import comtypes.client

def sap_open():
    # create API helper object
    helper = comtypes.client.CreateObject('SAP2000v1.Helper')
    helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)
    sap_object = helper.GetObject("CSI.SAP2000.API.SapObject")
    if sap_object is None:
        sap_object = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
        sap_object.ApplicationStart() 

    return sap_object

def sap_close(sap_object):
    ret = sap_object.ApplicationExit(False)
    sap_object = None

def sap_initialize_model(base_file_path, sap_object):

    sap_model = sap_object.SapModel
    ret = sap_model.File.OpenFile(base_file_path)
    return sap_model

def sap_create_frame(sap_model, bottom_chord_points, top_chord_points, diagonal_web_points, vertical_web_points,
                     bottom_chord_section, top_chord_section, web_section):
    
    # generate bottom chord
    bottom_chord_frames = []
    top_chord_frames = []
    diagonal_web_frames = []
    vertical_web_frames = []

    for i in range(len(bottom_chord_points)-1):
        bottom_chord_frames.append(sap_model.FrameObj.AddByCoord(*bottom_chord_points[i], 
                                                                 *bottom_chord_points[i+1], 'foo', bottom_chord_section)[0])
    # generate top chord
    for i in range(len(top_chord_points)-1):
        top_chord_frames.append(sap_model.FrameObj.AddByCoord(*top_chord_points[i], 
                                                              *top_chord_points[i+1], 'foo', top_chord_section)[0])
    # generate diagonal webs
    for i in range(len(diagonal_web_points)-1):
        diagonal_web_frames.append(sap_model.FrameObj.AddByCoord(*diagonal_web_points[i], 
                                                                 *diagonal_web_points[i+1], 'foo', web_section)[0])
    # generate vertical webs
    for i in range(int(len(vertical_web_points)/2)):
        vertical_web_frames.append(sap_model.FrameObj.AddByCoord(*vertical_web_points[i*2], 
                                                                 *vertical_web_points[i*2+1], 'foo', web_section)[0])

    return bottom_chord_frames, top_chord_frames, diagonal_web_frames, vertical_web_frames

def sap_set_restraints(sap_model, vertical_web_frames, num_spans):

    point_1 = ''
    point_2 = ''

    # for 2D truss, need to restrain the corners of every module in the y (out of plane) direction
    # this allows the truss to sway in the y direction but not completely shift out of plane
    y_translation_restraint = [False, True, False, False, False, False]
    # these points are all points of the vertical members
    for i in range(len(vertical_web_frames)):
        point_1, point_2, ret = sap_model.FrameObj.GetPoints(vertical_web_frames[i], point_1, point_2)
        ret = sap_model.PointObj.SetRestraint(point_1, y_translation_restraint)
        ret = sap_model.PointObj.SetRestraint(point_2, y_translation_restraint)

    # assign pin to base of every span end (every 2 modules)
    # left pin corresponds with the first node of the first entry to vertical web list
    # this overrides the y_translation_restraint 
    pin_restraint = [True, True, True, False, False, False]

    for i in range(num_spans+1):
        point_1, point_2, ret = sap_model.FrameObj.GetPoints(vertical_web_frames[i*2], point_1, point_2)
        ret = sap_model.PointObj.SetRestraint(point_1, pin_restraint)

    # should also pin the corners of the edge modules
    point_1, point_2, ret = sap_model.FrameObj.GetPoints(vertical_web_frames[0], point_1, point_2)
    ret = sap_model.PointObj.SetRestraint(point_2, pin_restraint)
    point_1, point_2, ret = sap_model.FrameObj.GetPoints(vertical_web_frames[-1], point_1, point_2)
    ret = sap_model.PointObj.SetRestraint(point_2, pin_restraint)

def sap_set_releases(sap_model, vertical_web_frames, bottom_chord_frames, top_chord_frames, 
                     diagonal_web_frames, num_modules, module_divisions):

    # set the release for the moment splice between modules. Release M3
    # applies to the vertical, diagonal, and the chord members to the LEFT of the splice
    # need to not have releases at the chord to the right 
    moment_release = [False, False, False, False, False, True]
    no_release = [False, False, False, False, False, False]
    startval = [0,0,0,0,0,0]
    endval = [0,0,0,0,0,0]
    # loop through verticals and release only the interior webs
    for i in range(len(vertical_web_frames)):
        if i == 0 or i == len(vertical_web_frames)-1:
            continue
        ret = sap_model.FrameObj.SetReleases(vertical_web_frames[i], moment_release, moment_release, startval, endval)

    # find the chords to the left of the moment splice
    # for the bottom chord there are module_divisions number of segments, so need to get the segment at index module_divisions-1 
    # for top chord, there will be module_divisions+1 number of segments, so get the segment at module_division
    # apply only to interior splices
    for i in range(num_modules-1):
        ret = sap_model.FrameObj.SetReleases(bottom_chord_frames[module_divisions-1+i*module_divisions], 
                                             no_release, moment_release, startval, endval)
        ret = sap_model.FrameObj.SetReleases(top_chord_frames[module_divisions + i*(module_divisions+1)], 
                                             no_release, moment_release, startval, endval)
        
    # find the diagonals to the left of the moment splice
    # for each module, there are module_divisions*2 diagonals, need to release the last one
    for i in range(num_modules-1):
        
        ret = sap_model.FrameObj.SetReleases(diagonal_web_frames[module_divisions*2-1+i*2*module_divisions], no_release, 
                                             moment_release, startval, endval)
        
def sap_central_node(sap_model, bottom_chord_frames):
    # get the displacement of node in center of of the middlemost span (need num_spans to be odd)
    # bottom_chord_frames has even number of frames. the frame with its right node at the point we want
    # is at the halfway point in the array
    index = int(len(bottom_chord_frames) / 2)
    target = bottom_chord_frames[index]
    point_1 = ''
    point_2 = ''
    point_1, point_2, ret = sap_model.FrameObj.GetPoints(target, point_1, point_2)
    return point_1

def sap_set_loads(sap_model, bottom_chord_frames, top_chord_frames, dead_factor, live_factor, 
                  wearing_surface_factor, concrete_deck_factor, snow_factor, live_UDL_vertical, 
                  live_UDL_horizontal, wearing_surface_UDL, concrete_deck_UDL, snow_UDL, roof_UDL):
    # define the load patterns (dead and live)
    # (case name, type (1=dead, 8=other), self weight multiplier, add linear static load case)
    ret = sap_model.LoadPatterns.Add('DEAD', 1, 1, True)
    ret = sap_model.LoadPatterns.Add('LIVE_VERTICAL', 8, 0, True)
    ret = sap_model.LoadPatterns.Add('LIVE_HORIZONTAL', 8, 0, True)
    ret = sap_model.LoadPatterns.Add('DECK', 8, 0, True)
    ret = sap_model.LoadPatterns.Add('WEARING SURFACE', 8, 0, True)
    ret = sap_model.LoadPatterns.Add('SNOW', 8, 0, True)
    ret = sap_model.LoadPatterns.Add('ROOF', 8, 0, True)

    for i in range(len(bottom_chord_frames)):
        # apply live, deck, and asphalt as UDL to bottom chord
        # (name, load case, type (1 is force per unit length, 2 is moment per unit length), 
        # integer indicating direction (10 is gravity dir), dist1, dist2, val1, val2)
        ret = sap_model.FrameObj.SetLoadDistributed(bottom_chord_frames[i], 'LIVE_VERTICAL', 1, 10, 0, 1, 
                                                    live_UDL_vertical, live_UDL_vertical, RelDist = True)
        # horizontal load is dir y (5)
        ret = sap_model.FrameObj.SetLoadDistributed(bottom_chord_frames[i], 'LIVE_HORIZONTAL', 1, 5, 0, 1, 
                                                    live_UDL_horizontal, live_UDL_horizontal, RelDist = True)
        ret = sap_model.FrameObj.SetLoadDistributed(bottom_chord_frames[i], 'DECK', 1, 10, 0, 1, 
                                                    concrete_deck_UDL, concrete_deck_UDL, RelDist = True)
        ret = sap_model.FrameObj.SetLoadDistributed(bottom_chord_frames[i], 'WEARING SURFACE', 1, 10, 0, 1, 
                                                    wearing_surface_UDL, wearing_surface_UDL, RelDist = True)
    # snow load applies to top chord
    # roof load also applies to top
    for i in range(len(top_chord_frames)):
        ret = sap_model.FrameObj.SetLoadDistributed(top_chord_frames[i], 'SNOW', 1, 10, 0, 1, 
                                                    snow_UDL, snow_UDL, RelDist = True)
        ret = sap_model.FrameObj.SetLoadDistributed(top_chord_frames[i], 'ROOF', 1, 10, 0, 1,
                                                    roof_UDL, roof_UDL, RelDist = True)

    ret = sap_model.LoadCases.StaticLinear.SetCase('ULS_1')
    ret = sap_model.LoadCases.StaticLinear.SetLoads(
        'ULS_1', 6, ['Load', 'Load', 'Load', 'Load', 'Load', 'Load'], ['DEAD', 'LIVE_VERTICAL', 
                                                               'LIVE_HORIZONTAL', 'DECK', 'WEARING SURFACE', 'ROOF'], 
        [dead_factor, live_factor, live_factor, wearing_surface_factor, concrete_deck_factor, dead_factor])
    
    ret = sap_model.LoadCases.StaticLinear.SetCase('ULS_5')
    ret = sap_model.LoadCases.StaticLinear.SetLoads(
        'ULS_5', 5, ['Load', 'Load', 'Load', 'Load', 'Load'], ['DEAD', 'SNOW', 'DECK', 'WEARING SURFACE', 'ROOF'], 
        [dead_factor, snow_factor, wearing_surface_factor, concrete_deck_factor, dead_factor])

    # for modal analysis, increase the number of modes (eigen) to 40
    ret = sap_model.LoadCases.ModalEigen.SetNumberModes('MODAL', 40, 20)

    # return list of load cases for later reference
    return ['ULS_1', 'ULS_5']
    
def sap_run_analysis(sap_model, file_path):
    ret = sap_model.File.Save(file_path)
    ret = sap_model.Analyze.RunAnalysis()

def sap_displacement(sap_model, load_cases, bottom_chord_frames):

    displacements = []

    for case in load_cases:
        # get the results from the 'ULS_1' load case
        ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
        ret = sap_model.Results.Setup.SetCaseSelectedForOutput(case)

        # get the central node vertical displacement
        center_node = sap_central_node(sap_model, bottom_chord_frames)    
        _, _, _, _, _, _, _, _, temp, _, _, _, ret = sap_model.Results.JointDispl(center_node,
                                            0, 0, [], [], [], [], [], [], [], [], [], [], []) 
        # return absolute value
        displacements.append(abs(temp[0]))

    return displacements

def sap_module_mass(sap_model, num_modules):
    
    # get the results from the 'DEAD' load case
    ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    ret = sap_model.Results.Setup.SetCaseSelectedForOutput('DEAD')

    _, _, _, _, _, _, reaction, _, _, _, _, _, _, ret = sap_model.Results.BaseReact(
        0, [], [], [], [], [], [], [], [], [], 0, 0, 0)
    
    # convert kN to kg
    total_mass = reaction[0] / 9.81 * 1000
    module_mass = total_mass / num_modules
    
    return module_mass

def sap_natural_frequency(sap_model, pedestrian_density):
    # get results from 'MODAL' load case
    ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    ret = sap_model.Results.Setup.SetCaseSelectedForOutput('MODAL')
    Uz = []
    sumUz = []
    period = []
    # get the list of modal participating mass ratios for each mode. we want to find the mode that has highest 
    # participating ratio in the Z (vertical) direction
    _, _, _, _, period, _, _, Uz, _, _, sumUz, _, _, _, _, _, _, ret = sap_model.Results.ModalParticipatingMassRatios(
        0, [], [], [], period, [], [], [], [], [], [], [], [], [], [], [], []
    )
    # get the mode that has the highest Uz
    max_Uz = max(Uz)
    mode_index = Uz.index(max_Uz)
    natural_period = period[mode_index]
    natural_frequency = 1 / natural_period
    
    # calculate forcing frequency of pedestrians (Hz)
    forcing_frequency = 0.099 * pedestrian_density**2 - 0.644*pedestrian_density + 2.188

    # calculate resonating harmonic
    m = natural_frequency / forcing_frequency

    return natural_frequency, m

def sap_steel_design(sap_model, load_cases):

    passed = []

    for case in load_cases:
        ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
        ret = sap_model.Results.Setup.SetCaseSelectedForOutput(case)

        ret = sap_model.DesignSteel.StartDesign()
        num_failed = 0
        _, num_failed, _, _, ret = sap_model.DesignSteel.VerifyPassed(0, num_failed, 0, [])

        if num_failed != 0: 
            passed.append(False)
        else:
            passed.append(True)
        
    return passed

