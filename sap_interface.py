import comtypes.client
import math


def sap_initialize_model(base_file_path):
    # create API helper object
    helper = comtypes.client.CreateObject('SAP2000v1.Helper')
    helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)

    sap_object = helper.GetObject("CSI.SAP2000.API.SapObject")
    if sap_object is None:
        sap_object = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
        sap_object.ApplicationStart() 

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

def sap_set_loads(sap_model, bottom_chord_frames, dead_factor, live_factor, 
                  wearing_surface_factor, concrete_deck_factor, live_UDL, 
                  wearing_surface_UDL, concrete_deck_UDL):
    # define the load patterns (dead and live)
    # (case name, type (1=dead, 8=other), self weight multiplier, add linear static load case)
    ret = sap_model.LoadPatterns.Add('DEAD', 1, 1, True)
    ret = sap_model.LoadPatterns.Add('LIVE', 8, 0, True)
    ret = sap_model.LoadPatterns.Add('DECK', 8, 0, True)
    ret = sap_model.LoadPatterns.Add('WEARING SURFACE', 8, 0, True)

    for i in range(len(bottom_chord_frames)):
        # apply live, deck, and asphalt as UDL to bottom chord
        # (name, load case, type (1 is force per unit length, 2 is moment per unit length), 
        # integer indicating direction (10 is gravity dir), dist1, dist2, val1, val2)
        ret = sap_model.FrameObj.SetLoadDistributed(bottom_chord_frames[i], 'LIVE', 1, 10, 0, 1, 
                                                    live_UDL, live_UDL, RelDist = True)
        ret = sap_model.FrameObj.SetLoadDistributed(bottom_chord_frames[i], 'DECK', 1, 10, 0, 1, 
                                                    concrete_deck_UDL, concrete_deck_UDL, RelDist = True)
        ret = sap_model.FrameObj.SetLoadDistributed(bottom_chord_frames[i], 'WEARING SURFACE', 1, 10, 0, 1, 
                                                    wearing_surface_UDL, wearing_surface_UDL, RelDist = True)
    ret = sap_model.LoadCases.StaticLinear.SetCase('FACTORED')
    ret = sap_model.LoadCases.StaticLinear.SetLoads(
        'FACTORED', 4, ['Load', 'Load', 'Load', 'Load'], ['DEAD', 'LIVE', 'DECK', 'WEARING SURFACE'], 
        [dead_factor, live_factor, wearing_surface_factor, concrete_deck_factor])
    
    # also need to add a load case to calculate the natural frequency of the span
    # apply a 1kN load to the central node, and then measure the displacement 
    # then calculate the natural frequency as 1/2pi*sqrt(P/delta*mL) where mL is mass per unit length
    center_node = sap_central_node(sap_model, bottom_chord_frames)
    ret = sap_model.LoadPatterns.Add('POINT UNIT LOAD', 8, 0, True)
    point_load = [0,0,-1,0,0,0]
    ret = sap_model.PointObj.SetLoadForce(center_node, 'POINT UNIT LOAD', point_load)
    ret = sap_model.LoadCases.StaticLinear.SetCase('POINT UNIT LOAD')
    ret = sap_model.LoadCases.StaticLinear.SetLoads('POINT UNIT LOAD', 1, ['Load'], ['POINT UNIT LOAD'], [1])  
    
def sap_run_analysis(sap_model, file_path):
    ret = sap_model.File.Save(file_path)
    ret = sap_model.Analyze.RunAnalysis()

def sap_factored_displacement(sap_model, bottom_chord_frames):
    # get the results from the 'FACTORED' load case
    ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    ret = sap_model.Results.Setup.SetCaseSelectedForOutput('FACTORED')

    # get the central node vertical displacement under the 'FACTORED' load case
    center_node = sap_central_node(sap_model, bottom_chord_frames)    
    _, _, _, _, _, _, _, _, vert_disp, _, _, _, ret = sap_model.Results.JointDispl(center_node,
                                        0, 0, [], [], [], [], [], [], [], [], [], [], []) 
    # return the abs value
    return abs(vert_disp[0])

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

def sap_natural_frequency(sap_model, module_mass, bottom_chord_frames):
    # get the results from the 'POINT UNIT LOAD' load case
    ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    ret = sap_model.Results.Setup.SetCaseSelectedForOutput('POINT UNIT LOAD')

    # calculate the natural frequency
    # get the displacement of center node under the 'POINT UNIT LOAD' case
    # calculate fn as 1/2pi*sqrt(P/(delta*mL)) where mL is mass per unit length
    # mass per unit length would be 2*module_mass / 
    center_node = sap_central_node(sap_model, bottom_chord_frames)   
    _, _, _, _, _, _, _, _, disp, _, _, _, ret = sap_model.Results.JointDispl(center_node,
                                        0, 0, [], [], [], [], [], [], [], [], [], [], []) 
    disp = disp[0]*-1
    
    # get the module mass per unit length
    span_mass = module_mass * 2

    f_n = (1 / (2 * math.pi)) * math.sqrt(1000 / (disp * span_mass))
    return f_n
