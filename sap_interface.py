import comtypes.client
import math

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
        
def sap_brace_bottom_chord(sap_model, bottom_chord_frames, num_spans, num_divisions):

    # loop over the bottom chord, and locate the frames to the left and right of the supports
    # get the location of their nodes, get the midpoint and add a point
    for i in range(num_spans):
        left_frame = bottom_chord_frames[i * num_divisions * 2]
        right_frame = bottom_chord_frames[(i * num_divisions * 2) + (num_divisions * 2 - 1)]
        
        frames = [left_frame, right_frame]
        points = []
        for frame in frames:
            # brace left at middle of frame by creating special joint
            point_1 = ''
            point_2 = ''
            point_1, point_2, ret = sap_model.FrameObj.GetPoints(frame)

            x_1, _, _, ret = sap_model.PointObj.GetCoordCartesian(point_1)
            x_2, _, _, ret = sap_model.PointObj.GetCoordCartesian(point_2)

            mid = (x_1 + x_2) / 2.

            points.append(sap_model.PointObj.AddCartesian(mid, 0, 0)[0])

        # set the restraint in the y-direction to the midpoints
        y_translation_restraint = [False, True, False, False, False, False]
        for point in points:
            ret = sap_model.PointObj.SetRestraint(point, y_translation_restraint)

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

def sap_barrier_load(sap_model, length, barrier_height, barrier_section, barrier_UDL):

    start = (0,0,barrier_height)
    end = (length,0,barrier_height)

    # barrier section is massless
    barrier = sap_model.FrameObj.AddByCoord(*start, *end, 'foo', barrier_section)[0]

    # add the barrier load pattern (not the case)
    ret = sap_model.LoadPatterns.Add('BARRIER_VERTICAL', 3, 0, True)
    ret = sap_model.LoadPatterns.Add('BARRIER_HORIZONTAL', 3, 0, True)

    ret = sap_model.FrameObj.SetLoadDistributed(barrier, 'BARRIER_VERTICAL', 1, 10, 0, 1, 
                                                barrier_UDL, barrier_UDL, RelDist = True)
    # horizontal load is dir y (5)
    ret = sap_model.FrameObj.SetLoadDistributed(barrier, 'BARRIER_HORIZONTAL', 1, 5, 0, 1, 
                                                barrier_UDL, barrier_UDL, RelDist = True)
    
    # divide frames at the intersections, to divide the barrier at the intersections w/ the webs
    # form rigid connection
    ret = sap_model.SelectObj.All()
    _, barriers, ret = sap_model.EditFrame.DivideAtIntersections(barrier, 0, [])
    return list(barriers)

def sap_set_loads(sap_model, bottom_chord_frames, top_chord_frames, dead_factor, live_factor, 
                  wearing_surface_factor, concrete_deck_factor, snow_factor, live_UDL,
                  wearing_surface_UDL, concrete_deck_UDL, snow_UDL, roof_UDL):
    # define the load patterns (ULS and SLS)
    # (case name, type (1=dead, 3=live, 8=other), self weight multiplier, add linear static load case)
    ret = sap_model.LoadPatterns.Add('DEAD', 1, 1, True)
    ret = sap_model.LoadPatterns.Add('LIVE', 3, 0, True)
    ret = sap_model.LoadPatterns.Add('DECK', 8, 0, True)
    ret = sap_model.LoadPatterns.Add('WEARING SURFACE', 8, 0, True)
    ret = sap_model.LoadPatterns.Add('SNOW', 8, 0, True)
    ret = sap_model.LoadPatterns.Add('ROOF', 8, 0, True)

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
    # snow load applies to top chord
    # roof load also applies to top
    for i in range(len(top_chord_frames)):
        ret = sap_model.FrameObj.SetLoadDistributed(top_chord_frames[i], 'SNOW', 1, 10, 0, 1, 
                                                    snow_UDL, snow_UDL, RelDist = True)
        ret = sap_model.FrameObj.SetLoadDistributed(top_chord_frames[i], 'ROOF', 1, 10, 0, 1,
                                                    roof_UDL, roof_UDL, RelDist = True)

    # for our governing ULS case, take 1.7 live and 1.5 snow
    ret = sap_model.LoadCases.StaticLinear.SetCase('ULS')
    ret = sap_model.LoadCases.StaticLinear.SetLoads(
        'ULS', 8, ['Load', 'Load', 'Load', 'Load', 'Load', 'Load', 'Load', 'Load'],
        ['DEAD', 'LIVE', 'BARRIER_VERTICAL', 'BARRIER_HORIZONTAL', 'DECK', 'WEARING SURFACE', 'ROOF', 'SNOW'], 
        [dead_factor, live_factor, live_factor, live_factor, 
         concrete_deck_factor, wearing_surface_factor, dead_factor, snow_factor])

    # for SLS its no snow, just 1.0 factors
    ret = sap_model.LoadCases.StaticLinear.SetCase('SLS')
    ret = sap_model.LoadCases.StaticLinear.SetLoads(
        'SLS', 7, ['Load', 'Load', 'Load', 'Load', 'Load', 'Load', 'Load'], 
        ['DEAD', 'LIVE', 'BARRIER_VERTICAL', 'BARRIER_HORIZONTAL', 'DECK', 'WEARING SURFACE', 'ROOF'], 
        [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0])

    # for modal analysis, increase the number of modes (eigen) to 40
    ret = sap_model.LoadCases.ModalEigen.SetNumberModes('MODAL', 40, 20)

def sap_gerber_modification(sap_model, vertical_web_frames, top_chord_frames, barrier_frames, 
                            num_spans, module_divisions):

    # track which frames that are deleted and remove from the list of frames
    web_deleted = []
    top_deleted = []

    # delete the vertical webs at the supports. These are every other entry to vertical_web_frames
    for index, frame in enumerate(vertical_web_frames):
        if index %  2 == 0:
            ret = sap_model.FrameObj.Delete(frame)
            web_deleted.append(frame)

    # delete the top chords to the left and right of the deleted vertical webs
    for i in range(num_spans):
        # if bottom chord is divided into module_divisions, top is divided into (module_divisions + 1)
        left_index = i * (module_divisions + 1) * 2
        right_index = left_index + (module_divisions + 1) * 2 - 1
        left_frame = top_chord_frames[left_index]
        right_frame = top_chord_frames[right_index]

        top_deleted.append(left_frame)
        top_deleted.append(right_frame)

        # deleting the vertical webs at the supports deletes the y-bracing. so need to rebrace at the 
        # corners of top chord of each gerber span
        # rebrace before deleting the top chord frames
        point_1 = ''
        point_2 = ''
        y_translation_restraint = [False, True, False, False, False, False]
        point_1, point_2, ret = sap_model.FrameObj.GetPoints(left_frame, point_1, point_2)
        # for the left frame we are about to delete, we want to brace the point_2 (right node)
        ret = sap_model.PointObj.SetRestraint(point_2, y_translation_restraint)
        point_1, point_2, ret = sap_model.FrameObj.GetPoints(right_frame, point_1, point_2)
        # for the right frame brace the point_1 (left node)
        ret = sap_model.PointObj.SetRestraint(point_1, y_translation_restraint)
        
        ret = sap_model.FrameObj.Delete(left_frame)
        ret = sap_model.FrameObj.Delete(right_frame)

    # delete the outermost barrier frame
    ret = sap_model.FrameObj.Delete(barrier_frames[0])
    ret = sap_model.FrameObj.Delete(barrier_frames[-1])

    # remove the deleted frames from their respective lists
    for web in web_deleted:
        vertical_web_frames.remove(web)
    for top in top_deleted:
        top_chord_frames.remove(top)
    del barrier_frames[0]
    del barrier_frames[-1]

    # return the modified lists of frames
    return vertical_web_frames, top_chord_frames, barrier_frames

def sap_run_analysis(sap_model, file_path):
    ret = sap_model.File.Save(file_path)
    ret = sap_model.Analyze.RunAnalysis()

def sap_deflection(sap_model, bottom_chord_frames, span_length):

    deflection_limit = span_length / 360.

    ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    ret = sap_model.Results.Setup.SetCaseSelectedForOutput('SLS')

    # get the central node vertical displacement
    center_node = sap_central_node(sap_model, bottom_chord_frames)    
    _, _, _, _, _, _, _, _, temp, _, _, _, ret = sap_model.Results.JointDispl(center_node,
                                        0, 0, [], [], [], [], [], [], [], [], [], [], []) 
    # return absolute value
    deflection = abs(temp[0])
    percentage = deflection / deflection_limit * 100

    return deflection, percentage

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

def sap_vibration_analysis(sap_model, pedestrian_density, concrete_deck_UDL, live_UDL, damping_ratio):
    # get results from 'MODAL' load case
    ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    ret = sap_model.Results.Setup.SetCaseSelectedForOutput('MODAL')
    Uz = []
    period = []
    # get the list of modal participating mass ratios for each mode. we want to find the mode that has highest 
    # participating ratio in the Z (vertical) direction
    _, _, _, _, period, _, _, Uz, _, _, _, _, _, _, _, _, _, ret = sap_model.Results.ModalParticipatingMassRatios(
        0, [], [], [], period, [], [], [], [], [], [], [], [], [], [], [], []
    )
    # get the mode that has the highest Uz
    max_Uz = max(Uz)
    mode_index = Uz.index(max_Uz)
    natural_period = period[mode_index]
    natural_frequency = 1 / natural_period

    if natural_frequency < 13.:
        in_crit_range = True
    else:
        in_crit_range = False

    # calculate rho_mass = sqrt((mu_D + mu_P)/mu_D) > 1.025 
    # where mu_D is deck mass per unit span
    # and mu_P is pedestrian mass per unit span
    mu_D = concrete_deck_UDL * 101.992 # kg/kN conversion
    mu_p = live_UDL * 101.992 
    rho_mass = math.sqrt((mu_D + mu_p) / mu_D)
    natural_frequency_occupied = natural_frequency / rho_mass

    # calculate forcing frequency of pedestrians (Hz)
    forcing_frequency = 0.099 * pedestrian_density**2 - 0.644*pedestrian_density + 2.188

    # calculate resonating harmonic
    m_empty = natural_frequency / forcing_frequency
    m_occupied = natural_frequency_occupied / forcing_frequency

    # if either of the resonating harmonics are below 5, further analysis is required
    """ if m_empty < 5. or m_occupied < 5.:
        m_crit = min(m_empty, m_occupied)
        m = round(m_crit)
        frequency = min(natural_frequency, natural_frequency_occupied)
        period = 1. / frequency

        alphas = [0.4, 0.1, 0.042, 0.041, 0.027]
        alpha = alphas[m-1]
        pedestrian_weight = 700. # N

        # dummy variables because idk what these are
        N = 1
        S = 1
        if pedestrian_density < 1.:
            n_eq = 10.8 * math.sqrt(damping_ratio * N) / S
        else:
            n_eq = 1.85 * math.sqrt(N) / S

        amplitude = n_eq * pedestrian_weight * alpha

        # create a cosine time history function
        ret = sap_model.Func.FuncTH.SetCosine('HARMONIC LOAD', period, 20, 10, amplitude)
        
        # create load pattern 
        ret = sap_model.LoadPatterns.Add('VIBRATION', 8, 0, True) """

    return natural_frequency, in_crit_range, natural_frequency_occupied, m_empty, m_occupied

def sap_steel_design(sap_model):

    ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    ret = sap_model.Results.Setup.SetCaseSelectedForOutput('ULS')

    ret = sap_model.DesignSteel.StartDesign()
    num_failed = 0
    names = []
    _, num_failed, _, names, ret = sap_model.DesignSteel.VerifyPassed(0, num_failed, 0, names)

    if num_failed != 0: 
        passed = False
    else:
        passed = True

    failed_sections = []
    for name in names:
        section, _, ret = sap_model.FrameObj.GetSection(name, '', '')
        failed_sections.append(section)
    unique_sections = list(set(failed_sections))

    if len(unique_sections) == 0:
        failed_section_names = 'None'
    else:
        failed_section_names = ', '.join(unique_sections)
   
    return passed, failed_section_names

