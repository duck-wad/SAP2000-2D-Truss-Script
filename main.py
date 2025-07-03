import os
import sys
import comtypes.client
import pandas as pd

''' ------------------------------------------------ GENERATE TRUSS GEOMETRY ------------------------------------------------ '''

def generate_warren(height, module_length, module_divisions, segment_length, num_modules):

    bottom_chord_points = []
    top_chord_points = []
    diagonal_web_points = []
    vertical_web_points = []

    # generate bottom chord points for first module
    for i in range(module_divisions + 1):
        bottom_chord_points.append((i*segment_length, 0.0, 0.0))
    # copy the chord over for the other modules
    # exclude the first point since it's already accounted for
    temp = len(bottom_chord_points) - 1
    for i in range(num_modules - 1):
        for j in range(temp):
            bottom_chord_points.append((bottom_chord_points[j+1][0] + (i+1)*module_length, 0.0, 0.0))

    # generate top chord points for first module
    for i in range(module_divisions + 2):
        if i == 0:
            top_chord_points.append((0.0, 0.0, height))
        elif i == module_divisions + 1:
            top_chord_points.append((module_length, 0.0, height))
        else:
            top_chord_points.append((i*segment_length - 0.5*segment_length, 0.0, height))
    # copy chord over for the other modules
    temp = len(top_chord_points) - 1
    for i in range(num_modules - 1):
        for j in range(temp):
            top_chord_points.append((top_chord_points[j+1][0] + (i+1)*module_length, 0.0, height))
    
    # generate the vertical web points which occur at x locations of multiples of module_length
    # for num_modules, there will be num_modules+1 vertical webs
    # in order (bottom_1, top_1, bottom_2, top_2, ....)
    for i in range(num_modules + 1):
        vertical_web_points.append((i*module_length, 0.0, 0.0))
        vertical_web_points.append((i*module_length, 0.0, height))

    # generate diagonal web points for first module
    bottom_counter = 0
    top_counter = 1
    for i in range(2 * module_divisions + 1):
        if i % 2 == 0:
            diagonal_web_points.append(bottom_chord_points[bottom_counter])
            bottom_counter += 1
        else:
            diagonal_web_points.append(top_chord_points[top_counter])
            top_counter += 1
    # copy webs over for the other modules
    temp = len(diagonal_web_points) - 1
    for i in range(num_modules - 1):
        for j in range(temp):
            diagonal_web_points.append((diagonal_web_points[j+1][0] + (i+1)*module_length, 0.0, 
                                        diagonal_web_points[j+1][2]))
    return bottom_chord_points, top_chord_points, diagonal_web_points, vertical_web_points

''' ------------------------------------------------ SAP INTERFACE FUNCTIONS ------------------------------------------------ '''

def initialize_model(model_path):

    # create API helper object
    helper = comtypes.client.CreateObject('SAP2000v1.Helper')
    helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)

    sap_object = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
    sap_object.ApplicationStart()
    sap_model = sap_object.SapModel
    ret = sap_model.File.OpenFile(model_path)
    return sap_model
    
def refresh_model(model_path):
    
    helper = comtypes.client.CreateObject('SAP2000v1.Helper')
    helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)

    sap_object = helper.GetObject("CSI.SAP2000.API.SapObject")
    sap_model = sap_object.SapModel
    ret = sap_model.File.OpenFile(model_path)
    return sap_model

def sap_create_frame(sap_model, bottom_chord_points, top_chord_points, diagonal_web_points, vertical_web_points,
                     bottom_chord_section, top_chord_section, diagonal_web_section, vertical_web_section):
    
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
                                                                 *diagonal_web_points[i+1], 'foo', diagonal_web_section)[0])
    # generate vertical webs
    for i in range(int(len(vertical_web_points)/2)):
        vertical_web_frames.append(sap_model.FrameObj.AddByCoord(*vertical_web_points[i*2], 
                                                                 *vertical_web_points[i*2+1], 'foo', vertical_web_section)[0])

    return bottom_chord_frames, top_chord_frames, diagonal_web_frames, vertical_web_frames

def sap_set_restraints(sap_model, vertical_web_frames):

    # assign restraints, pin to left end and roller to right
    # left pin corresponds with the first node of the first entry to vertical web list
    # right roller corresponds with the first node of the last entry to vertical web list
    pin_restraint = [True, True, True, False, False, False]
    roller_restraint = [False, False, True, False, False, False]
    # set the left restraint
    point_1 = ''
    point_2 = ''
    point_1, point_2, ret = sap_model.FrameObj.GetPoints(vertical_web_frames[0], point_1, point_2)
    ret = sap_model.PointObj.SetRestraint(point_1, pin_restraint)
    # set the right restraint
    point_1, point_2, ret = sap_model.FrameObj.GetPoints(vertical_web_frames[-1], point_1, point_2)
    ret = sap_model.PointObj.SetRestraint(point_1, roller_restraint)

def sap_set_releases(sap_model, vertical_web_frames, bottom_chord_frames, top_chord_frames):

    # set the release for the moment splice between modules. Release M3
    # applies to the vertical and the chord members to the LEFT of the splice
    # need to not have releases at the chord to the right (releasing would be redundant)
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

def sap_set_loads(sap_model, bottom_chord_frames, live_load):
    # define the load patterns (dead and live)
    # (case name, type (1=dead, 8=other), self weight multiplier, add linear static load case)
    ret = sap_model.LoadPatterns.Add('DEAD', 1, 1, True)
    ret = sap_model.LoadPatterns.Add('LIVE', 8, 0, True)

    for i in range(len(bottom_chord_frames)):
        # (name, load case, type (1 is force per unit length, 2 is moment per unit length), 
        # integer indicating direction (10 is gravity dir), dist1, dist2, val1, val2)
        ret = sap_model.FrameObj.SetLoadDistributed(bottom_chord_frames[i], 'LIVE', 1, 10, 0, 1, 
                                                    live_load, live_load, RelDist = True)
    ret = sap_model.LoadCases.StaticLinear.SetCase('FACTORED')
    ret = sap_model.LoadCases.StaticLinear.SetLoads('FACTORED', 2, ['Load', 'Load'], ['DEAD', 'LIVE'], [1.25, 1.5])

''' ------------------------------------------------ MAIN WORKFLOW ------------------------------------------------ '''

if __name__ == '__main__':

    height = 3.0
    module_length = 15.0
    module_divisions = 5 # applies to the bottom chord
    segment_length = module_length / module_divisions
    num_modules = 2

    live_load = 20.0 # kN/m applied to the bottom chord (deck load)

    API_path = 'C:\\Users\\Nick\\source\\repos\\Capstone\\SAP Truss Script'
    file_name = 'BASE.sdb'
    path = API_path + os.sep + file_name

    # import sections from excel 
    df = pd.read_excel('./HSS_Sections.xlsx')
    hss_round = df['HSS Round'].dropna().tolist() if 'HSS Round' in df.columns else []
    hss_box = df['HSS Box'].dropna().tolist() if 'HSS Box' in df.columns else []

    #initialize_model(path)
    sap_model = refresh_model(path)

    bottom_chord_points, top_chord_points, diagonal_web_points, vertical_web_points = generate_warren(
        height, module_length, module_divisions, segment_length, num_modules)

    # TEMPORARY SET SECTIONS
    bottom_chord_section = hss_round[10]
    top_chord_section = hss_round[10]
    diagonal_web_section = hss_round[10]
    vertical_web_section = hss_round[10]

    # generate frames
    bottom_chord_frames, top_chord_frames, diagonal_web_frames, vertical_web_frames = sap_create_frame(
        sap_model, bottom_chord_points, top_chord_points, diagonal_web_points, vertical_web_points,
        bottom_chord_section, top_chord_section, diagonal_web_section, vertical_web_section)

    # set the restraints
    sap_set_restraints(sap_model, vertical_web_frames)

    # set the releases for moment splice between modules
    sap_set_releases(sap_model, vertical_web_frames, bottom_chord_frames, top_chord_frames)

    # set the load case, apply deck load to bottom chord
    sap_set_loads(sap_model, bottom_chord_frames, live_load)

    
