import xml.etree.ElementTree as ET
import os
import sys
import comtypes.client

# 
def import_sections():
    
    # load xml file from SAP2000 installation folder
    tree = ET.parse(r'C:\Program Files\Computers and Structures\SAP2000 26\Property Libraries\Sections\CISC10.xml')
    root = tree.getroot()

    ns = {'csi': 'http://www.csiberkeley.com'}

    # find all steel pipe (round) 
    round_pipe = root.findall('.//csi:STEEL_PIPE', ns)

    # filter the labels
    hss_round = []
    for pipe in round_pipe:
        label = pipe.find('csi:LABEL', ns)
        if label is not None and label.text.startswith('HS'):
            hss_round.append(label.text)

    # find all HSS sections
    hss = root.findall('.//csi:STEEL_BOX', ns)

    hss_box = []
    for pipe in hss:
        label = pipe.find('csi:LABEL', ns)
        if label is not None and label.text.startswith('HS'):
            hss_box.append(label.text)
    
    return hss_round, hss_box

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
            diagonal_web_points.append((diagonal_web_points[j+1][0] + (i+1)*module_length, 0.0, diagonal_web_points[j+1][2]))
    return bottom_chord_points, top_chord_points, diagonal_web_points, vertical_web_points

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

if __name__ == '__main__':

    height = 3.0
    module_length = 15.0
    module_divisions = 5
    segment_length = module_length / module_divisions
    num_modules = 2

    bottom_chord_points, top_chord_points, diagonal_web_points, vertical_web_points = generate_warren(height, module_length, module_divisions, segment_length, num_modules)

    hss_round, hss_box = import_sections()

    API_path = 'D:\\Nick\\Documents\\School\\CIVE\\4A\\CIVE 400\\SAP Truss Script'
    file_name = 'BASE.sdb'

    path = API_path + os.sep + file_name

    #initialize_model(path)
    sap_model = refresh_model(path)

    # generate bottom chord
    bottom_chord_frames = []
    top_chord_frames = []
    diagonal_web_frames = []
    vertical_web_frames = []

    for i in range(len(bottom_chord_points)-1):
        bottom_chord_frames.append(sap_model.FrameObj.AddByCoord(*bottom_chord_points[i], *bottom_chord_points[i+1], 'foo', hss_round[0]))
    # generate top chord
    for i in range(len(top_chord_points)-1):
        top_chord_frames.append(sap_model.FrameObj.AddByCoord(*top_chord_points[i], *top_chord_points[i+1], 'foo', hss_round[0]))
    # generate diagonal webs
    for i in range(len(diagonal_web_points)-1):
        diagonal_web_frames.append(sap_model.FrameObj.AddByCoord(*diagonal_web_points[i], *diagonal_web_points[i+1], 'foo', hss_round[0]))
    # generate vertical webs
    for i in range(int(len(vertical_web_points)/2)):
        vertical_web_frames.append(sap_model.FrameObj.AddByCoord(*vertical_web_points[i*2], *vertical_web_points[i*2+1], 'foo', hss_round[0]))
    print(bottom_chord_frames)
    print(top_chord_frames)
    print(diagonal_web_frames)
    print(vertical_web_frames)


    # assign restraints
    # for our model we want a span of 30m, so 2 modules, assign a pin to left end and roller to right




