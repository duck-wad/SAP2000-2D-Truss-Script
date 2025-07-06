import os
import pandas as pd

from sap_interface import *
from define_geometry import *

if __name__ == '__main__':

    ''' ------------------------ DEFINE MODEL PARAMETERS ------------------------ '''

    height = 3.0
    module_length = 15.0
    module_divisions = 4 # applies to the bottom chord
    segment_length = module_length / module_divisions
    # each span contains 2 modules, so there will be num_spans x 2 modules
    # num_spans should be an odd number
    num_spans = 5
    assert num_spans % 2 != 0
    num_modules = 2*num_spans

    ''' ------------------------ DEFINE LOADS ------------------------ '''

    deck_width = 3.0 # meters
    deck_thickness = 0.15 # 150mm concrete deck
    concrete_density = 24 # kN/m3
    asphalt_thickness = 0.004 # 4mm
    asphalt_density = 21 # kN/m3
    pedestrian_pressure = 4.25 # kPa
    barrier_load = 1.2 # kN/m

    # load factors
    dead_factor = 1.1
    live_factor = 1.7
    wearing_surface_factor = 1.5
    concrete_deck_factor = 1.2

    # compute UDLs
    live_UDL = barrier_load + deck_width / 2 * pedestrian_pressure
    wearing_surface_UDL = asphalt_density * asphalt_thickness * deck_width / 2
    concrete_deck_UDL = concrete_density * deck_thickness * deck_width / 2

    ''' ------------------------ IMPORT SECTIONS ------------------------ '''

    # import sections from excel 
    df = pd.read_excel('./sections.xlsx')
    hss_round = df['HSS Round'].dropna().tolist() if 'HSS Round' in df.columns else []
    hss_box = df['HSS Box'].dropna().tolist() if 'HSS Box' in df.columns else []

    ''' ------------------------ GENERATE GEOMETRY ------------------------ '''

    bottom_chord_points, top_chord_points, diagonal_web_points, vertical_web_points = generate_warren(
        height, module_length, module_divisions, segment_length, num_modules)

    ''' ------------------------ INITIALIZE MODEL ------------------------ '''

    root_path = os.getcwd()
    base_file_path = root_path + '/BASE.sdb'
    os.makedirs('./models', exist_ok=True)
    model_path = root_path + '/models/MODEL.sdb'

    # initialize fresh model from BASE in root folder
    sap_model = sap_initialize_model(base_file_path)

    # initialize the list of results for each section combination
    results = []

    # TEMPORARY SET SECTIONS
    bottom_chord_section = hss_round[10]
    top_chord_section = hss_round[10]
    web_section = hss_round[10]

    ''' ------------------------ CREATE SAP MODEL ------------------------ '''

    # generate frames
    bottom_chord_frames, top_chord_frames, diagonal_web_frames, vertical_web_frames = sap_create_frame(
        sap_model, bottom_chord_points, top_chord_points, diagonal_web_points, vertical_web_points,
        bottom_chord_section, top_chord_section, web_section)

    # set the restraints
    sap_set_restraints(sap_model, vertical_web_frames, num_spans)

    # set the releases for moment splice between modules
    sap_set_releases(sap_model, vertical_web_frames, bottom_chord_frames, top_chord_frames, 
                     diagonal_web_frames, num_modules, module_divisions)

    # set the load case, apply deck load to bottom chord
    sap_set_loads(sap_model, bottom_chord_frames, dead_factor, live_factor, 
                  wearing_surface_factor, concrete_deck_factor, live_UDL, 
                  wearing_surface_UDL, concrete_deck_UDL)
    
    ''' ------------------------ RUN MODEL AND COLLECT RESULTS ------------------------ '''
    
    # save the file to a new file in the models folder (so don't override the BASE file)
    # return the vertical displacement of the central node
    vert_disp = sap_run_analysis(sap_model, model_path, bottom_chord_frames)
    print(vert_disp)
    # get the reaction output from dead case and divide by num_modules
    module_mass = sap_module_mass(sap_model, num_modules)
    print(module_mass)

    results.append({
        'Bottom chord': bottom_chord_section,
        'Top chord': top_chord_section,
        'Web members': web_section,
        'Max vertical displacement (m)': vert_disp,
        'Module mass (kg)': module_mass
    })

    ''' ------------------------ EXPORT RESULTS ------------------------ '''
    df = pd.DataFrame(results)
    df.to_excel('output.xlsx', index=False)