import os
import pandas as pd
from tqdm import tqdm

from sap_interface import *
from define_geometry import *
from sections import create_section_combinations

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
    span_length = 2*module_length

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

    ''' ------------------------ INITIALIZE MODEL ------------------------ '''
    
    # generate geometry
    bottom_chord_points, top_chord_points, diagonal_web_points, vertical_web_points = generate_warren(
        height, module_length, module_divisions, segment_length, num_modules)
    
    # import the section combinations
    section_combinations = create_section_combinations()
    
    # set file paths
    root_path = os.getcwd()
    base_file_path = root_path + '/BASE.sdb'
    os.makedirs('./models', exist_ok=True)
    model_path = root_path + '/models/MODEL.sdb'

    # initialize the list of results for each section combination
    results = []

    for i in tqdm(range(len(section_combinations))):

        top_chord_section = section_combinations[i][0]
        bottom_chord_section = section_combinations[i][1]
        web_section = section_combinations[i][2]

        # initialize fresh model from BASE in root folder
        sap_model = sap_initialize_model(base_file_path)

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
        sap_run_analysis(sap_model, model_path)

        vert_disp = sap_factored_displacement(sap_model, bottom_chord_frames)
        # get the reaction output from dead case and divide by num_modules
        module_mass = sap_module_mass(sap_model, num_modules)

        natural_frequency = sap_natural_frequency(sap_model, module_mass, bottom_chord_frames)

        results.append({
            'Bottom chord': bottom_chord_section,
            'Top chord': top_chord_section,
            'Web members': web_section,
            'Max vertical displacement (m)': vert_disp,
            'Module mass (kg)': module_mass
        })

        # log results 
        tqdm.write(f'Bottom chord section: {bottom_chord_section}, Top chord section: {top_chord_section}, Web member section: {web_section}')
        tqdm.write(f'Displacement of central node (mm): {vert_disp * 1000}')
        tqdm.write(f'Mass of single module (kg): {module_mass}')
        tqdm.write(f'Natural frequency of span (Hz): {natural_frequency}')

    ''' ------------------------ EXPORT RESULTS ------------------------ '''
    df = pd.DataFrame(results)
    df.to_excel('output.xlsx', index=False)