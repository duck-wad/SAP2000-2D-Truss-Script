import os
from tqdm import tqdm

from sap_interface import *
from define_geometry import *
from define_sections import *

if __name__ == '__main__':

    ''' ------------------------ DEFINE MODEL PARAMETERS ------------------------ '''

    height = 3.0
    module_length = 15.0
    module_divisions = 3 # applies to the bottom chord
    segment_length = module_length / module_divisions
    # each span contains 2 modules, so there will be num_spans x 2 modules
    # num_spans should be an odd number
    num_spans = 5
    assert num_spans % 2 != 0
    num_modules = 2*num_spans
    span_length = 2*module_length
    total_length = span_length * num_spans
    barrier_height = 1.37
    barrier_section = 'Barrier'
    damping_ratio = 0.003

    # default analysis is for through truss design. for gerber design, set is_gerber to true
    is_gerber = True

    ''' ------------------------ DEFINE LOADS ------------------------ '''

    deck_width = 3.0 # meters
    #deck_thickness = 0.15 # 150mm concrete deck
    #concrete_density = 24 # kN/m3
    deck_pressure = 1.85 # kPa (Simon calculation)
    asphalt_thickness = 0.004 # 4mm
    asphalt_density = 21 # kN/m3
    pedestrian_pressure = 3.86 # kPa (Simon calculation)
    pedestrian_density = 1.5 # p/m2
    barrier_load = 1.2 # kN/m
    snow_pressure = 1.3 # kPa (Simon calculation)
    trib_area = deck_width / 2.0

    # load factors
    dead_factor = 1.1
    live_factor = 1.7
    wearing_surface_factor = 1.5
    concrete_deck_factor = 1.2
    snow_factor = 1.5

    # compute UDLs
    live_UDL = trib_area * pedestrian_pressure
    barrier_UDL = barrier_load
    wearing_surface_UDL = asphalt_density * asphalt_thickness * trib_area
    concrete_deck_UDL = deck_pressure * trib_area
    snow_UDL = snow_pressure * trib_area
    roof_UDL = 0.5 # kN/m (Simon calculation)

    ''' ------------------------ INITIALIZE MODEL ------------------------ '''
    
    # generate geometry
    bottom_chord_points, top_chord_points, diagonal_web_points, vertical_web_points = generate_warren(
        height, module_length, module_divisions, segment_length, num_modules)
    
    # import the section combinations
    # will be a list of lists, each sublist has the combinations for different section types
    # ex. round, box, round-box combo
    section_combinations = create_section_combinations()
    
    # set file paths
    root_path = os.getcwd()
    base_file_path = root_path + '/BASE.sdb'
    os.makedirs('./models', exist_ok=True)
    model_path = root_path + '/models/MODEL.sdb'

    # open SAP application
    sap_object = sap_open()

    # delete old output file if it exists
    results_file = 'output.xlsx'
    results_path = root_path + os.sep + results_file
    if os.path.exists(results_path):
        os.remove(results_path)
    sheet_names = ['Box Box Box', 'Box Box Round']

    first_write = True
    for index, combination_type in enumerate(section_combinations):
        results = []
        sheet_name = sheet_names[index]

        for combo_index, combination in enumerate(tqdm(combination_type)):

            top_chord_section = combination[0]
            bottom_chord_section = combination[1]
            web_section = combination[2]

            # initialize fresh model from BASE in root folder
            sap_model = sap_initialize_model(base_file_path, sap_object)

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
            
            # brace bottom chord at the midpoint of each frame to the left and right of the support
            # this is to prevent those members from failing the kl/r_y check
            sap_brace_bottom_chord(sap_model, bottom_chord_frames, num_spans, module_divisions)
            
            # create barrier and apply vertical and horizontal barrier load patterns (cases created in sap_set_loads)
            barrier_frames = sap_barrier_load(sap_model, total_length, barrier_height, barrier_section, barrier_UDL)

            # set the load case, apply deck load to bottom chord
            sap_set_loads(sap_model, bottom_chord_frames, top_chord_frames, dead_factor, live_factor, 
                        wearing_surface_factor, concrete_deck_factor, snow_factor, live_UDL,
                        wearing_surface_UDL, concrete_deck_UDL, snow_UDL, roof_UDL)
            
            if is_gerber:
                vertical_web_frames, top_chord_frames, barrier_frames = sap_gerber_modification(
                    sap_model, vertical_web_frames, top_chord_frames, barrier_frames, num_spans, module_divisions)
                                    
            ''' ------------------------ RUN MODEL AND COLLECT RESULTS ------------------------ '''

            # save the file to a new file in the models folder (so don't override the BASE file)
            sap_run_analysis(sap_model, model_path)

            deflection, deflection_percentage = sap_deflection(sap_model, bottom_chord_frames, span_length)
            # get the reaction output from dead case and divide by num_modules
            module_mass = sap_module_mass(sap_model, num_modules)

            natural_frequency, in_crit_range, natural_frequency_occupied, resonating_harmonic, resonating_harmonic_occupied = sap_vibration_analysis(
                sap_model, pedestrian_density, concrete_deck_UDL, live_UDL, damping_ratio)
            # verify frames pass steel design check, and get list of sections that fail if ULS does not pass
            passed, failed_section_names = sap_steel_design(sap_model)

            results.append({
                'Top chord': top_chord_section,
                'Bottom chord': bottom_chord_section,
                'Web members': web_section,
                'Max vertical deflection for SLS (m)': deflection,
                'Percentage of deflection limit for SLS (%)': deflection_percentage,
                'Module mass (kg)': module_mass,
                'Natural frequency (Hz)': natural_frequency,
                'Natural frequency in critical range': in_crit_range,
                'Natural frequency occupied (Hz)': natural_frequency_occupied,
                'Resonating harmonic': resonating_harmonic,
                'Resonating harmonic occupied': resonating_harmonic_occupied,
                'Passed steel design check for ULS': passed,
                'Failed section': failed_section_names
            })

            # log results to console
            tqdm.write(f'Top chord section: {top_chord_section}, Bottom chord section: {bottom_chord_section}, Web member section: {web_section}')
            tqdm.write(f'Deflection of central node for SLS (mm): {round(deflection * 1000, 4)}')
            tqdm.write(f'Percentage of deflection limit for SLS (%): {round(deflection_percentage)}')
            tqdm.write(f'Mass of single module (kg): {round(module_mass, 4)}')
            tqdm.write(f'Natural frequency of span (Hz): {round(natural_frequency, 4)}')
            tqdm.write(f'Natural frequency in critical range?: {in_crit_range}')
            tqdm.write(f'Natural frequency of span occupied (Hz): {round(natural_frequency_occupied, 4)}')
            tqdm.write(f'Resonating harmonic of span: {round(resonating_harmonic, 4)}')
            tqdm.write(f'Resonating harmonic of span occupied: {round(resonating_harmonic_occupied, 4)}')
            tqdm.write(f'Passed steel design check for ULS: {passed}')
            tqdm.write(f'Failed section: {failed_section_names}')

            # write result to excel
            if combo_index % 10 == 0:
                write_to_excel(results, results_path, sheet_name, first_write)
                tqdm.write("Successfully updated output file.")
                first_write = False
            
            sap_model = None

    sap_close(sap_object)