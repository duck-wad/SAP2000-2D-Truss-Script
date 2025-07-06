import xml.etree.ElementTree as ET
import pandas as pd

def load_xml():
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

    # write sections to excel file for easier reference
    max_len = max(len(hss_round), len(hss_box))
    hss_round_excel = hss_round.copy()
    hss_box_excel = hss_box.copy()
    hss_round_excel += [None] * (max_len - len(hss_round))
    hss_box_excel += [None] * (max_len - len(hss_box))

    # Create DataFrame
    df = pd.DataFrame({
        'HSS Round': hss_round_excel,
        'HSS Box': hss_box_excel
    })

    output_path = './sections.xlsx'
    df.to_excel(output_path, index=False)

    return hss_round, hss_box

def filter_HSS_sections(sections, min_depth, min_thick, max_depth, max_thick):
    filtered_sections = []
    for section in sections:
        # name is in format HS###X## (round) and HS###X###X## (box)
        # split section name into before and after 'X'
        # depth is always the first dim and thickness always the last dim
        # don't need to filter based on width
        parts = section.split('X')
        # get diameter as everything after 'HS' and before 'X', and convert to int
        depth = int(parts[0][2:])
        # get numbers after 'X' and convert to float
        thick = float(parts[-1])

        if depth >= min_depth and thick >= min_thick and depth <= max_depth and thick <= max_thick:
            filtered_sections.append(section)

    return filtered_sections

def valid_combinations(top_sections, bottom_sections, web_sections):
    combinations = []
    # to limit number of combinations, specify a minimum difference between the top and bottom chord
    # as well as web and bottom chord of 50mm
    min_d_diff = 50
    for top in top_sections:
        temp = top.split('X')
        top_d = int(temp[0][2:])
        for bottom in bottom_sections:
            temp = bottom.split('X')
            bottom_d = int(temp[0][2:])
            if (top_d - bottom_d) >= min_d_diff:
                for web in web_sections:
                    temp = web.split('X')
                    web_d = int(temp[0][2:])
                    if (bottom_d - web_d) >= min_d_diff:
                        combinations.append([top, bottom, web])
    return combinations

def create_section_combinations():

    round, box = load_xml()

    ''' ------------------------ FILTER ROUND SECTIONS ------------------------ '''
    # top chord is largest because under compression. limit size to be above 219mm diam. and 8mm thick
    top_chord_round = filter_HSS_sections(round, 200, 8, 500, 20)
    # bottom chord is smaller because under tension
    # limit diam to be between 150-350 and thickness 6.4-13
    bottom_chord_round = filter_HSS_sections(round, 150, 6.4, 350, 13)
    # web is smallest, limit diam between 100-250 and thickness 4-10
    web_round = filter_HSS_sections(round, 100, 4, 250, 10)

    # [top, bottom, web]
    round_combinations = valid_combinations(top_chord_round, bottom_chord_round, web_round)

    ''' ------------------------ FILTER BOX SECTIONS ------------------------ '''

    # sort box sections based on size (depth, width, thickness)
    # in xml they are ordered by square first then rectangle
    box = sorted(box, key=lambda x: (
    int(x[2:x.index('X')]),                  
    int(x[x.index('X')+1:x.index('X', x.index('X')+1)]),  
    float(x[x.index('X', x.index('X')+1)+1:])  
    ))
    # reverse list
    box.reverse()
    
    # limit top chord to be depth 250-356 and thickness 9-16
    top_chord_box = filter_HSS_sections(box, 250, 9, 356, 16)
    # limit bottom chord depth 200-305 and thickness 6-13
    bottom_chord_box = filter_HSS_sections(box, 200, 6, 305, 13)
    # limit web depth 150-254 and thickness 4.8-13
    web_box = filter_HSS_sections(box, 150, 4.8, 254, 13)

    box_combinations = valid_combinations(top_chord_box, bottom_chord_box, web_box)

    # get combinations for top and bottom chord box and round web
    box_round_combinations = valid_combinations(top_chord_box, bottom_chord_box, web_round)

    #return [round_combinations, box_combinations, box_round_combinations]
    return [round_combinations[0:5], box_combinations[0:5], box_round_combinations[0:5]]
