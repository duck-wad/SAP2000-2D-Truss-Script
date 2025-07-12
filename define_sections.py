import xml.etree.ElementTree as ET
import pandas as pd
import os
from openpyxl import load_workbook, Workbook

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

def filter_HSS_sections(sections, min_depth, min_thick, max_depth, max_thick, asym = False):
    filtered_sections = []
    for section in sections:
        # name is in format HS###X## (round) and HS###X###X## (box)
        # split section name into before and after 'X'
        # depth is always the first dim and thickness always the last dim
        parts = section.split('X')
        # get diameter as everything after 'HS' and before 'X', and convert to int
        depth = int(parts[0][2:])
        # get numbers after 'X' and convert to float
        thick = float(parts[-1])
        # get width if applicable
        width = depth
        if len(parts) == 3:
            width = int(parts[1])

        if depth >= min_depth and thick >= min_thick and depth <= max_depth and thick <= max_thick:
            if asym == True and width != depth:
                pass
            else:
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

# we are limiting bottom and top chord to be box only for connection purposes
# top chord > bottom chord > web
# for box limit to square section no rectangle (for now to limit combinations)
def create_section_combinations():

    round, box = load_xml()

    ''' ------------------------ FILTER ROUND SECTIONS ------------------------ '''
    # web is smaller than bottom chord. don't specify a min or max diameter, that will be taken care of in 
    # the valid_combinations function. limit thickness 7.9-9.5
    web_round = filter_HSS_sections(round, 0, 7.9, 500, 9.5)

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
    
    # limit top chord to be depth 200+ and thickness 7.9-9.5
    top_chord_box = filter_HSS_sections(box, 200, 7.9, 356, 9.5)
    # limit bottom chord depth no bottom limit, 305 top, 7.9-9.5
    bottom_chord_box = filter_HSS_sections(box, 0, 7.9, 305, 9.5)
    # limit web depth no bottom limit, top limit 203
    # limit the web to be only square sections
    web_box = filter_HSS_sections(box, 0, 7.9, 203, 9.5, asym=True)

    ''' ------------------------ CREATE COMBINATIONS ------------------------ '''
    
    # [top, bottom, web]
    box_box_box = valid_combinations(top_chord_box, bottom_chord_box, web_box)
    box_box_round = valid_combinations(top_chord_box, bottom_chord_box, web_round)   
    return [box_box_box, box_box_round]

def write_to_excel(results, path, sheet, first_write=False):

    df = pd.DataFrame(results)

    if first_write:
        with pd.ExcelWriter(path, mode='w') as writer:
                df.to_excel(writer, sheet_name=sheet, index=False)

    else:
        with pd.ExcelWriter(path, mode='a', if_sheet_exists='replace',) as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)
