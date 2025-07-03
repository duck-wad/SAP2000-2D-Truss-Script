import xml.etree.ElementTree as ET
import pandas as pd

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

# write sections to excel file 
max_len = max(len(hss_round), len(hss_box))
hss_round += [None] * (max_len - len(hss_round))
hss_box += [None] * (max_len - len(hss_box))

# Create DataFrame
df = pd.DataFrame({
    'HSS Round': hss_round,
    'HSS Box': hss_box
})

output_path = './HSS_Sections.xlsx'
df.to_excel(output_path, index=False)