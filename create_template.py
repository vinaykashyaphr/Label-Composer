from lxml import etree
import pathlib
import os
import sys
import pandas as pd
from generalfunctions.helper_classes import (
    Exclusion, Validate_Entities, NameAndCode)
from generalfunctions.nlp_functions import Case_Correction

folderpath = pathlib.Path(sys.argv[1])
os.chdir(folderpath)
all_files = Exclusion().parsable_list(folderpath)
par = etree.XMLParser(no_network=True, recover=True)
pmc = list(folderpath.glob('PMC-HON*-*-0*-*.xml'))[0]
Validate_Entities().valent(pmc, folderpath)
pmroot = etree.parse(pmc, par).getroot()
pmes = pmroot.xpath(
    './/*[self:: pmEntry[not(parent:: pmEntry)] or self:: applicCrossRefTableRef]')
FILE, TECH, INFO = [], [], []


for pme in pmes:
    dmcodes = pme.findall('.//dmCode')
    for dmcode in dmcodes:
        dmname = NameAndCode().name_from_dmcode(dmcode.attrib)
        print(dmname, '<<<<')
        filepath = list(folderpath.glob(f'{dmname}*.xml'))[0]
        if filepath.name in all_files:
            print(filepath.name)
            Validate_Entities().valent(filepath.name, folderpath)
            dmroot = etree.parse(filepath.name, par).getroot()
            if dmname not in FILE:
                FILE.append(dmname)
                TECH.append(Case_Correction().sentence_case(
                    str(dmroot.find('.//techName').text)))
                INFO.append(dmroot.find('.//infoName').text)

pct = list(folderpath.glob('DMC-HON*-*-*-*-*-*-00P*-C*.xml'))[0]
print(pct.name)
pct_root = etree.parse(pct.name, par).getroot()
FILE.append(NameAndCode().only_name(pct.name))
TECH.append(Case_Correction().sentence_case(pct_root.find('.//techName').text))
INFO.append(pct_root.find('.//infoName').text)

excel = pd.ExcelWriter('dm_label.xlsx')
df = pd.DataFrame({'Filename': FILE, 'Techname': TECH, 'Infoname': INFO})
df.to_excel(excel, 'LABEL', index=False)
excel.close()
