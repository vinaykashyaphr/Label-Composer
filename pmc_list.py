from lxml import etree
import pathlib
from generalfunctions.helper_classes import Exclusion, NameAndCode

folderpath = pathlib.Path(input('Path:  '))
pmc = list(folderpath.glob('PMC-HON*-*.xml'))[0]
par = etree.XMLParser(no_network=True, recover=True)
pmroot = etree.parse(pmc, par).getroot()

dmcodeattribs = pmroot.findall('.//dmCode')

with open('pmc_list.txt', 'w', encoding='utf-8') as pmlist:
    for each in dmcodeattribs:
        if not str(NameAndCode().name_from_dmcode(each.attrib)).startswith('DMC-HONAERO'):
            pmlist.write(str(NameAndCode().name_from_dmcode(each.attrib))+'\n')