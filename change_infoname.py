import pathlib
import os
import re
import sys
from lxml import etree
import pandas as pd
from generalfunctions.helper_classes import (
    Exclusion, Write_DOC, NameAndCode, Validate_Entities)
from generalfunctions.nlp_functions import Case_Correction

app_path = pathlib.Path(os.getcwd())
FILENAME, TECHNAME, INFONAME = [], [], []


def get_dms_from_pm(folder: pathlib.Path, par: etree.XMLParser):
    pmc = list(folder.glob('PMC-HON*-*-0*-*.xml'))[0]
    pmroot = etree.parse(pmc, par).getroot()
    all_dmcodes = pmroot.findall('.//dmCode')
    return [{str(NameAndCode().name_from_dmcode(dmcode.attrib)): dmcode} for dmcode in all_dmcodes]


def techname_update(root: etree._Element, module: str, choice: bool, dm_list: list):
    techname = root.find('.//techName')
    if choice != False:
        steptitle = root.xpath(
            './/*[self::proceduralStep[not(parent::proceduralStep)] or self::levelledPara[not(parent::levelledPara)]]/title')
        figtitle = root.xpath(
            './/illustratedPartsCatalog/*[self::figure[title]]')
        if techname != None:
            if steptitle != []:
                techname.text = Case_Correction().sentence_case(
                    str(steptitle[0].text))
            elif figtitle != []:
                techname.text = Case_Correction().sentence_case(
                    str(figtitle[0].find('title').text))
            else:
                for each in dm_list:
                    if list(each.keys())[0] == module:
                        try:
                            techname.text = Case_Correction().sentence_case(each[module].xpath(
                                'ancestor:: pmEntry[1][pmEntryTitle]')[0].find('pmEntryTitle').text)
                        except IndexError:
                            techname.text = Case_Correction().sentence_case(techname.text)
    TECHNAME.append(techname.text)


def techname_change(root: etree._Element, content: str, choice: bool):
    techname = root.find('.//techName')
    if (choice != False) and (techname != None):
        techname.text = Case_Correction().sentence_case(content)
    TECHNAME.append(techname.text)


def infoname_change(root: etree._Element, content: str, choice: bool):
    infoname = root.find('.//infoName')
    if (choice != True) and ((infoname != None)):
        infoname.text = Case_Correction().sentence_case(content)
    INFONAME.append(infoname.text)


def infoname_update(root: etree._Element, choice: bool):
    infoname = root.find('.//infoName')
    if choice != True:
        infocode = root.find(
            './/dmIdent/dmCode/[@infoCode]').attrib['infoCode']
        brex = app_path.joinpath(r'infocodes.xml')
        brexroot = etree.parse(brex).getroot()
        query = brexroot.xpath(
            './/*[self::structureObjectRule/objectPath[@allowedObjectFlag="2"][text()="//dmIdent/dmCode/@infoCode"]]')[0]
        query_infoname = query.find(
            f'objectValue/[@valueForm="single"]/[@valueAllowed="{infocode}"]')
        if query_infoname != None:
            infoname.text = query_infoname.text
        else:
            query_patterns = query.findall(
                'objectValue/[@valueForm="pattern"]/[@valueAllowed]')
            for pattern in query_patterns:
                if re.match(pattern.attrib['valueAllowed'], infocode) != None:
                    infoname.text = pattern.text
    INFONAME.append(infoname.text)


def choose_bool(user_ip):
    if user_ip == 1:
        bool_ip = True
    elif user_ip == 0:
        bool_ip = False
    else:
        bool_ip = None
    return bool_ip


def infotech_change(folder: pathlib.Path, choice: bool, parser: etree.XMLParser):
    df = pd.read_excel(folder.joinpath('dm_label.xlsx'), 'LABEL')
    filenames = list(df['Filename'])
    technames = list(df['Techname'])
    infonames = list(df['Infoname'])
    for i, each in enumerate(filenames):
        filepath = list(folder.glob(f'{each}*.xml'))[0]
        Validate_Entities().valent(filepath.name, folder)
        dmroot = etree.parse(filepath.name, parser).getroot()
        FILENAME.append(NameAndCode().only_name(str(each)))
        print(filepath.name)
        techname_change(dmroot, technames[i], choice)
        infoname_change(dmroot, infonames[i], choice)
        Write_DOC().data_module(dmroot, filepath.name, folder)


def infotech_update(all_files: list, folder: pathlib.Path,
                    choice: bool, parser: etree.XMLParser, dm_list: list):
    for each in all_files:
        if each.startswith('DMC-HON') and (each.endswith('.xml') or each.endswith('.XML')):
            Validate_Entities().valent(each, folder)
            dmroot = etree.parse(each, parser).getroot()
            dmname = NameAndCode().only_name(str(each))
            FILENAME.append(dmname)
            print(each)
            techname_update(dmroot, dmname, choice, dm_list)
            infoname_update(dmroot, choice)
            Write_DOC().data_module(dmroot, each, folder)


def initiate():
    folderpath = pathlib.Path(sys.argv[1])
    all_files = Exclusion().parsable_list(folderpath)
    par = etree.XMLParser(no_network=True, recover=True)
    action_choice = choose_bool(int(sys.argv[2]))
    print(f'Action choice: {choose_bool(action_choice)}')
    function_choice = choose_bool(int(sys.argv[3]))
    print(f'Function choice: {choose_bool(function_choice)}')
    os.chdir(folderpath)
    dms_in_pmc = get_dms_from_pm(folderpath, par)
    [infotech_update(all_files, folderpath, function_choice, par, dms_in_pmc)
        if (action_choice == True) else infotech_change(folderpath, function_choice, par)]
    infotech_df = pd.DataFrame(
        {'Filename': FILENAME, 'Techname': TECHNAME, 'Infoname': INFONAME})
    excel = pd.ExcelWriter('dm_label.xlsx')
    infotech_df.to_excel(excel, 'LABEL', index=False)
    excel.close()


initiate()
