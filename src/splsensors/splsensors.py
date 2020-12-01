# -*- coding: utf-8 -*-
###############################################################
# Author:       patrice.ponchant@furgo.com  (Fugro Brasil)    #
# Created:      27/11/2020                                    #
# Python :      3.x                                           #
###############################################################

# The future package will provide support for running your code on Python 2.6, 2.7, and 3.3+ mostly unchanged.
# http://python-future.org/quickstart.html
from __future__ import (absolute_import, division,
                        print_function, unicode_literals)
from builtins import *

##### Sensor reader packages #####
from pyXTF import * # https://github.com/pktrigg/pyxtf
from pyall import * # https://github.com/pktrigg/pyall
from obspy import read

##### Basic packages #####
import datetime
import sys, os, glob, shutil
import subprocess
import pandas as pd
import numpy as np
import math

##### CMD packages #####
from tqdm import tqdm
#from tabulate import tabulate

##### GUI packages #####
from gooey import Gooey, GooeyParser
from colored import stylize, attr, fg

# 417574686f723a205061747269636520506f6e6368616e74
##########################################################
#                       Main code                        #
##########################################################
# this needs to be *before* the @Gooey decorator!
# (this code allows to only use Gooey when no arguments are passed to the script)
if len(sys.argv) >= 2:
    if not '--ignore-gooey' in sys.argv:
        sys.argv.append('--ignore-gooey')
        cmd = True 
    else:
        cmd = False  
        
# GUI Configuration
@Gooey(
    program_name='Linename comparison and rename tool between SPL and sensors',
    progress_regex=r"^progress: (?P<current>\d+)/(?P<total>\d+)$",
    progress_expr="current / total * 100",
    hide_progress_msg=True,
    richtext_controls=True,
    #richtext_controls=True,
    terminal_font_family = 'Courier New', # for tabulate table nice formatation
    #dump_build_config=True,
    #load_build_config="gooey_config.json",
    default_size=(800, 750),
    timing_options={        
        'show_time_remaining':True,
        'hide_time_remaining_on_complete':True
        },
    tabbed_groups=True,
    navigation='Tabbed',
    header_bg_color = '#95ACC8',
    #body_bg_color = '#95ACC8',
    menu=[{
        'name': 'File',
        'items': [{
                'type': 'AboutDialog',
                'menuTitle': 'About',
                'name': 'spl-sensors-comp-ren',
                'description': 'Linename comparison and rename tool between SPL and sensors',
                'version': '0.1.3',
                'copyright': '2020',
                'website': 'https://github.com/Shadoward/spl-sensors-comp-ren',
                'developer': 'patrice.ponchant@fugro.com',
                'license': 'MIT'
                }]
        },{
        'name': 'Help',
        'items': [{
            'type': 'Link',
            'menuTitle': 'Documentation',
            'url': ''
            }]
        }]
    )

def main():
    desc = "Linename comparison and rename tool between SPL and sensors"    
    parser = GooeyParser(description=desc)
    
    splopt = parser.add_argument_group('SPL Options', gooey_options={'columns': 1})
    sensorsopt = parser.add_argument_group('Sensors Options', gooey_options={'columns': 1})
    outputsopt = parser.add_argument_group('Output Options', gooey_options={'columns': 1})
    additionalopt = parser.add_argument_group('Additional Options', gooey_options={'columns': 1})
    renameopt = parser.add_argument_group('Rename Options')
     
    # SPL Arguments
    splopt.add_argument(
        '-i',
        '--splFolder', 
        #action='store',
        dest='splFolder',       
        metavar='SPL Root Path', 
        help='This is the path where the *.fbf/*.fbz/*.pos files to process are. (Root Session Folder)',
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\NEL',
        #default_path='S:\\JOBS\\2020\\20030002_Shell_FBR_MF\\B2B_FromVessel\\Navigation\\Starfix_Logging\\RawData', 
        widget='DirChooser')
        #type=str,)
    splopt.add_argument(
        '-p', '--splPosition', 
        #action='store',
        dest='splPosition',
        metavar='SPL Position File Name', 
        widget='TextField',
        #type=str,
        #default='FugroBrasilis-CRP-Position',
        help='SPL position file to be use to rename the sensor without extention.',)
        
    # Sensors Arguments
    sensorsopt.add_argument(
        '-A', '--allFolder', 
        #action='store',
        dest='allFolder',        
        metavar='ALL Folder Path',
        help='ALL Root path. This is the path where the *.all files to process are.',
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\ALL', 
        widget='DirChooser')
        #type=str,)
    sensorsopt.add_argument(
        '-X', '--xtfFolder', 
        #action='store',
        dest='xtfFolder',        
        metavar='XTF Folder Path',
        help='XTF Root path. This is the path where the *.xtf files to process are.',
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\XTF', 
        widget='DirChooser')
        #type=str,)
    sensorsopt.add_argument(
        '-S', '--sgySBPFolder', 
        #action='store',
        dest='sgySBPFolder',        
        metavar='SGY/SEG/SEGY SBP Folder Path',
        help='SGY/SEG/SEGY SBP Root path. This is the path where the *.sgy/*.seg/*.segy files to process are.',
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\SGYSBP', 
        widget='DirChooser')
        #type=str,)
    #sensorsopt.add_argument(
    #    '-E', '--ses3SBPFolder', 
    #    #action='store',
    #    dest='ses3SBPFolder',        
    #    metavar='SES3 SBP Folder Path',
    #    help='SES3 SBP Root path. This is the path where the *.ses3 files to process are.',
    #    default='C:\\Users\\patrice.ponchant\\Downloads\\SES3SBP', 
    #    widget='DirChooser',
    #    type=str,
    # )
    sensorsopt.add_argument(
        '-M', '--csvMAGFolder', 
        #action='store',
        dest='csvMAGFolder',        
        metavar='CSV MAG Folder Path',
        help='CSV MAG Root path. This is the path where the *.csv files to process are.',
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\CSVMAG', 
        widget='DirChooser')
        #type=str,)
    sensorsopt.add_argument(
        '-H', '--sgySUHRSFolder', 
        #action='store',
        dest='sgySUHRSFolder',        
        metavar='SGY/SEG/SEGY SUHRS Folder Path',
        help='SGY/SEG/SEGY SUHRS Root path. This is the path where the *.sgy/*.seg/*.segy files to process are.',
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\SGYSUHRS', 
        widget='DirChooser')
        #type=str,)
    
    # Output Arguments
    outputsopt.add_argument(
        '-o', '--output',
        #action='store',
        dest='outputFolder',
        metavar='Output Logs Folder',  
        help='Output folder to save all the logs files.',      
        #type=str,
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\LOGS',
        widget='DirChooser')
    
    # Additional Arguments
    additionalopt.add_argument(
        '-r', '--recursive',
        dest='recursive',
        metavar='Recurse into the subfolders?', 
        choices=['yes', 'no'], 
        default='yes')
    additionalopt.add_argument(
        '-m', '--move',
        dest='move',
        metavar='Move MAG and SUHRS in the vessel folder?', 
        help='This will create and vessel folder in the sensor folder basaed on the SPL name vessel and move the files to this.',
        choices=['yes', 'no'], 
        default='yes')
    additionalopt.add_argument(
        '-e', '--excludeFolder', 
        #action='store',
        dest='excludeFolder',
        metavar='List of folder to be exclude', 
        widget='TextField',
        type=str,
        #default='DNP, DoNotProcess',
        help='List all folder that need to be excluded from the recurcive search.\n(eg.: DNP,DoNotProcess) Comma separated and NO WHITESPACE!\nNote: This just apply to the sensors folders')       

    # Rename Option
    renameopt.add_argument(
        '-n', '--rename',
        dest='rename',
        metavar='Rename the files?', 
        choices=['yes', 'no'], 
        default='no')
    
    # Use to create help readme.md. TO BE COMMENT WHEN DONE
    # if len(sys.argv)==1:
    #    parser.print_help()
    #    sys.exit(1)   
        
    args = parser.parse_args()
    process(args, cmd)

def process(args, cmd):
    """
    Uses this if called as __main__.
    """
    splFolder = args.splFolder
    splPosition = args.splPosition
    vessel = splPosition.split('-')[0]
    allFolder = args.allFolder
    xtfFolder = args.xtfFolder
    #ses3SBPFolder = args.ses3SBPFolder
    sgySBPFolder = args.sgySBPFolder
    csvMAGFolder = args.csvMAGFolder
    sgySUHRSFolder = args.sgySUHRSFolder
    outputFolder = args.outputFolder
    excludeFolder = args.excludeFolder
    move = args.move
    
    # Defined Global Dataframe
    dfFINAL = pd.DataFrame(columns = ["Session Start", "Session End", "Vessel Name", "SPL", "MBES", "SBP", "SSS", "MAG", "SUHRS"])
    dfSPL = pd.DataFrame(columns = ["Session Start", "Session End", "SPL LineName"])       
    dfer = pd.DataFrame(columns = ["SPLPath"])
    dfSummary = pd.DataFrame(columns = ["Sensor", "Processed Files", "Duplicated Files", "Wrong Timestamp (SBP)", "Moved Files"])
    dfMissingSPL = pd.DataFrame(columns = ["Session Start", "Session End", "Vessel Name", "Sensor Start",
                                           "FilePath", "Sensor FileName", "SPL LineName", "Sensor New LineName"])
    dfDuplSensor = pd.DataFrame(columns = ["Session Start", "Session End", "Vessel Name", "Sensor Start",
                                           "FilePath", "Sensor FileName", "SPL LineName", "Sensor New LineName"])
    dfSkip = pd.DataFrame(columns = ["FilePath", "File Size [MB]"])
    
    # Check if SPL is a position file
    if splPosition.find('-Position') == -1:
        print('')
        sys.exit(stylize(f"The SPL file {splPosition} is not a position file, quitting", fg('red')))
    
    # Check if SPL folder is defined
    if not splFolder:
        print ('')
        sys.exit(stylize('No SPL Folder was defined, quitting', fg('red')))
    
    # Check if Output folder is defined    
    if not outputFolder:
        print ('')
        sys.exit(stylize('No Output Folder was defined, quitting', fg('red')))
    
    # Check if file is open before continue
    for fi in glob.glob(outputFolder + "\\*"):
        try:
            fp = open(fi, "r+")
        except IOError:
            print('')
            sys.exit(stylize(f'The following file is lock ({fi}).\nPlease close the files in the "{outputFolder}" folder', fg('red')))      
    
    ##########################################################
    #                 Listing the files                      #
    ##########################################################  
    print('')
    print('##################################################')
    print('LISTING FILES. PLEASE WAIT....')
    print('##################################################')
    nowLS = datetime.datetime.now() # record time of the subprocess
    allListFile = []
    xtfListFile = []
    sbpListFile = []
    #ses3ListFile = []
    magListFile = []
    suhrsListFile = []
    fbzListFile = []
    fbfListFile = []
    posListFile = []
       
    if args.recursive == 'no':
        print(f'Start Listing for {allFolder}') if args.allFolder is not None else []
        allListFile = glob.glob(allFolder + "\\*.all") if args.allFolder is not None else []
        print(f'Start Listing for {xtfFolder}') if args.xtfFolder is not None else []
        xtfListFile = glob.glob(xtfFolder + "\\*.xtf") if args.xtfFolder is not None else []
        print(f'Start Listing for {sgySBPFolder}') if args.sgySBPFolder is not None else []
        sbpListFile = glob.glob(sgySBPFolder + "\\*.s*g*") if args.sgySBPFolder is not None else []
        #print(f'Start Listing for {ses3SBPFolder}') if args.ses3SBPFolder is not None else []
        #ses3ListFile = glob.glob(ses3SBPFolder + "\\*.ses3") if args.ses3SBPFolder is not None else []
        print(f'Start Listing for {csvMAGFolder}') if args.csvMAGFolder is not None else []
        magListFile = glob.glob(csvMAGFolder + "\\*.csv") if args.csvMAGFolder is not None else []
        print(f'Start Listing for {sgySUHRSFolder}') if args.sgySUHRSFolder is not None else []
        suhrsListFile = glob.glob(sgySUHRSFolder + "\\*.s*g*") if args.sgySUHRSFolder is not None else []        
    else:
        exclude = excludeFolder.split(",") if args.excludeFolder is not None else []
        print(f'The following folders(s) will be excluded: {exclude} \n') if args.excludeFolder is not None else []
        #exclude = ['DNP', 'DoNotProcess']
        allListFile = listFile(allFolder, "all", set(exclude)) if args.allFolder is not None else []
        xtfListFile = listFile(xtfFolder, "xtf", set(exclude)) if args.xtfFolder is not None else []
        sbpListFile = listFile(sgySBPFolder, ('seg', 'sgy', 'segy'), set(exclude)) if args.sgySBPFolder is not None else []
        #ses3ListFile = listFile(ses3SBPFolder, "ses3", set(exclude)) if args.ses3SBPFolder is not None else []
        magListFile = listFile(csvMAGFolder, "csv", set(exclude)) if args.csvMAGFolder is not None else []
        suhrsListFile = listFile(sgySUHRSFolder, ('seg', 'sgy', 'segy'), set(exclude)) if args.sgySUHRSFolder is not None else []
    
    #### Not very efficient, will need to improve later   
    print(f'Start Listing for {splFolder}')
    fbzListFile = glob.glob(splFolder + "\\**\\" + splPosition + ".fbz", recursive=True)
    fbfListFile = glob.glob(splFolder + "\\**\\" + splPosition + ".fbf", recursive=True)
    posListFile = glob.glob(splFolder + "\\**\\" + splPosition + ".pos", recursive=True)            
    print("Subprocess Duration: ", (datetime.datetime.now() - nowLS))
    
    # Check if SPL files found
    if not fbfListFile and not fbzListFile:
        print ('')
        sys.exit(stylize('No SPL files were found, quitting', fg('red')))
    
    print('')
    print(f'Total of files that will be processed: \n {len(fbfListFile)} *.fbf \n {len(fbzListFile)} *.fbz \n {len(posListFile)} *.pos \n {len(allListFile)} *.all \n {len(xtfListFile)} *.xtf \n {len(sbpListFile)} *.sgy/*.seg/*.segy (SBP) \n {len(magListFile)} *.csv (MAG) \n {len(suhrsListFile)} *.sgy/*.seg/*.segy (SUHRS)') 
    
    ##########################################################
    #                     Reading SPL                        #
    ##########################################################    
    dfSPL, dfer = splfc(fbfListFile, 'FBF', dfSPL, dfer, outputFolder, cmd)
    dfSPL, dfer = splfc(fbzListFile, 'FBZ', dfSPL, dfer, outputFolder, cmd)
    dfSPL, dfer = splfc(posListFile, 'POS', dfSPL, dfer, outputFolder, cmd)
    
    # Copy the needed info from dfSPl to the dfFINAL
    dfFINAL['Session Start'] = dfSPL['Session Start']  
    dfFINAL['Session End'] = dfSPL['Session End']
    dfFINAL['SPL'] = dfSPL['SPL LineName']
    dfFINAL['Vessel Name'] = vessel
    
    ##########################################################
    #                    Reading Sensors                     #
    ##########################################################
    if allListFile:
       dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip = sensorsfc(allListFile, 'MBES', '.all', cmd, args.rename, outputFolder,
                                                                  dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, vessel)
    if xtfListFile: 
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip = sensorsfc(xtfListFile, 'SSS', '.xtf', cmd, args.rename, outputFolder,
                                                                   dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, vessel)
    if sbpListFile: 
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip = sensorsfc(sbpListFile, 'SBP', '.sgy/*.seg/*.segy', cmd, args.rename, outputFolder,
                                                                   dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, vessel)
    #if args.ses3SBPFolder is not None: 
    #    sensorsfc(ses3ListFile, 'SES3', '.ses3', cmd, args.rename, outputFolder, dfSensors, dfSPL, dftmp, vessel)
    
    if magListFile: 
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip = sensorsfc(magListFile, 'MAG', '.csv', cmd, args.rename, outputFolder,
                                                                   dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, vessel)
        if move == 'yes':
            lsFile = outputFolder + "\\" + vessel + "_MAG_Full_Log.csv"
            VesselFolder = os.path.join(csvMAGFolder, vessel)
            WrongFolder = os.path.join(csvMAGFolder, 'WRONG')
            dfSummary = mvSensorFile (lsFile, VesselFolder, WrongFolder, cmd, 'MAG', dfSummary) 
    
    if suhrsListFile: 
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip = sensorsfc(suhrsListFile, 'SUHRS', '.sgy/*.seg/*.segy', cmd, args.rename, outputFolder,
                                                                   dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, vessel)
        if move == 'yes':
            lsFile = outputFolder + "\\" + vessel + "_SUHRS_Full_Log.csv"
            VesselFolder = os.path.join(sgySUHRSFolder, vessel)
            WrongFolder = os.path.join(sgySUHRSFolder, 'WRONG')
            dfSummary = mvSensorFile (lsFile, VesselFolder, WrongFolder, cmd, 'SUHRS', dfSummary) 

    ##########################################################
    #                  Excel Exportation                     #
    ##########################################################
    print('')
    print('##################################################')
    print('EXPORTING THE LOGS IN EXCEL')
    print('##################################################')       
    print('')
    nowExcel = datetime.datetime.now()  # record time of the subprocess 
    
    # For stats
    Tfbf_fbz = len(fbfListFile) + len(fbzListFile) + len(posListFile)
    Tsensors = len(allListFile) + len(xtfListFile) + len(sbpListFile) + len(magListFile) + len(suhrsListFile)
    
    # creating the df for every sheet in excel
    dfMissingSPL = dfMissingSPL[dfMissingSPL['Session Start'].isnull()] # df Missing SPL
    dfMissingSPL = dfMissingSPL.drop(columns=["Session Start", "Session End", "Vessel Name", "SPL LineName", "Sensor New LineName"])
    dfMissingSPL = dfMissingSPL[["Sensor Start", "Sensor FileName", "FilePath"]]
    
    dfSPLProblem = dfFINAL[dfFINAL['SPL'].isin(['NoLineNameFound', 'EmptySPL', 'SPLtoSmall'])]
    dfSPLProblem = dfSPLProblem.drop_duplicates(subset='Session Start', keep='first')
    
    dfMBES = dfFINAL[dfFINAL.apply(lambda row: '[WRONG]' in str(row.MBES), axis=1)].sort_values('Session Start')
    dfMBES = dfMBES.append(dfFINAL[dfFINAL.MBES.isnull()].sort_values('Session Start')) #https://stackoverflow.com/questions/29314033/drop-rows-containing-empty-cells-from-a-pandas-dataframe/56708633#56708633
    
    dfSSS = dfFINAL[dfFINAL.apply(lambda row: '[WRONG]' in str(row.SSS), axis=1)].sort_values('Session Start')
    dfSSS = dfSSS.append(dfFINAL[dfFINAL.SSS.isnull()].sort_values('Session Start'))
    dfSSS = dfSSS[["Session Start", "Session End", "Vessel Name", "SPL", "SSS", "MBES", "SBP", "MAG", "SUHRS"]]
    
    dfSBP = dfFINAL[dfFINAL.apply(lambda row: '[WRONG]' in str(row.SBP), axis=1)].sort_values('Session Start')
    dfSBP = dfSBP.append(dfFINAL[dfFINAL.SBP.isnull()].sort_values('Session Start'))
    dfSBP = dfSBP[["Session Start", "Session End", "Vessel Name", "SPL", "SBP", "MBES", "SSS", "MAG", "SUHRS"]]
    
    dfMAG = dfFINAL[dfFINAL.apply(lambda row: '[WRONG]' in str(row.MAG), axis=1)].sort_values('Session Start')
    dfMAG = dfMAG.append(dfFINAL[dfFINAL.MAG.isnull()].sort_values('Session Start'))
    dfMAG = dfMAG[["Session Start", "Session End", "Vessel Name", "SPL", "MAG", "MBES", "SSS", "SBP", "SUHRS"]]
    
    dfSUHRS = dfFINAL[dfFINAL.apply(lambda row: '[WRONG]' in str(row.SUHRS), axis=1)].sort_values('Session Start')
    dfSUHRS = dfSUHRS.append(dfFINAL[dfFINAL.SUHRS.isnull()].sort_values('Session Start'))
    dfSUHRS = dfSUHRS[["Session Start", "Session End", "Vessel Name", "SPL", "SUHRS", "MBES", "SSS", "SBP", "MAG"]]
    
    dfDuplSPL = dfFINAL[dfFINAL.duplicated(subset='SPL', keep=False)]
    #print(dfFINAL.apply(lambda row: str(row.MBES) in '[WRONG]', axis=1))
    
    writer = pd.ExcelWriter(outputFolder + "\\_" + vessel + '_FINAL_Log.xlsx', engine='xlsxwriter', datetime_format='dd mmm yyyy hh:mm:ss.000')
            
    dfSummary.to_excel(writer, sheet_name='Summary_Process_Log', startrow=5)    
    dfFINAL.sort_values('Session Start').to_excel(writer, sheet_name='Full_List') 
    dfMissingSPL.sort_values('Sensor Start').to_excel(writer, sheet_name='Missing_SPL')   
    dfMBES.to_excel(writer, sheet_name='MBES_NotMatching')
    dfSSS.to_excel(writer, sheet_name='SSS_NotMatching')
    dfSBP.to_excel(writer, sheet_name='SBP_NotMatching')
    dfMAG.to_excel(writer, sheet_name='MAG_NotMatching')
    dfSUHRS.to_excel(writer, sheet_name='SUHRS_NotMatching')    
    dfDuplSPL.sort_values('SPL').to_excel(writer, sheet_name='Duplicated_SPL_Name')
    dfDuplSensor.sort_values('Sensor Start').to_excel(writer, sheet_name='Duplicated_Sensor_Data')    
    dfSPLProblem.sort_values('Session Start').to_excel(writer, sheet_name='SPL_Problem')
    dfSkip.to_excel(writer, sheet_name='Skip_SSS_Files')
    
    workbook  = writer.book        
    worksheetS = writer.sheets['Summary_Process_Log']
    worksheetS.hide_gridlines(2)
    worksheetFull = writer.sheets['Full_List']
    worksheetMBES = writer.sheets['MBES_NotMatching']
    worksheetSSS = writer.sheets['SSS_NotMatching']
    worksheetSBP = writer.sheets['SBP_NotMatching']
    worksheetMAG = writer.sheets['MAG_NotMatching']
    worksheetSUHRS = writer.sheets['SUHRS_NotMatching']
    worksheetDuplSPL = writer.sheets['Duplicated_SPL_Name']
    worksheetDuplSensor = writer.sheets['Duplicated_Sensor_Data']
    worksheetMissingSPL = writer.sheets['Missing_SPL']
    worksheetSkip = writer.sheets['SPL_Problem']
    
    ListDF = [dfFINAL, dfMBES, dfSSS, dfSBP, dfMAG, dfSUHRS, dfDuplSPL, dfDuplSensor, dfMissingSPL, dfSPLProblem, dfSkip]
    ListWS = [worksheetFull, worksheetMBES, worksheetSSS, worksheetSBP, worksheetMAG, worksheetSUHRS, 
                worksheetDuplSPL, worksheetDuplSensor, worksheetMissingSPL, worksheetSkip]
    ListFC = [worksheetFull, worksheetMBES, worksheetSSS, worksheetSBP, worksheetMAG, worksheetSUHRS]
    
    #### Set format       
    # Summary Sheet
    bold = workbook.add_format({'bold': True,
                                'font_name': 'Segoe UI',
                                'font_size': 10,
                                'valign': 'vcenter',})
    normal = workbook.add_format({'bold': False,
                                'font_name': 'Segoe UI',
                                'font_size': 10,
                                'valign': 'vcenter',})
    hlink = workbook.add_format({'bold': False,
                                 'font_color': '#0250AE',
                                 'underline': True,
                                 'font_name': 'Segoe UI',
                                 'font_size': 10,
                                 'valign': 'vcenter',})
  
    for i, width in enumerate(get_col_widths(dfSummary)):
        worksheetS.set_column(i, i, width)
        
    text1 = 'A total of ' + str(Tsensors) + ' sensors files were processed and ' + str(Tfbf_fbz) + ' SPL Sessions.'
    text2 = 'A total of ' + str(len(dfer)) + '/' + str(Tfbf_fbz) + ' Session SPL has/have no Linename information.' 
    textS = [bold, 'Summary_Process_Log', normal, ': Summary log of the processing']
    textFull = [bold, 'Full_List', normal, ': Full log list of all sensors without duplicated and skip files']
    textMissingSPL = [bold, 'Missing_SPL', normal, ': List of all missing SPL file that do have sensors recorded']
    textMBES = [bold, 'MBES_NotMatching', normal, ': MBES log list of all files that do not match the SPL name; without duplicated and skip files']
    textSSS = [bold, 'SSS_NotMatching', normal, ': SSS log list of all files that do not match the SPL name; without duplicated and skip files']
    textSBP = [bold, 'SBP_NotMatching', normal, ': SBP log list of all files that do not match the SPL name; without duplicated and skip files']
    textMAG = [bold, 'MAG_NotMatching', normal, ': MAG log list of all files that do not match the SPL name; without duplicated and skip files']
    textSUHRS = [bold, 'SUHRS_NotMatching', normal, ': SUHRS log list of all files that do not match the SPL name; without duplicated and skip files']
    textDuplSPL = [bold, 'Duplicated_SPL_Name', normal, ': List of all duplicated SPL name']
    textDuplSensor = [bold, 'Duplicated_Sensor_Data', normal, ': List of all duplicated sensors files; Based on the start time']
    textSkip = [bold, 'SPL_Problem', normal, ': List of all SPL session without a line name in the columns LineName or are empty ou too small']
    
    ListT = [textS, textMissingSPL, textFull, textMBES, textSSS, textSBP, textMAG, textSUHRS, textDuplSPL, textDuplSensor, textSkip]
    ListHL = ['internal:Summary_Process_Log!A1', 'internal:Full_List!A1', 'internal:Missing_SPL!A1', 'internal:MBES_NotMatching!A1', 
              'internal:SSS_NotMatching!A1', 'internal:SBP_NotMatching!A1', 'internal:MAG_NotMatching!A1','internal:SUHRS_NotMatching!A1',
              'internal:Duplicated_SPL_Name!A1', 'internal:Duplicated_Sensor_Data!A1', 'internal:SPL_Problem!A1']
                
    worksheetS.write(0, 0, text1, bold)
    worksheetS.write(1, 0, text2, bold)
    
    icount = 0
    for e, l in zip(ListT,ListHL):
        worksheetS.write_rich_string(dfSummary.shape[0] + 7 + icount, 1, *e)
        worksheetS.write_url(dfSummary.shape[0] + 7 + icount, 0, l, hlink, string='Link')
        icount += 1
    
    for ws in ListWS:
        ws.write_url(0, 0, 'internal:Summary_Process_Log!A1', hlink, string='Summary')
    
    # Others Sheet
    cell_format = workbook.add_format({'text_wrap': True,
                                       'font_name': 'Segoe UI',
                                       'font_size': 10,
                                       'valign': 'vcenter',
                                       'align': 'left',
                                       'border_color': '#000000',
                                       'border': 1})
    
    header_format = workbook.add_format({'bold': True,
                                         'font_name': 'Segoe UI',
                                         'font_size': 12,
                                         'text_wrap': False,
                                         'valign': 'vcenter',
                                         'align': 'left',
                                         'fg_color': '#011E41',
                                         'font_color': '#FFFFFF',
                                         'border_color': '#FFFFFF',
                                         'border': 1})
    
    for df, ws in zip(ListDF, ListWS):
        ws.autofilter(0, 0, df.shape[0], df.shape[1])
        ws.set_column(0, 0, 15, cell_format)
        ws.set_column(1, 4, 24, cell_format)
        ws.set_column(5, df.shape[1], 50, cell_format)
        for col_num, value in enumerate(df.columns.values):
            ws.set_row(0, 25)
            ws.write(0, col_num + 1, value, header_format)
    #    for i, width in enumerate(get_col_widths(df)): # Autosize will not work because of the "\n" in the text
    #        ws.set_column(i, i, width, cell_format)

    # Define our range for the color formatting
    color_range1 = "E2:E{}".format(len(dfFINAL.index)+1)
    color_range2 = "F2:J{}".format(len(dfFINAL.index)+1)

    # Add a format. Light red fill with dark red text.
    format1 = workbook.add_format({'bg_color': '#FFC7CE',
                                'font_color': '#9C0006'})
    format2 = workbook.add_format({'bg_color': '#C6EFCE',
                                'font_color': '#006100'})
    format3 = workbook.add_format({'bg_color': '#FFFFFF',
                                'font_color': '#000000'})
    format4 = workbook.add_format({'bg_color': '#2385FC',
                                'font_color': '#FFFFFF'})
    
    # Highlight the values
    for i in ListFC:
        i.conditional_format(color_range1, {'type': 'duplicate',
                                            'format': format4})
        i.conditional_format(color_range2, {'type': 'blanks',
                                            'format': format3})
        i.conditional_format(color_range2, {'type': 'text',
                                            'criteria': 'containing',
                                            'value':    '[WRONG]',
                                            #'criteria': '=NOT(ISNUMBER(SEARCH($E2,F2)))',
                                            'format': format1})
        i.conditional_format(color_range2, {'type': 'text',
                                            'criteria': 'containing',
                                            'value':    '[OK]',
                                            'format': format2}) 

    text3 = 'Process Duration: ' + str((datetime.datetime.now() - now))
    worksheetS.write(3, 0, text3, bold)
        
    #workbook.close()
    writer.save()
        
    print('')
    print("Subprocess Duration: ", (datetime.datetime.now() - nowExcel))
    print('')
    print('##################################################')
    print('SUMMARY')
    print('##################################################')       
    print('')
    print(f'A total of {Tsensors} sensors files were processed and {Tfbf_fbz} SPL Sessions \n')
    print(dfSummary.to_markdown(tablefmt="github", index=False)) 
    # Empty Linename columns Log
    if not dfer.empty:
        print("")
        print(f"A total of {len(dfer)}/{Tfbf_fbz} Session SPL has/have no Linename information or Empty Session.")
        print(f"Please check the SPL_Problem sheet in the _{vessel}_FINAL_Log.xlsx for more information.")
        #dfer.to_csv(outputFolder + "\\" + vessel + "_SPL_Problem_log.csv", index=True)
    print('')
    print(f'Logs can be found in {outputFolder}.\n')
          
##########################################################
#                       Functions                        #
########################################################## 

##### Convert FBF/FBZ to CSV #####
def SPL2CSV(SPLFileName, Path):
    ##### Convert FBZ to CSV #####
    FileName = os.path.splitext(os.path.basename(SPLFileName))[0] 
    #print(SPLFileName)   
    SPLFilePath = Path + "\\" + FileName + '.txt'
    cmd = 'for %i in ("' + SPLFileName + '") do C:\ProgramData\Fugro\Starfix2018\Fugro.DescribedData2Ascii.exe -n3 %i Time LineName > "' + SPLFilePath + '"'
    #cmd = 'for %i in ("' + SPLFileName + '") do fbf2asc -n 3 -i %i Time LineName > "' + SPLFilePath + '"'  ## OLD TOOL   
    
    try:
        subprocess.call(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, close_fds=True) #https://github.com/pyinstaller/pyinstaller/wiki/Recipe-subprocess
        #subprocess.call(cmd, shell=True) ### For debugging
    except subprocess.CalledProcessError:
        os.remove(SPLFilePath)
        print('')        
        sys.exit(stylize(f'The following file is lock ({SPLFileName}).', fg('red')))

    # created the variables
    #dfS = pd.read_csv(SPLFilePath, header=None, skipinitialspace=True, usecols=[0,8], nrows=1)
    # from https://stackoverflow.com/questions/3346430/what-is-the-most-efficient-way-to-get-first-and-last-line-of-a-text-file/3346788
    if os.path.getsize(SPLFilePath) == 0:
        er = SPLFileName
        lnValue = "EmptySPL"
        SessionStart = np.datetime64('NaT')
        SessionEnd = np.datetime64('NaT')
        return SessionStart, SessionEnd, lnValue, er
        
    else:
        with open(SPLFilePath, 'rb') as fh:
            # 2020/10/17 03:15:24.420, "M1308"
            first = next(fh).decode()
            try:
                fh.seek(-200, os.SEEK_END) # 2 times size of the line
                last = fh.readlines()[-1].decode()
            except: # file less than 3 lines not working
                er = SPLFileName
                lnValue = "SPLtoSmall"
                SessionStart = first.split(',')[0] 
                SessionEnd = np.datetime64('NaT')
                return SessionStart, SessionEnd, lnValue, er

    LineName = first.split(',')[1].replace(' ', '', 1).replace('"', '').replace('\n', '').replace('\r', '')

    ## OLD TOOL
    #first = first.replace(' ', ',', 2).replace(',', ' ', 1)
    #last = last.replace(' ', ',', 2).replace(',', ' ', 1)
    #LineName = first.split(',')[1].replace('\n', '').replace(' \r', '')

    SessionStart = first.split(',')[0]    
    SessionEnd = last.split(',')[0]    

    #cleaning  
    os.remove(SPLFilePath)
    
    #checking if linename is empty as is use in all other process
    if not LineName: 
        er = SPLFileName
        lnValue = "NoLineNameFound"
        return SessionStart, SessionEnd, lnValue, er       
    else:
        er = ""
        lnValue = LineName
        return SessionStart, SessionEnd, lnValue, er

# SPL convertion
def splfc(splList, SPLFormat, dfSPL, dfer, outputFolder, cmd):
    """
    Reading SPL Files (FBF and FBZ)
    """
    print('')
    print('##################################################')
    print(f'READING {SPLFormat} FILES')
    print('##################################################')
    nowSPL = datetime.datetime.now() # record time of the subprocess
    pbar = tqdm(total=len(splList)) if cmd else print(f"Note: Output show file counting every {math.ceil(len(splList)/10)}") #cmd vs GUI 
            
    for index, n in enumerate(splList): 
        SessionStart, SessionEnd, LineName, er = SPL2CSV(n, outputFolder)        
        dfSPL = dfSPL.append(pd.Series([SessionStart, SessionEnd, LineName], 
                                index=dfSPL.columns ), ignore_index=True)        
        if er: dfer = dfer.append(pd.Series([er], index=dfer.columns ), ignore_index=True)        
        progressBar(cmd, pbar, index, splList)                   

    # Format datetime
    dfSPL['Session Start'] = pd.to_datetime(dfSPL['Session Start'], format='%Y/%m/%d %H:%M:%S.%f') # format='%d/%m/%Y %H:%M:%S.%f' format='%Y/%m/%d %H:%M:%S.%f' 
    dfSPL['Session End'] = pd.to_datetime(dfSPL['Session End'], format='%Y/%m/%d %H:%M:%S.%f')

    ## OLD TOOL
    #dfSPL['Session Start'] = pd.to_datetime(dfSPL['Session Start'], format='%d/%m/%Y %H:%M:%S.%f') # format='%d/%m/%Y %H:%M:%S.%f' format='%Y/%m/%d %H:%M:%S.%f' 
    #dfSPL['Session End'] = pd.to_datetime(dfSPL['Session End'], format='%d/%m/%Y %H:%M:%S.%f')

    pbar.close() if cmd else print("Subprocess Duration: ", (datetime.datetime.now() - nowSPL)) # cmd vs GUI
    
    return dfSPL, dfer
  
# Sensors convertion and logs creation
def sensorsfc(lsFile, ssFormat, ext, cmd, rename, outputFolder, dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, vessel):
    """
    Main function to create and rename the sensors files.    
    """
    print('')
    print('##################################################')
    print(f'READING *{ext} ({ssFormat}) FILES')
    print('##################################################')
    
    # Define Dataframe
    # Need to be declare fully in case of manipulated df (DO NOT CHANGE)
    dfSensors = pd.DataFrame(columns = ["Session Start", "Session End", "Vessel Name", "Sensor Start",
                                        "FilePath", "Sensor FileName", "SPL LineName", "Sensor New LineName"])
    dftmp = pd.DataFrame(columns = ["Session Start", "Session End", "Vessel Name", "Sensor Start",
                                        "FilePath", "Sensor FileName", "SPL LineName", "Sensor New LineName"])

    nowSensor = datetime.datetime.now()  # record time of the subprocess     
    pbar = tqdm(total=len(lsFile)) if cmd else print(f"Note: Output show file counting every {math.ceil(len(lsFile)/10)}") # cmd vs GUI
          
    # Reading the sensors files 
    for index, f in enumerate(lsFile):
        fName = os.path.splitext(os.path.basename(f))[0]        
        # MBES *.all Format https://github.com/pktrigg/pyall
        if ssFormat == 'MBES':
            r = ALLReader(f)          
            numberOfBytes, STX, typeOfDatagram, EMModel, RecordDate, RecordTime = r.readDatagramHeader() # read the common header for any datagram.
            fStart = to_DateTime(RecordDate, RecordTime)
            r.rewind()
            r.close() 
            
        # SSS *.xtf Files https://github.com/pktrigg/pyxtf
        if ssFormat == 'SSS':
            file_size = os.path.getsize(f)/(1024*1024)
            if file_size > 1: # skip file smaller than 1 MB
                r = XTFReader(f)            
                pingHdr = r.readPacket()
                if pingHdr != None:
                    fStart = datetime.datetime(pingHdr.Year, pingHdr.Month, pingHdr.Day, pingHdr.Hour, pingHdr.Minute, pingHdr.Second, pingHdr.HSeconds * 10000)
                    #FileNameinXTF = r.XTFFileHdr.ThisFileName                     
                r.rewind()
                r.close()
            else:
                dfSkip = dfSkip.append(pd.Series([f, file_size], index=dfSkip.columns ), ignore_index=True)
            
        # SBP/SUHRS *.sgy Files https://docs.obspy.org/master/packages/obspy.io.segy.html
        if ssFormat in ['SBP', 'SUHRS']:
            r = read(f, headonly=False)
            rStart = str(r[0].stats.starttime) # 2020-05-22T17:26:47.000000Z
            #print(f'START: {rStart}  END: {rEnd}')
            fStart = datetime.datetime.strptime(rStart, '%Y-%m-%dT%H:%M:%S.%fZ') # .split(' | ')[1].split(' - ')[0]
            #print(r[0])
            #print(r[150])          
        
        # SBP *.ses3 Files
        # Future if python convertion exist
               
        # MAG *.csv
        # Headers:
        #       Date,Time,Line_Name,X_Corrected,Y_Corrected,Towfish_Heading,SOG,STW,Altitude,
        #       Layback,Raw_Total_Field,Signal_Strength,Clean_Total_Field,Background_Field,Anomaly_Residual_Signal,
        #       Geological_Residual_Signal,Transverse_Gradient,Analytic_Signal
        if ssFormat == 'MAG':
            r = pd.read_csv(f, usecols=[0,1,2], nrows=1, parse_dates=[['Date', 'Time']])
            #print(r)
            fStart = r['Date_Time'].iloc[0]
             
        # Add the Sensor Info in a df
        #print(f'fStart {ssFormat} = {fStart}')
        dfSensors = dfSensors.append(pd.Series(["","", "", fStart, f, fName, "", ""], index=dfSensors.columns), ignore_index=True)        
        progressBar(cmd, pbar, index, lsFile)
                      
    pbar.close() if cmd else print("Subprocess Duration: ", (datetime.datetime.now() - nowSensor)) # cmd vs GUI

    # Format datetime
    dfSensors['Sensor Start'] = pd.to_datetime(dfSensors['Sensor Start'])  # format='%d/%m/%Y %H:%M:%S.%f' format='%Y/%m/%d %H:%M:%S.%f'
     
    print('')
    print(f'Listing and generating the rename list the *{ext} ({ssFormat}) files')
    nowListing = datetime.datetime.now()  # record time of the subprocess
    if rename == 'yes':
        pbar = tqdm(total=len(lsFile)) if cmd else print(f"Note: Output show file counting every {math.ceil(len(lsFile)/10)}") # cmd vs GUI   
    else:
        print("No file was rename. Option Rename was unselected")
        
    # Logs and renaming the Sensors files
    rcount = 0    
    for index, row in dfSPL.iterrows():                  
        splStart = row['Session Start']
        splEnd = row['Session End']
        splName = row['SPL LineName']   
        dffilter = dfSensors[dfSensors['Sensor Start'].between(splStart, splEnd)]        
        for index, el in dffilter.iterrows():
            #print(el)
            SensorFile =  el['FilePath']
            SensorStart = el['Sensor Start']
            FolderName = os.path.split(SensorFile)[0]
            SensorName = os.path.splitext(os.path.basename(SensorFile))[0]
            if splName in SensorName: # use for conditional formating and because I group linename under the same Session some sensor can contain or not the SPL linename
                SNameCond = SensorName + ' [OK]'
            else:
                SNameCond = SensorName + ' [WRONG]'
            SensorExt =  os.path.splitext(os.path.basename(SensorFile))[1]                              
            SensorNewName = FolderName + '\\' + SensorName + '_' + splName + SensorExt
            dftmp = dftmp.append(pd.Series([splStart, splEnd, vessel, SensorStart, SensorFile, SNameCond, splName, SensorNewName], 
                                index=dftmp.columns), ignore_index=True)
            
            
            # rename the sensor file           
            if rename == 'yes':
                if cmd:
                    os.rename(SensorFile, SensorNewName)
                    rcount += 1
                    pbar.update(1)
                else:
                    os.rename(SensorFile, SensorNewName)
                    rcount += 1
                    print_progress(index, len(lsFile)) # to have a nice progress bar in the GUI      
                    if index % math.ceil(len(lsFile)/10) == 0: # decimate print
                        print(f"Files Process: {index+1}/{len(lsFile)}")                                
    
    if rename == 'yes':      
        pbar.close() if cmd else print("Subprocess Duration: ", (datetime.datetime.now() - nowListing)) # cmd vs GUI
       
    print('')
    print('Creating logs. Please wait....')
    # Format datetime
    dfSensors['Session Start'] = pd.to_datetime(dfSensors['Session Start'], unit='s')
    dfSensors['Session End'] = pd.to_datetime(dfSensors['Session End'], unit='s')
    
    # Timestamps error in SGY file Log
    if ssFormat == 'SBP':
        date_before = max(dfSensors['Sensor Start']) - datetime.timedelta(days=2*365)
        dfsgy = dfSensors[dfSensors['Sensor Start'] < date_before]
        if not dfsgy.empty:
            print("")
            print(f"A total of {len(dfsgy)} *{ext} ({ssFormat}) has/have wrong Start Time information.")
            print(f"Please check the {vessel}_{ssFormat}_WrongStartTime_log.csv for more information.")
            dfsgy.to_csv(outputFolder + "\\" + vessel + "_" + ssFormat + "_WrongStartTime_log.csv", index=True,
                         columns=["Sensor Start", "Vessel Name", "FilePath", "Sensor FileName"])  
       
    # Droping duplicated and creating a log.
    dfCountDupl = dftmp[dftmp.duplicated(subset='Sensor Start', keep=False)].sort_values('Sensor Start')
    dfDuplSensor = dfDuplSensor.append(dftmp[dftmp.duplicated(subset='Sensor Start', keep=False)].sort_values('Sensor Start'))
    if not dfCountDupl.empty:
        print("")
        print(f"A total of {int(len(dfCountDupl.index)/2)} *{ext} file(s) was/were duplicated.")
        print(f"Please check the Duplicated_Sensor_Data sheet in the _{vessel}_FINAL_Log.xlsx for more information.")
        #dfDuplSensor.to_csv(outputFolder + "\\" + vessel + "_" + ssFormat + "_Duplicate_Log.csv", index=True)
        dftmp = dftmp.drop_duplicates(subset='Sensor Start', keep='first')
        dfSensors = dfSensors.drop_duplicates(subset='Sensor Start', keep='first')
    
    # Updated the final dataframe to export the log
    dfSensors.set_index('Sensor Start', inplace=True)
    dftmp.set_index('Sensor Start', inplace=True)    
    dfSensors.update(dftmp)    
    
    # Reset index if not the column Sensor Start will be missing
    dfSensors = dfSensors.reset_index()
    
    # Format datetime
    dfSensors['Session Start'] = pd.to_datetime(dfSensors['Session Start'], unit='s')
    dfSensors['Session End'] = pd.to_datetime(dfSensors['Session End'], unit='s')

    # Saving the Full log
    dfSensors.to_csv(outputFolder + "\\" + vessel + "_" + ssFormat + "_Full_Log.csv", index=False)
    
    # Creating the FINAL df
    dfMissingSPL = dfMissingSPL.append(dfSensors)
    dfMissingSPL['Sensor Type'] = ssFormat

    dfFINAL.set_index('Session Start', inplace=True)
    dfSensors = dfSensors.dropna(subset=['Session Start']) # Drop Missing SPL
    dfSensors.rename(columns={"Sensor FileName": ssFormat}, inplace = True)
    
    dfSensors = dfSensors.groupby('Session Start')[ssFormat].apply(lambda x: '\n'.join(x)).to_frame().reset_index()
    dfSensors.set_index('Session Start', inplace=True)
    dfFINAL.update(dfSensors)
    
    dfFINAL = dfFINAL.reset_index() # Reset index if not the column Session Start will be missing
    #print(dfFINAL.to_markdown(tablefmt="github", index=False))
    
    # Add the Sumary Info in a df
    dfSummary = dfSummary.append(pd.Series([ssFormat, len(lsFile), int(len(dfCountDupl.index)/2), len(dfsgy) if ssFormat == 'SBP' else np.nan, rcount], 
                index=dfSummary.columns ), ignore_index=True) 
    
    return dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip

# Move MAG and SUHRS function
def mvSensorFile (lsFile, VesselFolder, WrongFolder, cmd, ssFormat, dfSummary):
    """
    Moving MAG and SUHRS in the vessel folder or WRONG folder based on SPL Files
    """
    print('')
    print('##################################################')
    print(f'MOVING {ssFormat} FILES')
    print('##################################################')
    nowmove = datetime.datetime.now() # record time of the subprocess
           
    dfmove = pd.read_csv(lsFile, usecols=[3,4,5])
    dfmove = dfmove[~dfmove['Vessel Name'].isnull()]
    
    dfOK = dfmove[dfmove.apply(lambda row: '[OK]' in str(row['Sensor FileName']), axis=1)]
    dfWRONG = dfmove[dfmove.apply(lambda row: '[WRONG]' in str(row['Sensor FileName']), axis=1)]
    lsOK = dfOK['FilePath'].tolist()
    lsWRONG = dfWRONG['FilePath'].tolist() 
    
    count = len(lsOK) + len(dfWRONG)
          
    if not os.path.exists(VesselFolder):
        os.mkdir(VesselFolder)
    if not os.path.exists(WrongFolder):
        os.mkdir(WrongFolder)
    
    pbar = tqdm(total=count) if cmd else print(f"Moving the OK files to {VesselFolder}.\nNote: Output show file counting every {math.ceil(count/10)}") #cmd vs GUI 
    
    i = 0
    for f in lsOK:
        i += 1
        newpath = os.path.join(VesselFolder, os.path.basename(f))            
        shutil.move(f, newpath)
        progressBar(cmd, pbar, i, lsOK)    
    for f in lsWRONG:
        newpath = newpath = os.path.join(WrongFolder, os.path.basename(f))
        shutil.move(f, newpath)
        
    # Add the Sumary Info in a df
    #df.at[index, col] = val https://stackoverflow.com/questions/13842088/set-value-for-particular-cell-in-pandas-dataframe-using-index
    dfSummary.set_index('Sensor', inplace=True)
    dfSummary.at[ssFormat, 'Moved Files'] = count
    dfSummary = dfSummary.reset_index() # Reset index if not the column Session Start will be missing              
    pbar.close() if cmd else print("Subprocess Duration: ", (datetime.datetime.now() - nowmove)) # cmd vs GUI
    
    return dfSummary
    

# Others function
# from https://www.pakstech.com/blog/python-gooey/
def print_progress(index, total):
    print(f"progress: {index+1}/{total}")
    sys.stdout.flush()
    
# Progrees bar GUI and CMD
def progressBar(cmd, pbar, index, ls):
    if cmd:
        pbar.update(1)
    else:
        print_progress(index, len(ls)) # to have a nice progress bar in the GU            
        if index % math.ceil(len(ls)/10) == 0: # decimate print
            print(f"Files Process: {index+1}/{len(ls)}") 

# List file in subfolder with exclude
def listFile(Folder, ext, exclude):
    ls = []
    for root, dirs, files in os.walk(Folder, topdown=True):
        dirs[:] = [d for d in dirs if d not in exclude]
        for filename in files:                
            if filename.lower().endswith(ext): #call lower to make the string lowercase before calling endswith
                filepath = os.path.join(root, filename)
                ls.append(filepath)
    print(f'Start Listing for {Folder}')
    return ls

# Simulate autofit column in xslxwriter https://stackoverflow.com/questions/29463274/simulate-autofit-column-in-xslxwriter
def get_col_widths(dataframe):
    # First we find the maximum length of the index column   
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]



##########################################################
#                        __main__                        #
########################################################## 
if __name__ == "__main__":
    now = datetime.datetime.now() # time the process
    # Preparing your script for packaging https://chriskiehl.com/article/packaging-gooey-with-pyinstaller
    # Prevent stdout buffering     
    #nonbuffered_stdout = os.fdopen(sys.stdout.fileno(), 'w') #https://stackoverflow.com/questions/45263064/how-can-i-fix-this-valueerror-cant-have-unbuffered-text-i-o-in-python-3/45263101
    #sys.stdout = nonbuffered_stdout
    main()
    print('')
    print("Process Duration: ", (datetime.datetime.now() - now)) # print the processing time. It is handy to keep an eye on processing performance.