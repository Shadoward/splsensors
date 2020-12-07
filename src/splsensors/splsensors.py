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
from xlsxwriter.utility import xl_rowcol_to_cell

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
    default_size=(930, 770),
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
                'version': '0.3.0',
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
    sensorsopt = parser.add_argument_group('Sensors Options (First Run)', description='Leave the field blank if you do not need to process the sensor', gooey_options={'columns': 1})
    sensor2sopt = parser.add_argument_group('Sensors Options (Others Runs)', description='This option can be use to speed up the creation of the final list.\nUse this option if you do not need to re-read the sensor files\nLeave the field blank if you do not need to process the sensor\nPLEASE THE FILES SHOULD BE IN A OTHER FOLDER THAT THE OUTPUT FOLDER SELECTED IN THE TOOL!!!!!', gooey_options={'columns': 1})
    outputsopt = parser.add_argument_group('Output Options', gooey_options={'columns': 1})
    additionalopt = parser.add_argument_group('Additional Options', gooey_options={'columns': 1})
    renameopt = parser.add_argument_group('Rename Options')
     
    # SPL Arguments
    splopt.add_argument(
        '-i',
        '--splFolder', 
        dest='splFolder',       
        metavar='SPL Root Path', 
        help='This is the path where the *.fbf/*.fbz/*.pos files to process are. (Root Session Folder)',
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\NEL',
        #default_path='S:\\JOBS\\2020\\20030002_Shell_FBR_MF\\B2B_FromVessel\\Navigation\\Starfix_Logging\\RawData', 
        widget='DirChooser')
        
    splopt.add_argument(
        '-p', '--splPosition', 
        dest='splPosition',
        metavar='SPL Position File Name', 
        widget='TextField',
        #default='FugroBrasilis-CRP-Position',
        help='SPL position file to be use to rename the sensor without extention.')
    
    splopt.add_argument(
        '-b', '--buffer',
        dest='buffer',
        metavar='Start Buffer [s]', 
        widget='TextField',
        #default='FugroBrasilis-CRP-Position',
        help='Start buffer [in second] to be used to included sensors that have start before the session start.')
        
    # Sensors Arguments
    sensorsopt.add_argument(
        '-A', '--allFolder', 
        dest='allFolder',        
        metavar='ALL Folder Path',
        help='ALL Root path. This is the path where the *.all files to process are.',
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\ALL', 
        widget='DirChooser')
        
    sensorsopt.add_argument(
        '-X', '--xtfFolder', 
        dest='xtfFolder',        
        metavar='XTF Folder Path',
        help='XTF Root path. This is the path where the *.xtf files to process are.',
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\XTF', 
        widget='DirChooser')
        
    sensorsopt.add_argument(
        '-S', '--sgySBPFolder', 
        dest='sgySBPFolder',        
        metavar='SGY/SEG/SEGY SBP Folder Path',
        help='SGY/SEG/SEGY SBP Root path. This is the path where the *.sgy/*.seg/*.segy files to process are.',
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\SGYSBP', 
        widget='DirChooser')

    sensorsopt.add_argument(
        '-M', '--csvMAGFolder', 
        dest='csvMAGFolder',        
        metavar='CSV MAG Folder Path',
        help='CSV MAG Root path. This is the path where the *.csv files to process are.',
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\CSVMAG', 
        widget='DirChooser')
        
    sensorsopt.add_argument(
        '-H', '--sgySUHRSFolder', 
        dest='sgySUHRSFolder',        
        metavar='SGY/SEG/SEGY SUHRS Folder Path',
        help='SGY/SEG/SEGY SUHRS Root path. This is the path where the *.sgy/*.seg/*.segy files to process are.',
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\SGYSUHRS', 
        widget='DirChooser')        
        
    # Sensors Arguments Other runs
    sensor2sopt.add_argument(
        '-A2', '--allFile', 
        dest='allFile',        
        metavar='*_MBES_Full_Log.csv File Path',
        help='This is the file that list all MBES and it start time generated by this tool on the first run.',
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\ALL', 
        widget='FileChooser',
        gooey_options={'wildcard': "Comma separated file (*.csv)|*_MBES_Full_Log.csv"})
        
    sensor2sopt.add_argument(
        '-X2', '--xtfFile', 
        dest='xtfFile',        
        metavar='*_SSS_Full_Log.csv File Path',
        help='This is the file that list all SSS and it start time generated by this tool on the first run.',
        widget='FileChooser',
        gooey_options={'wildcard': "Comma separated file (*.csv)|*_SSS_Full_Log.csv"})
        
    sensor2sopt.add_argument(
        '-S2', '--sgySBPFile', 
        dest='sgySBPFile',        
        metavar='*_SBP_Full_Log.csv File Path',
        help='This is the file that list all SBP and it start time generated by this tool on the first run.',
        widget='FileChooser',
        gooey_options={'wildcard': "Comma separated file (*.csv)|*_SBP_Full_Log.csv"})
        
    sensor2sopt.add_argument(
        '-M2', '--csvMAGFile', 
        dest='csvMAGFile',        
        metavar='*_MAG_Full_Log.csv File Path',
        help='This is the file that list all MAG and it start time generated by this tool on the first run.',
        widget='FileChooser',
        gooey_options={'wildcard': "Comma separated file (*.csv)|*_MAG_Full_Log.csv"})
        
    sensor2sopt.add_argument(
        '-H2', '--sgySUHRSFile', 
        dest='sgySUHRSFile',        
        metavar='*_SUHRS_Full_Log.csv File Path',
        help='This is the file that list all SUHRS and it start time generated by this tool on the first run.',
        widget='FileChooser',
        gooey_options={'wildcard': "Comma separated file (*.csv)|*_SUHRS_Full_Log.csv"})
        
    # Output Arguments
    outputsopt.add_argument(
        '-o', '--output',
        dest='outputFolder',
        metavar='Output Logs Folder',  
        help='Output folder to save all the logs files.',      
        #type=str,
        #default_path='C:\\Users\\patrice.ponchant\\Downloads\\LOGS',
        widget='DirChooser')
    
    # Additional Arguments
    additionalopt.add_argument(
        '-m', '--move',
        dest='move',
        metavar='Move MAG and SUHRS in the vessel folder?', 
        help='This will create and vessel folder in the sensor folder basaed on the SPL name vessel and move the files to this.',
        choices=['yes', 'no'], 
        default='yes')
    additionalopt.add_argument(
        '-e', '--excludeFolder', 
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
    splPosition = str(args.splPosition)
    vessel = splPosition.split('-')[0]
    buffer = args.buffer if args.buffer is not None else 0
    
    allFolder = args.allFolder
    xtfFolder = args.xtfFolder
    sgySBPFolder = args.sgySBPFolder
    csvMAGFolder = args.csvMAGFolder
    sgySUHRSFolder = args.sgySUHRSFolder
    
    allFile = args.allFile
    xtfFile = args.xtfFile
    sgySBPFile = args.sgySBPFile
    csvMAGFile = args.csvMAGFile
    sgySUHRSFile = args.sgySUHRSFile
    
    outputFolder = args.outputFolder
    
    excludeFolder = args.excludeFolder
    move = args.move
    
    # Defined Global Dataframe
    col = ["Session Start", "Session End", "Session Name", "Session MaxGap", "Vessel Name", "Sensor Start",
            "FilePath", "Sensor FileName", "SPL LineName", "Sensor New LineName"]
    dfFINAL = pd.DataFrame(columns = ["Session Start", "Session End", "Session Name", "Session MaxGap", "Vessel Name", 
                                      "SPL", "MBES", "SBP", "SSS", "MAG", "SUHRS"])
    dfSPL = pd.DataFrame(columns = ["Session Start", "Session End", "SPL LineName", "Session MaxGap", "Session Name"])       
    dfer = pd.DataFrame(columns = ["SPLPath"])
    dfSummary = pd.DataFrame(columns = ["Sensor", "Processed Files", "Duplicated Files", "Wrong Timestamp (SBP)",
                                        "Moved Files", "Processing Time"])
    dfMissingSPL = pd.DataFrame(columns = col)
    dfDuplSensor = pd.DataFrame(columns = col)
    dfsgy = pd.DataFrame(columns = col)
    dfSkip = pd.DataFrame(columns = ["FilePath", "File Size [MB]"])
    
    ##########################################################
    #              Checking before continuing                #
    ##########################################################   
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
    
    # Check if LOGS files used for listing is not in same folder that LOGS that will be logs 
    tpmList = [allFile, xtfFile, sgySBPFile, csvMAGFile, sgySUHRSFile]
    for e in tpmList:
        if e is not None:
            if os.path.dirname(os.path.abspath(e)) == outputFolder:
                print ('')
                sys.exit(stylize('The Output Folder is the same as the *_Full_Log_.csv, quitting', fg('red')))
    
    ##########################################################
    #                 Listing the files                      #
    ##########################################################  
    print('')
    print('##################################################')
    print('LISTING FILES. PLEASE WAIT....')
    print('##################################################')
    nowLS = datetime.datetime.now() # record time of the subprocess
    
    exclude = excludeFolder.split(",") if args.excludeFolder is not None else []
    print(f'The following folders(s) will be excluded: {exclude} \n') if args.excludeFolder is not None else []
    
    if args.allFile is not None:
        allListFile = pd.read_csv(allFile, usecols=[0,4], parse_dates=['Sensor Start'])
    elif args.allFolder is not None:
        allListFile = listFile(allFolder, "all", set(exclude))
    else:
        allListFile = []

    if args.xtfFile is not None:
        xtfListFile = pd.read_csv(xtfFile, usecols=[0,4], parse_dates=['Sensor Start'])
    elif args.xtfFolder is not None:
        xtfListFile = listFile(xtfFolder, "xtf", set(exclude))
    else:
        xtfListFile = []

    if args.sgySBPFile is not None:
        sbpListFile = pd.read_csv(sgySBPFile, usecols=[0,4], parse_dates=['Sensor Start'])
    elif args.sgySBPFolder is not None:
        sbpListFile = listFile(sgySBPFolder, ('seg', 'sgy', 'segy'), set(exclude))
    else:
        sbpListFile = []

    if args.csvMAGFile is not None:
        magListFile = pd.read_csv(csvMAGFile, usecols=[0,4], parse_dates=['Sensor Start'])
    elif args.csvMAGFolder is not None:
        magListFile = listFile(csvMAGFolder, "csv", set(exclude))
    else:
        magListFile = []        

    if args.sgySUHRSFile is not None:
        suhrsListFile = pd.read_csv(sgySUHRSFile, usecols=[0,4], parse_dates=['Sensor Start'])
    elif args.sgySUHRSFolder is not None:
        suhrsListFile = listFile(sgySUHRSFolder, ('seg', 'sgy', 'segy'), set(exclude))
    else:
        suhrsListFile = []
    
    #### Not very efficient, will need to improve later   
    print(f'Start Listing for {splFolder}')
    fbzListFile = glob.glob(splFolder + "\\**\\" + splPosition + ".fbz", recursive=True)
    fbfListFile = glob.glob(splFolder + "\\**\\" + splPosition + ".fbf", recursive=True)
    posListFile = glob.glob(splFolder + "\\**\\" + splPosition + ".pos", recursive=True)           
    print("Subprocess Duration: ", (datetime.datetime.now() - nowLS))
    
    # Check if SPL files found
    if not fbfListFile and not fbzListFile and not posListFile:
        print ('')
        sys.exit(stylize('No SPL files were found, quitting', fg('red')))
    
    print('')
    print(f'Total of files that will be processed: \n {len(fbfListFile)} *.fbf \n {len(fbzListFile)} *.fbz \n {len(posListFile)} *.pos \n {len(allListFile)} *.all \n {len(xtfListFile)} *.xtf \n {len(sbpListFile)} *.sgy/*.seg/*.segy (SBP) \n {len(magListFile)} *.csv (MAG) \n {len(suhrsListFile)} *.sgy/*.seg/*.segy (SUHRS)') 
    
    ##########################################################
    #                     Reading SPL                        #
    ##########################################################    
    dfSPL, dfer, dfSummary = splfc(fbfListFile, 'FBF', dfSPL, dfer, dfSummary, outputFolder, cmd)
    dfSPL, dfer, dfSummary = splfc(fbzListFile, 'FBZ', dfSPL, dfer, dfSummary, outputFolder, cmd)
    dfSPL, dfer, dfSummary = splfc(posListFile, 'POS', dfSPL, dfer, dfSummary, outputFolder, cmd)
    
    # Copy the needed info from dfSPl to the dfFINAL
    dfFINAL['Session Start'] = dfSPL['Session Start']
    dfFINAL['Session End'] = dfSPL['Session End']
    dfFINAL['SPL'] = dfSPL['SPL LineName']
    dfFINAL['Session MaxGap'] = dfSPL['Session MaxGap']
    dfFINAL['Session Name'] = dfSPL['Session Name']
    
    ##########################################################
    #                    Reading Sensors                     #
    ##########################################################
    # MBES
    if args.allFile is not None:
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy = sensorsfc('File', allListFile, 'MBES', '.all', cmd, buffer, outputFolder,
                                                                            dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy, vessel)
    elif args.allFolder is not None:
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy = sensorsfc('Folder', allListFile, 'MBES', '.all', cmd, buffer, outputFolder,
                                                                            dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy, vessel)
    # XTF
    if args.xtfFile is not None:
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy = sensorsfc('File', xtfListFile, 'SSS', '.xtf', cmd, buffer, outputFolder,
                                                                            dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy, vessel)
    elif args.xtfFolder is not None:
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy = sensorsfc('Folder', xtfListFile, 'SSS', '.xtf', cmd, buffer, outputFolder,
                                                                            dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy, vessel)
    # SBP
    if args.sgySBPFile is not None:
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy = sensorsfc('File', sbpListFile, 'SBP', '.sgy/*.seg/*.segy', cmd, buffer, outputFolder,
                                                                            dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy, vessel)
    elif args.sgySBPFolder is not None:
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy = sensorsfc('Folder', sbpListFile, 'SBP', '.sgy/*.seg/*.segy', cmd, buffer, outputFolder,
                                                                            dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy, vessel)
    # MAG
    if args.csvMAGFile is not None:
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy = sensorsfc('File', magListFile, 'MAG', '.csv', cmd, buffer, outputFolder,
                                                                            dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy, vessel)
        if move == 'yes':
            lsFile = outputFolder + "\\" + vessel + "_MAG_Full_Log.csv"
            VesselFolder = os.path.join(csvMAGFolder, vessel)
            WrongFolder = os.path.join(csvMAGFolder, 'WRONG')
            dfSummary = mvSensorFile(lsFile, VesselFolder, WrongFolder, cmd, 'MAG', dfSummary) 
    elif args.csvMAGFolder is not None:
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy = sensorsfc('Folder', magListFile, 'MAG', '.csv', cmd, buffer, outputFolder,
                                                                   dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy, vessel)
        if move == 'yes':
            lsFile = outputFolder + "\\" + vessel + "_MAG_Full_Log.csv"
            VesselFolder = os.path.join(csvMAGFolder, vessel)
            WrongFolder = os.path.join(csvMAGFolder, 'WRONG')
            dfSummary = mvSensorFile(lsFile, VesselFolder, WrongFolder, cmd, 'MAG', dfSummary) 
    # SUHRS
    if args.sgySUHRSFile is not None:
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy = sensorsfc('File', suhrsListFile, 'SUHRS', '.sgy/*.seg/*.segy', cmd, buffer, outputFolder,
                                                                   dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy, vessel)
        if move == 'yes':
            lsFile = outputFolder + "\\" + vessel + "_SUHRS_Full_Log.csv"
            VesselFolder = os.path.join(sgySUHRSFolder, vessel)
            WrongFolder = os.path.join(sgySUHRSFolder, 'WRONG')
            dfSummary = mvSensorFile(lsFile, VesselFolder, WrongFolder, cmd, 'SUHRS', dfSummary)         
    elif args.sgySUHRSFolder is not None:
        dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy = sensorsfc('Folder', suhrsListFile, 'SUHRS', '.sgy/*.seg/*.segy', cmd, buffer, outputFolder,
                                                                   dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy, vessel)
        if move == 'yes':
            lsFile = outputFolder + "\\" + vessel + "_SUHRS_Full_Log.csv"
            VesselFolder = os.path.join(sgySUHRSFolder, vessel)
            WrongFolder = os.path.join(sgySUHRSFolder, 'WRONG')
            dfSummary = mvSensorFile(lsFile, VesselFolder, WrongFolder, cmd, 'SUHRS', dfSummary) 


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
    log_filenames = glob.glob(outputFolder + '\\*_Full_Log.csv')
    dfALL = pd.concat([pd.read_csv(f) for f in log_filenames])
    dtFormat = ['Sensor Start', 'Session Start', 'Session End']
    for dt in dtFormat:
        dfALL[dt] = pd.to_datetime(dfALL[dt])
    dfALL['Difference Start [s]'] = dfALL['Session Start'] - dfALL['Sensor Start']
    dfALL['Difference Start [s]'] = dfALL['Difference Start [s]'] / np.timedelta64(1, 's')
    dfALL = movecol(dfALL, cols_to_move=['Difference Start [s]'], ref_col='Sensor Start', place='After')
    
    coldrop =["Session Start", "Session End", "Session Name", "Session MaxGap", "Vessel Name", "SPL LineName", "Sensor New LineName"]
    dfMissingSPL = dfMissingSPL[dfMissingSPL['Session Start'].isnull()] # df Missing SPL
    dfMissingSPL = dfMissingSPL.drop(columns=coldrop)
    dfMissingSPL = dfMissingSPL[["Sensor Start", "Sensor FileName", "FilePath"]]
    
    dfsgy = dfsgy.drop(columns=coldrop)
    dfsgy = dfsgy[["Sensor Start", "Sensor FileName", "FilePath"]]
    
    dfSPLProblem = dfFINAL[dfFINAL['SPL'].isin(['NoLineNameFound', 'EmptySPL', 'SPLtoSmall'])]
    dfSPLProblem = dfSPLProblem.drop_duplicates(subset='Session Start', keep='first')
    
    lsWRONG = ["MBES", "SSS", "SBP", "MAG", "SUHRS"]    
    d = {name: pd.DataFrame() for name in lsWRONG}    
    for name, df in d.items():
        d[name] = dfFINAL[dfFINAL.apply(lambda row: '[WRONG]' in str(row[name]), axis=1)].sort_values('Session Start')
        d[name] = d[name].append(dfFINAL[dfFINAL[name].isnull()].sort_values('Session Start')) #https://stackoverflow.com/questions/29314033/drop-rows-containing-empty-cells-from-a-pandas-dataframe/56708633#56708633
        d[name] = movecol(d[name], cols_to_move=[name], ref_col='SPL', place='After')
    
    dfDuplSPL = dfFINAL[dfFINAL.duplicated(subset='SPL', keep=False)]
    #print(dfFINAL.apply(lambda row: str(row.MBES) in '[WRONG]', axis=1))
    
    writer = pd.ExcelWriter(outputFolder + "\\_" + vessel + '_FINAL_Log.xlsx', engine='xlsxwriter', datetime_format='dd mmm yyyy hh:mm:ss.000')
      
    sheet_names = ['Summary_Process_Log', 'Full_List', 'List_Transposed', 'Missing_SPL', 'MBES_NotMatching', 'SSS_NotMatching',
                   'SBP_NotMatching', 'MAG_NotMatching', 'SUHRS_NotMatching', 'Duplicated_SPL_Name', 'Duplicated_Sensor_Data',
                   'SPL_Problem', 'Skip_SSS_Files', 'Wrong_SBP_Time']
              
    dfSummary.to_excel(writer, sheet_name='Summary_Process_Log', startrow=5)
    
    dfALL.sort_values('Sensor Start').to_excel(writer, sheet_name='Full_List')    
    dfFINAL.sort_values('Session Start').to_excel(writer, sheet_name='List_Transposed') 
    dfMissingSPL.sort_values('Sensor Start').to_excel(writer, sheet_name='Missing_SPL')
    
    for name, df in d.items():
        d[name].to_excel(writer, sheet_name = name +'_NotMatching')
        
    dfDuplSPL.sort_values('SPL').to_excel(writer, sheet_name='Duplicated_SPL_Name')
    dfDuplSensor.sort_values('Sensor Start').to_excel(writer, sheet_name='Duplicated_Sensor_Data')    
    dfSPLProblem.sort_values('Session Start').to_excel(writer, sheet_name='SPL_Problem')
    dfSkip.to_excel(writer, sheet_name='Skip_SSS_Files')
    dfsgy.to_excel(writer, sheet_name='Wrong_SBP_Time')
    
    workbook  = writer.book    

    # worksheet list to format
    w = {name: writer.sheets[name] for name in sheet_names}
    w['Summary_Process_Log'].hide_gridlines(2)
    
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
    
    cell_format = workbook.add_format({'text_wrap': True,
                                       'font_name': 'Segoe UI',
                                       'font_size': 10,
                                       'valign': 'vcenter',
                                       'align': 'left',
                                       'border_color': '#000000',
                                       'border': 1})
    
    session_format = workbook.add_format({'num_format': '0',
                                          'text_wrap': True,
                                          'font_name': 'Segoe UI',
                                          'font_size': 10,
                                          'valign': 'vcenter',
                                          'align': 'left',
                                          'border_color': '#000000',
                                          'border': 1})
    
    cell_time = workbook.add_format({'num_format': 'hh:mm:ss.000',
                                    'text_wrap': False,
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
    
  
    for i, width in enumerate(get_col_widths(dfSummary)):
        w['Summary_Process_Log'].set_column(i, i, width)
        
    text1 = 'A total of ' + str(Tsensors) + ' sensors files were processed and ' + str(Tfbf_fbz) + ' SPL Sessions.'
    text2 = 'A total of ' + str(len(dfer)) + '/' + str(Tfbf_fbz) + ' Session SPL has/have no Linename information.' 
    textS = [bold, 'Summary_Process_Log', normal, ': Summary log of the processing']
    textFull = [bold, 'Full_List', normal, ': Full log list of all sensors without duplicated and skip files. (Sensors Not Transposed)']
    textTrans = [bold, 'List_Transposed', normal, ': Log list of all sensors transposed and matching all sessions)']
    textMissingSPL = [bold, 'Missing_SPL', normal, ': List of all sensors that have missing SPL file.']
    textMBES = [bold, 'MBES_NotMatching', normal, ': MBES log list of all files that do not match the SPL name; without duplicated and skip files']
    textSSS = [bold, 'SSS_NotMatching', normal, ': SSS log list of all files that do not match the SPL name; without duplicated and skip files']
    textSBP = [bold, 'SBP_NotMatching', normal, ': SBP log list of all files that do not match the SPL name; without duplicated and skip files']
    textMAG = [bold, 'MAG_NotMatching', normal, ': MAG log list of all files that do not match the SPL name; without duplicated and skip files']
    textSUHRS = [bold, 'SUHRS_NotMatching', normal, ': SUHRS log list of all files that do not match the SPL name; without duplicated and skip files']
    textDuplSPL = [bold, 'Duplicated_SPL_Name', normal, ': List of all duplicated SPL name']
    textDuplSensor = [bold, 'Duplicated_Sensor_Data', normal, ': List of all duplicated sensors files; Based on the start time']
    textSPLProblem = [bold, 'SPL_Problem', normal, ': List of all SPL session without a line name in the columns LineName, are empty or too small']
    textSkip = [bold, 'Skip_SSS_Files', normal, ': List of all SSS data that have a file size less than 1 MB']
    textsgy = [bold, 'Wrong_SBP_Time', normal, ': List of all SBP data that have a wrong timestamp']
    
    ListT = [textS, textFull, textTrans, textMissingSPL, textMBES, textSSS, textSBP, textMAG, textSUHRS, textDuplSPL, textDuplSensor, 
             textSPLProblem, textSkip, textsgy]
    ListHL = ['internal:Summary_Process_Log!A1', 'internal:Full_List!A1', 'internal:List_Transposed!A1', 'internal:Missing_SPL!A1', 
              'internal:MBES_NotMatching!A1', 'internal:SSS_NotMatching!A1', 'internal:SBP_NotMatching!A1', 'internal:MAG_NotMatching!A1',
              'internal:SUHRS_NotMatching!A1','internal:Duplicated_SPL_Name!A1', 'internal:Duplicated_Sensor_Data!A1', 
              'internal:SPL_Problem!A1', 'internal:Skip_SSS_Files!A1', 'internal:Wrong_SBP_Time!A1']
                
    w['Summary_Process_Log'].write(0, 0, text1, bold)
    w['Summary_Process_Log'].write(1, 0, text2, bold)
    
    icount = 0
    for e, l in zip(ListT,ListHL):
        w['Summary_Process_Log'].write_rich_string(dfSummary.shape[0] + 7 + icount, 1, *e)
        w['Summary_Process_Log'].write_url(dfSummary.shape[0] + 7 + icount, 0, l, hlink, string='Link')
        icount += 1
    
    range_table = "B7:G{}".format(dfSummary.shape[0]+6)
    range_time = "G7:G{}".format(dfSummary.shape[0]+6)
    w['Summary_Process_Log'].conditional_format(range_time, {'type': 'no_blanks',
                                               'format': cell_time})
    w['Summary_Process_Log'].conditional_format(range_table, {'type': 'no_blanks',
                                                'format': cell_format})

    # Others Sheet
    for name, ws in w.items():
        if name != 'Summary_Process_Log':
            ws.write_url(0, 0, 'internal:Summary_Process_Log!A1', hlink, string='Summary')
    
    ListDF = [dfSummary, dfALL, dfFINAL, dfMissingSPL, d['MBES'], d['SSS'], d['SBP'], d['MAG'], d['SUHRS'], dfDuplSPL, dfDuplSensor, 
              dfSPLProblem, dfSkip, dfsgy]
    
    for df, (namews, ws) in zip(ListDF, w.items()):
        if namews != 'Summary_Process_Log':
            list1 = ['List_Transposed', 'Missing_SPL', 'MBES_NotMatching', 'SSS_NotMatching', 'SBP_NotMatching', 'MAG_NotMatching', 'SUHRS_NotMatching']
            list2 = ['Missing_SPL', 'Skip_SSS_Files', 'Duplicated_Sensor_Data', 'Wrong_SBP_Time']
            for col_num, value in enumerate(df.columns.values):
                ws.set_row(0, 25)
                ws.write(0, col_num + 1, value, header_format)                
            if namews == 'Full_List':                
                ws.autofilter(0, 0, df.shape[0], df.shape[1])
                ws.set_column(0, 0, 11, cell_format) # ID
                ws.set_column(df.columns.get_loc('Sensor Start')+1, df.columns.get_loc('Session End')+1, 24, cell_format) # DateTime
                ws.set_column(df.columns.get_loc('Session Name')+1, df.columns.get_loc('Session Name')+1, 20, session_format) # Session Name
                ws.set_column(df.columns.get_loc('Session MaxGap')+1, df.columns.get_loc('Vessel Name')+1, 20, cell_format) # Session Info
                ws.set_column(df.columns.get_loc('FilePath')+1, df.columns.get_loc('FilePath')+1, 150, cell_format)
                ws.set_column(df.columns.get_loc('Sensor FileName')+1, df.columns.get_loc('Sensor FileName')+1, 50, cell_format)
                ws.set_column(df.columns.get_loc('SPL LineName')+1, df.columns.get_loc('SPL LineName')+1, 20, cell_format)
                ws.set_column(df.columns.get_loc('SPL LineName')+2, df.shape[1], 150, cell_format) # SPL Name
            elif namews in list2:                
                ws.autofilter(0, 0, df.shape[0], df.shape[1])
                ws.set_column(0, 0, 11, cell_format) # ID
                ws.set_column(1, df.shape[1], 50, cell_format) 
                #ws.set_column(3, df.shape[1], 150, cell_format) # Path
            elif namews in list1:
                ws.autofilter(0, 0, df.shape[0], df.shape[1])
                ws.set_column(0, 0, 11, cell_format) # ID
                ws.set_column(df.columns.get_loc('Session Start')+1, df.columns.get_loc('Session End')+1, 22, cell_format) # DateTime
                ws.set_column(df.columns.get_loc('Session Name')+1, df.columns.get_loc('Session Name')+1, 20, session_format) # Session Name
                ws.set_column(df.columns.get_loc('Session MaxGap')+1, df.columns.get_loc('SPL')+1, 20, cell_format) # Session Info
                ws.set_column(df.columns.get_loc('SPL')+2, df.shape[1], 50, cell_format) # Sensors           

        #    for i, width in enumerate(get_col_widths(df)): # Autosize will not work because of the "\n" in the text
        #        ws.set_column(i, i, width, cell_format)

    # Add a format.
    fWRONG = workbook.add_format({'bg_color': '#FFC7CE',
                                'font_color': '#9C0006'})
    fOK = workbook.add_format({'bg_color': '#C6EFCE',
                                'font_color': '#006100'})
    fBLANK = workbook.add_format({'bg_color': '#FFFFFF',
                                'font_color': '#000000'})
    fDUPL = workbook.add_format({'bg_color': '#2385FC',
                                'font_color': '#FFFFFF'})
    fWSPL = workbook.add_format({'bg_color': '#C90119',
                                'font_color': '#FFFFFF'})
    
    # Highlight the values (first is overwrite the others below.....)
    FMaxGap_start = xl_rowcol_to_cell(1, dfALL.columns.get_loc('Session MaxGap')+1, row_abs=True, col_abs=True)
    FMaxGap_end = xl_rowcol_to_cell(dfALL.shape[0]+1, dfALL.columns.get_loc('Session MaxGap')+1, row_abs=True, col_abs=True)
    w['Full_List'].conditional_format('%s:%s' % (FMaxGap_start, FMaxGap_end), {'type':     'cell',
                                                                          'criteria': 'greater than or equal to',
                                                                          'value':    1,
                                                                          'format':   fWRONG})
    
    FDiff_start = xl_rowcol_to_cell(1, dfALL.columns.get_loc('Difference Start [s]')+1, row_abs=True, col_abs=True)
    FDiff_end = xl_rowcol_to_cell(dfALL.shape[0]+1, dfALL.columns.get_loc('Difference Start [s]')+1, row_abs=True, col_abs=True)
    w['Full_List'].conditional_format('%s:%s' % (FDiff_start, FDiff_end), {'type':     'cell',
                                                                          'criteria': 'greater than',
                                                                          'value':    0,
                                                                          'format':   fWRONG})
    
    FFilename_start = xl_rowcol_to_cell(1, dfALL.columns.get_loc('Sensor FileName')+1, row_abs=True, col_abs=True)
    FFilename_end = xl_rowcol_to_cell(dfALL.shape[0]+1, dfALL.columns.get_loc('Sensor FileName')+1, row_abs=True, col_abs=True)
    w['Full_List'].conditional_format('%s:%s' % (FFilename_start, FFilename_end), {'type': 'text',
                                                                                    'criteria': 'containing',
                                                                                    'value':    '[WRONG]',
                                                                                    #'criteria': '=NOT(ISNUMBER(SEARCH($E2,F2)))',
                                                                                    'format': fWRONG})
    w['Full_List'].conditional_format('%s:%s' % (FFilename_start, FFilename_end), {'type': 'text',
                                                                                    'criteria': 'containing',
                                                                                    'value':    '[OK]',
                                                                                    'format': fOK})
    
    
    ListFC = [w['List_Transposed'], w['MBES_NotMatching'], w['SSS_NotMatching'], w['SBP_NotMatching'], 
              w['MAG_NotMatching'], w['SUHRS_NotMatching']]
    
    # Define our range for the color formatting
    MaxGap_start = xl_rowcol_to_cell(1, dfFINAL.columns.get_loc('Session MaxGap')+1, row_abs=True, col_abs=True)
    MaxGap_end = xl_rowcol_to_cell(dfFINAL.shape[0]+1, dfFINAL.columns.get_loc('Session MaxGap')+1, row_abs=True, col_abs=True)
    
    SPL_start = xl_rowcol_to_cell(1, dfFINAL.columns.get_loc('SPL')+1, row_abs=True, col_abs=True)
    SPL_end = xl_rowcol_to_cell(dfFINAL.shape[0]+1, dfFINAL.columns.get_loc('SPL')+1, row_abs=True, col_abs=True)
    
    Sensors_start = xl_rowcol_to_cell(1, dfFINAL.columns.get_loc('SPL')+2, row_abs=True, col_abs=True)
    Sensors_end = xl_rowcol_to_cell(dfFINAL.shape[0]+1, dfFINAL.shape[1], row_abs=True, col_abs=True)
    
    for i in ListFC:
        i.conditional_format('%s:%s' % (MaxGap_start, MaxGap_end), {'type':     'cell',
                                                                    'criteria': 'greater than or equal to',
                                                                    'value':    1,
                                                                    'format':   fWRONG})
        
        i.conditional_format('%s:%s' % (SPL_start, SPL_end), {'type': 'text',
                                                                'criteria': 'containing',
                                                                'value':    'SPLtoSmall',
                                                                'format': fWSPL})
        i.conditional_format('%s:%s' % (SPL_start, SPL_end), {'type': 'text',
                                                                'criteria': 'containing',
                                                                'value':    'NoLineNameFound',
                                                                'format': fWSPL})
        i.conditional_format('%s:%s' % (SPL_start, SPL_end), {'type': 'text',
                                                                'criteria': 'containing',
                                                                'value':    'EmptySPL',
                                                                'format': fWSPL})
        i.conditional_format('%s:%s' % (SPL_start, SPL_end), {'type': 'duplicate',
                                                                'format': fDUPL})
        
        i.conditional_format('%s:%s' % (Sensors_start, Sensors_end), {'type': 'blanks',
                                                                        'format': fBLANK})
        i.conditional_format('%s:%s' % (Sensors_start, Sensors_end), {'type': 'text',
                                                                        'criteria': 'containing',
                                                                        'value':    '[WRONG]',
                                                                        #'criteria': '=NOT(ISNUMBER(SEARCH($E2,F2)))',
                                                                        'format': fWRONG})
        i.conditional_format('%s:%s' % (Sensors_start, Sensors_end), {'type': 'text',
                                                                        'criteria': 'containing',
                                                                        'value':    '[OK]',
                                                                        'format': fOK}) 

    text3 = 'Process Duration: ' + str((datetime.datetime.now() - now))
    w['Summary_Process_Log'].write(3, 0, text3, bold)
        
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
    if os.path.getsize(SPLFilePath) == 0:
        er = SPLFileName
        lnValue = "EmptySPL"
        SessionStart = np.datetime64('NaT')
        SessionEnd = np.datetime64('NaT')
        maxGap = np.nan
        SessionName = "EmptySPL"
        return SessionStart, SessionEnd, lnValue, er, maxGap, SessionName
    
    dfS = pd.read_csv(SPLFilePath, header=None, skipinitialspace=True, names=["Time", "LineName"], parse_dates=["Time"])
    LineName = dfS.iloc[0][-1]
    SessionStart = dfS.iloc[0][0]
    SessionEnd = dfS.iloc[-1][0]
    dfS['timedelta'] = (dfS.Time-dfS.Time.shift())
    dfS.timedelta = dfS.timedelta / np.timedelta64(1, 's')
    maxGap = dfS.timedelta.max()
    SessionName = SessionStart.strftime('%Y%m%d%H%M')
    #DeviceDuration = (pd.to_datetime(DeviceTime[1])-pd.to_datetime(DeviceTime[0]))
    
    # from https://stackoverflow.com/questions/3346430/what-is-the-most-efficient-way-to-get-first-and-last-line-of-a-text-file/3346788  
    # Fast to read first and last row but now use anymore. Leave for future use
    # else:
    #     with open(SPLFilePath, 'rb') as fh:
    #         # 2020/10/17 03:15:24.420, "M1308"
    #         first = next(fh).decode()
    #         try:
    #             fh.seek(-200, os.SEEK_END) # 2 times size of the line
    #             last = fh.readlines()[-1].decode()
    #         except: # file less than 3 lines not working
    #             er = SPLFileName
    #             lnValue = "SPLtoSmall"
    #             SessionStart = first.split(',')[0] 
    #             SessionEnd = np.datetime64('NaT')
    #             return SessionStart, SessionEnd, lnValue, er
    #LineName = first.split(',')[1].replace(' ', '', 1).replace('"', '').replace('\n', '').replace('\r', '')
    #SessionStart = first.split(',')[0]    
    #SessionEnd = last.split(',')[0]    

    #cleaning  
    os.remove(SPLFilePath)
    
    #checking if linename is empty as is use in all other process
    if len(dfS) < 5:
        er = SPLFileName
        lnValue = "SPLtoSmall"
        return SessionStart, SessionEnd, lnValue, er, maxGap, SessionName
    if not LineName: 
        er = SPLFileName
        lnValue = "NoLineNameFound"
        return SessionStart, SessionEnd, lnValue, er, maxGap, SessionName      
    else:
        er = ""
        lnValue = LineName
        return SessionStart, SessionEnd, lnValue, er, maxGap, SessionName

# SPL convertion
def splfc(splList, SPLFormat, dfSPL, dfer, dfSummary, outputFolder, cmd):
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
        SessionStart, SessionEnd, LineName, er, maxGap, SessionName = SPL2CSV(n, outputFolder)        
        dfSPL = dfSPL.append(pd.Series([SessionStart, SessionEnd, LineName, maxGap, SessionName], 
                                index=dfSPL.columns ), ignore_index=True)        
        if er: dfer = dfer.append(pd.Series([er], index=dfer.columns ), ignore_index=True)        
        progressBar(cmd, pbar, index, splList)                   

    # Format datetime
    dfSPL['Session Start'] = pd.to_datetime(dfSPL['Session Start'], format='%Y/%m/%d %H:%M:%S.%f') # format='%d/%m/%Y %H:%M:%S.%f' format='%Y/%m/%d %H:%M:%S.%f' 
    dfSPL['Session End'] = pd.to_datetime(dfSPL['Session End'], format='%Y/%m/%d %H:%M:%S.%f')
    
    dfCount = dfSPL[dfSPL.duplicated(subset='Session Start', keep=False)]

    ## OLD TOOL
    #dfSPL['Session Start'] = pd.to_datetime(dfSPL['Session Start'], format='%d/%m/%Y %H:%M:%S.%f') # format='%d/%m/%Y %H:%M:%S.%f' format='%Y/%m/%d %H:%M:%S.%f' 
    #dfSPL['Session End'] = pd.to_datetime(dfSPL['Session End'], format='%d/%m/%Y %H:%M:%S.%f')
    
    # Add the Sumary Info in a df
    dfSummary = dfSummary.append(pd.Series([SPLFormat, len(splList), int(len(dfCount.index)/2), 'nan', 'nan', datetime.datetime.now() - nowSPL], 
                index=dfSummary.columns ), ignore_index=True) 

    pbar.close() if cmd else print("Subprocess Duration: ", (datetime.datetime.now() - nowSPL)) # cmd vs GUI
    
    return dfSPL, dfer, dfSummary
  
# Sensors convertion and logs creation
def sensorsfc(firstrun, lsFile, ssFormat, ext, cmd, buffer, outputFolder, dfSPL, dfSummary, dfFINAL, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy, vessel):
    """
    Main function to create and rename the sensors files.    
    """
    print('')
    print('##################################################')
    print(f'READING *{ext} ({ssFormat}) FILES')
    print('##################################################')
    
    # Define Dataframe
    # Need to be declare fully in case of manipulated df (DO NOT CHANGE)
    col = ["Session Start", "Session End", "Session Name", "Session MaxGap", "Vessel Name", "Sensor Start",
           "FilePath", "Sensor FileName", "SPL LineName", "Sensor New LineName"]
    dfSensors = pd.DataFrame(columns = col)
    dftmp = pd.DataFrame(columns = col)

    nowSensor = datetime.datetime.now()  # record time of the subprocess
    nowMain = datetime.datetime.now()  # record time of the main process       
    pbar = tqdm(total=len(lsFile)) if cmd else print(f"Note: Output show file counting every {math.ceil(len(lsFile)/10)}") # cmd vs GUI
          
    # Reading the sensors files 
    if firstrun == 'Folder':    
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
                
            # MAG *.csv
            if ssFormat == 'MAG':
                r = pd.read_csv(f, usecols=[0,1,2], nrows=1, parse_dates=[['Date', 'Time']])
                #print(r)
                fStart = r['Date_Time'].iloc[0]
            
            progressBar(cmd, pbar, index, lsFile)
            # Add the Sensor Info in a df
            dfSensors = dfSensors.append(pd.Series(["","", "", "", "", fStart, f, fName, "", ""], index=dfSensors.columns), 
                                         ignore_index=True) 
        
    if firstrun == 'File':
        for index, row in lsFile.iterrows():
            f = row['FilePath']
            fName = os.path.splitext(os.path.basename(f))[0]
            fStart = row['Sensor Start']             
            progressBar(cmd, pbar, index, lsFile)                   
            # Add the Sensor Info in a df
            dfSensors = dfSensors.append(pd.Series(["","", "", "", "", fStart, f, fName, "", ""], index=dfSensors.columns), 
                                         ignore_index=True)        
                          
    pbar.close() if cmd else print("Subprocess Duration: ", (datetime.datetime.now() - nowSensor)) # cmd vs GUI

    # Format datetime
    dfSensors['Sensor Start'] = pd.to_datetime(dfSensors['Sensor Start'])  # format='%d/%m/%Y %H:%M:%S.%f' format='%Y/%m/%d %H:%M:%S.%f'
     
    print('')
    print(f'Listing the *{ext} ({ssFormat}) files.\nPlease wait......')
    nowListing = datetime.datetime.now()  # record time of the subprocess
    
    # Logs and renaming the Sensors files
    # TODO change the order to handle buffer. First occurence
    rcount = 0    
    for index, row in dfSPL.iterrows():                  
        splStart = row['Session Start']
        splEnd = row['Session End']
        splName = row['SPL LineName']
        SessionGap = row['Session MaxGap']
        SessionName = row['Session Name']
        dffilter = dfSensors[dfSensors['Sensor Start'].between(splStart-datetime.timedelta(seconds=int(buffer)), splEnd)]   
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
            dftmp = dftmp.append(pd.Series([splStart, splEnd, SessionName, SessionGap, vessel, SensorStart, SensorFile, SNameCond, splName, SensorNewName], 
                                index=dftmp.columns), ignore_index=True)
 
    print("Subprocess Duration: ", (datetime.datetime.now() - nowListing)) # cmd vs GUI
    
           
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
            #dfsgy.to_csv(outputFolder + "\\" + vessel + "_" + ssFormat + "_WrongStartTime_log.csv", index=True,
            #            columns=["Sensor Start", "Vessel Name", "FilePath", "Sensor FileName"])  
       
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
    #print(f'dfSensors:\n{dfSensors}')   
    #print(f'dftmp:\n{dftmp}') 
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
    dfSummary = dfSummary.append(pd.Series([ssFormat, len(lsFile), int(len(dfCountDupl.index)/2), len(dfsgy) if ssFormat == 'SBP' else 'nan', 'nan', datetime.datetime.now() - nowMain], 
                index=dfSummary.columns), ignore_index=True) 
    
    return dfFINAL, dfSummary, dfMissingSPL, dfDuplSensor, dfSkip, dfsgy

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
    dfSummary.at[ssFormat, 'Processing Time'] = dfSummary.loc[ssFormat].at['Processing Time'] + (datetime.datetime.now() - nowmove)
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

# https://towardsdatascience.com/reordering-pandas-dataframe-columns-thumbs-down-on-standard-solutions-1ff0bc2941d5
def movecol(df, cols_to_move=[], ref_col='', place='After'):
    
    cols = df.columns.tolist()
    if place == 'After':
        seg1 = cols[:list(cols).index(ref_col) + 1]
        seg2 = cols_to_move
    if place == 'Before':
        seg1 = cols[:list(cols).index(ref_col)]
        seg2 = cols_to_move + [ref_col]
    
    seg1 = [i for i in seg1 if i not in seg2]
    seg3 = [i for i in cols if i not in seg1 + seg2]
    
    return(df[seg1 + seg2 + seg3])



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