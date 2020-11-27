# Linename comparison and rename tool between SPL and sensors

## Introduction

This python script will compare Linename (creating logs) and remane (if needed) the sensors based on the *-Position.fbf or *-.fbz file.

## Setup

Several modules need to be install before using the script. You will need:

+ `$ pushd somepath\spl-sensors-comp-ren`
+ `$ pip install .`

## Usage

The tool is a GUI based on Gooey but the tool can be use in cmd too. See below:

```
usage: splsensors.py [-h] [-r Recurse into the subfolders?] [-n Rename the files?] [-e List of folder to be exclude]
                     [-i SPL Root Path] [-p SPL Position File Name] [-A ALL Folder Path] [-X XTF Folder Path]
                     [-S SGY SBP Folder Path] [-M CSV MAG Folder Path] [-H SGY SUHRS Folder Path]
                     [-o Output Logs Folder]

Linename comparison and rename tool between SPL and sensors

optional arguments:
  -h, --help            show this help message and exit

SPL Options:
  -i SPL Root Path, --splFolder SPL Root Path
                        This is the path where the *.fbf/*.fbz files to process are. (Root Session Folder)
  -p SPL Position File Name, --splPosition SPL Position File Name
                        SPL position file to be use to rename the sensor without extention.

Sensors Options:
  -A ALL Folder Path, --allFolder ALL Folder Path
                        ALL Root path. This is the path where the *.all files to process are.
  -X XTF Folder Path, --xtfFolder XTF Folder Path
                        XTF Root path. This is the path where the *.xtf files to process are.
  -S SGY SBP Folder Path, --sgySBPFolder SGY SBP Folder Path
                        SGY SBP Root path. This is the path where the *.sgy files to process are.
  -M CSV MAG Folder Path, --csvMAGFolder CSV MAG Folder Path
                        CSV MAG Root path. This is the path where the *.csv files to process are.
  -H SGY SUHRS Folder Path, --sgySUHRSFolder SGY SUHRS Folder Path
                        SGY SUHRS Root path. This is the path where the *.sgy files to process are.

Output Options:
  -o Output Logs Folder, --output Output Logs Folder
                        Output folder to save all the logs files.

Additional Options:
  -r Recurse into the subfolders?, --recursive Recurse into the subfolders?
                        Choise = [yes, no] Default = no
  -n Rename the files?, -rename Rename the files?
                        Choise = [yes, no] Default = no
  -e List of folder to be exclude, --excludeFolder List of folder to be exclude
                        List all folder that need to be excluded from the recurcive search. (eg.: DNP,DoNotProcess)
                        Comma separated and NO WHITESPACE! Note: This just apply to the sensors folders

Example:
 To rename the sensors files use python splsensors.py -r no -n no -e DNP,DoNotProcess -i c:\temp\spl\ -p FugroBrasilis-CRP-Position -A c:\temp\spl\ -X c:\temp\xtf\ -o c:\temp\log\
```

## Export products

+ Logs files with all information needed to QC the data
  + *_Full_Log.csv (Full log per sensors)
  + _*_FINAL_Log.xlsx (Log used to compare the LineName between sensors)
+ Optional: Rename the sensor files

## TO DO

+ TrackPlot.shp
+ Improve Renaming otion
+ Improve Missing SPL Sheet
