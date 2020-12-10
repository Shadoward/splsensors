# Linename comparison tool between SPL and sensors

## Introduction

This python script will compare Linename (creating logs) the sensors based on the *-Position.fbf or *-.fbz file.

## Setup

Beacuse of a bug in the Starfix converter you need to copy the .\starfixRC\Fugro.DescribedData2Ascii.exe in the following Starfix folder: C:\ProgramData\Fugro\Starfix2018\ or C:\ProgramData\Fugro\Starfix2020\

## Usage

The tool is a GUI based on Gooey.

## Export products

+ Logs files with all information needed to QC the data
  + *_Full_Log.csv (Full log per sensors)
  + _*_FINAL_Log.xlsx (Log used to compare the LineName between sensors)

## TO DO

+ speed up read sgy process
