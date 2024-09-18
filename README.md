# WDBXlsxConverter

This program allows you to convert the WDB database files from the FF13 game trilogy to excel (xlsx) format. the program should be launched from a command prompt terminal with a few argument switches to perform the conversion function.

**Game Codes:**
<br>``-ff131`` For FF13-1 WDB files
<br>``-ff132`` For FF13-2 and FF13-LR 's WDB files

<br>**Optional switches:**
<br>``-?`` Display the help page. will also display few argument examples too.
<br>``-i`` Converts the WDB file's data into a new xlsx file without the fieldnames (only when gamecode is -ff131)

<br>Commandline usage examples:
<br>``WDBXlsxConverter.exe -? ``
<br>``WDBXlsxConverter.exe -ff131 "auto_clip.wdb" ``
<br>``WDBXlsxConverter.exe -ff131 "auto_clip.wdb" -i ``

## Important notes
- The WDB file has to be specified after the game code argument switch.

- Field names will be present in the xlsx file only for some of 13-1's WDB files. refer to this [page](https://github.com/LR-Research-Team/Datalog/wiki/WDB-Field-Names) for information about the field names.

- When using the `-ff131` game code switch, you can specify an optional `-i` argument switch after the WDB file, to prevent generating field names for the records.

## For developers
- The following package was used for writing to xlsx format:
<br>**ClosedXML** - https://github.com/ClosedXML/ClosedXML

- Refer to this [page](https://github.com/LR-Research-Team/Datalog/wiki/WDB) for information about the file structure of the WDB file.
