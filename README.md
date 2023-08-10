#SpreadshetsFileUnion
Is a console application .NET Core that find file csv or xlsx in folders with same header and generate one union file of it.

The library used to generate excel file is Syncfusion.XlsIO.Net.Core

Example :

FOLDER A
    FileCSV1.csv
    FileCSV2.csv

Generate a file named FolderA_merged_datetime.csv with the union of content of FileCSV1 and FileCSV2 with same header not repeted.

FOLDER B
    Filexls1.xlsx
    Filexls2.xlsx

Generate a file named FolderB_merged_datetime.xlsx with the union of content of Filexls1 and Filexls2 with same header not repeted
