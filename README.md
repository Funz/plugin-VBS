[![Build Status](https://travis-ci.org/Funz/plugin-VBS.png)](https://travis-ci.org/Funz/plugin-VBS)

# Funz plugin: VBS

This plugin is dedicated to launch VBS calculations from Funz.
It supports the following syntax and features:

  * Input
    * file type supported: '*.vbs', any other format for resources
    * parameter syntax: 
      * variable syntax: `$(...)`
      * formula syntax: `@{...}`
      * comment char: `'`
    * example input file (associated with dependency file 'sheet.xlsx'):
        ```
        Set xl = CreateObject("Excel.Application")
        Set wb = xl.Workbooks.Open("sheet.xlsx", 0, True) 
        xl.DisplayAlerts = False
        
        '' Can also use name of sheet:
        'wb.Sheets("Feuil1").Range("A1").Value = 123
        wb.Worksheets(1).Range("A1").Value = $(value)
        
        WScript.StdOut.WriteLine("z=" & wb.Worksheets(1).Range("A3").Value)
        
        wb.Close False
        xl.Quit
        Set wb = Nothing
        Set xl = Nothing
        ```
      * will identify input:
        * value, expected to vary inside [0,1] (by default)

  * Output
    * file type supported: 'out.txt' (which is standard output stream)
    * read any named value XXX printed with `WScript.StdOut.WriteLine("XXX=" & ...`
    * example output file:
        ```
        Microsoft (R) Windows Script Host Version 5.8
        Copyright (C) Microsoft Corporation. All rights reserved.
        
        z=125
        ```
        * will return output:
          * z=125


![Analytics](https://ga-beacon.appspot.com/UA-109580-20/plugin-VBS)
