# Excel Interop

    Usage: excel [OPTIONS] ExcelFile PasteFile Cell1 Cell2 OutCell1 OutCell2 [Chart1 [Chart2 ..]]
       or: excel =FORMULA

    Opens ExcelFile and places the contents of PasteFile from position given.
    Position is the range between Cell1 and Cell2. Unused cells are cleared.
    Echos all filled rows in the range between OutCell1 and OutCell2.
    Any charts named will be saved to <ExcelFile_ChartN>.png.
    'Sheet.ChartN' can be used if multiple charts has the same name.

        -p PasteSheet   Select the sheet that should be pasted to.
        -s OutSheet     Select the sheet that should be outputted.
        -m Macro        Run macro after paste. If -m is used multiples
                        times, more than 1 macro can be executed.
        -n              Do not save workbook
        -w              display Excel dialogs. Default is to surpress.
        -#              Replace errors with blanks in output
        -b [CHAR]       Set cell-spacing character in output to
                        'space' or 'CHAR' (default is 'tab')
        -t              Insert PasteFile as text instead of numbers

    If '~' is specified as PasteFile no file is loaded and Cell1 and Cell2
    should not be specified.


    Supported Cell formats:
      A2
      A,2
      1,2
      $A$2
      $1,$2
      R1C2

      A2 A26
      A2:A26
      A:A

    Examples:

      Sort numbers in ascending order. Numbers outputted to sorted_ascending.txt
        excel Sort.xlsx numbers.txt A2:A26 C2:C26 > sorted_ascending.txt
        
      Sort numbers in descending order. Numbers outputted to sorted_descending.txt
        excel Sort.xlsx numbers.txt A2:A26 D2:D26 > sorted_descending.txt
        
      Sort numbers and also save the selected figures.
        excel Sort.xlsx numbers.txt A2:A26 C2:C26 "Chart 1" > sorted_ascending.txt
        excel Sort.xlsx numbers.txt A2:A26 C2:C26 "Chart 2" > sorted_descending.txt
        
      Sort numbers and also save the selected figures. Output both results (column C and D).
        excel Sort.xlsx numbers.txt A2:A26 C2:D26 "Chart 1" "Chart 2" > sorted_asc_and_desc.txt

      Run macro to sort numbers.
        excel Sort_Macro.xlsm -m "Macro1" numbers.txt A2:A26 C2:D26 > sorted_asc_and_desc.txt

      Using a excel file with multiple sheets. 
        excel Sort_Sheets.xlsx numbers.txt  -p "Input" A2:A26  -s "Output"  B2:C26 > sorted_asc_and_desc.txt
        
      
      
      If only ExcelFile is given, the content of first spreadsheet is written:
        excel Sort_Sheets.xlsx > all_content.txt
        
      If no cells are given, Pastefile is written from A1, and the content of first spreadsheet is written:
        excel Sort_Sheets.xlsx numbers.txt > all_content.txt

      Sort numbers in descending order. Numbers outputted to sorted_descending.txt.
      Note that since no row numbers are given, paste starts at A1
        excel Sort.xlsx numbers.txt A:A D:D > sorted_descending.txt