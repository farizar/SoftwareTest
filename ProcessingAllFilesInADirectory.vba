'
    Dim rowNum, colNum As Integer
    Dim currCell As Range
    Dim rowNum2, colNum2 As Integer
    Dim currCell2 As Range
    Dim lastLoadRow As Double
    
    'Turn off automatic calculation and display alerts which speeds up macro
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
        
    'Clear previous results
    Sheets(1).Range("C4:P139").ClearContents
            
    'Find name of this workbook
    MyWorkbook = ThisWorkbook.Name
    
    'Set the path manually
    MyPath = ThisWorkbook.Path & "\"
    'Retrieve the first file in the directory
    MyName = Dir(MyPath, vbDirectory)
    Do While MyName <> ""
        'Ignore the current directory and the encompassing directory.
        If MyName <> "." And MyName <> ".." And MyName Like "*.o*" Then
            'Use bitwise comparison to make sure MyName is not a directory.
            If (GetAttr(MyPath & MyName) And vbDirectory) <> vbDirectory Then
                
                'Open File
                Workbooks.OpenText Filename:=MyPath & MyName, _
                    Origin:=xlMSDOS, StartRow:=1, DataType:=xlDelimited, _
                    TextQualifier:=xlDoubleQuote, _
                    ConsecutiveDelimiter:=True, Tab:=False, _
                    Semicolon:=False, Comma:=False, Space:=True, _
                    Other:=False, TrailingMinusNumbers:=True, _
                    Local:=True
        
                'Process the file
                Application.Run "Extract_Test_Results"
                
                'Close the File
                Windows(MyName).Close False
            
            End If
        End If
        'Get next entry
        MyName = Dir
    Loop
    
    'Turn on automatic calculation which speeds up macro
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True

End Sub
