Public MyName As String
Public MyWorkbook As String

Sub ProcessMultipleTextFilesInADirectory()
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
Sub Extract_Test_Results()
'
'
    Dim rowNum As Integer, colNum As Integer, currCell As Range
   
    'Initialize pointer to store results to the correct case
    temp = Split(MyName, ".")
    testCase = Mid(temp(0), 5)
    Set currCell = Workbooks(MyWorkbook).Sheets(1).Range("A:A"). _
        Find(What:=testCase, After:=Range("A1"), LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        
    'Determine ending time of simulation based on convection data with a minimum of 2 hours
    Cells.Find(What:="Convection", After:=Range("A1"), LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Select
    endTime = WorksheetFunction.Max(2, ActiveCell.Offset(3, 0).Value)
    
    'Look for keyword "output" from the bottom to enter the OUTPUT section
    Cells.Find(What:="output", After:=Range("A1"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:= _
        False, SearchFormat:=False).Select

    '===============================================================
    'PART 1: FIRST GATHER RESULTS FOR EAB, LPZ AND CONTROL ROOM DOSE
    '===============================================================
    
    'Look for keyword endTime to find output results
    Cells.Find(What:=endTime, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Select
    'Keep looking for endTime until it is a time value in the OUTPUT section
    Do While ActiveCell.Offset(0, -3).Value <> "Time"
        Cells.FindNext(After:=ActiveCell).Select
    Loop
    
    'Once the first instance of endTime is found, check what value it is and store data accordingly
    If ActiveCell.Offset(-2, -3).Value = "Exclusion" Then
        currCell.Offset(0, 2) = ActiveCell.Offset(2, 0).Value
        currCell.Offset(0, 3) = ActiveCell.Offset(2, 1).Value
        currCell.Offset(0, 4) = ActiveCell.Offset(2, 2).Value
    ElseIf ActiveCell.Offset(-2, -3).Value = "Low" Then
        currCell.Offset(0, 5) = ActiveCell.Offset(2, 0).Value
        currCell.Offset(0, 6) = ActiveCell.Offset(2, 1).Value
        currCell.Offset(0, 7) = ActiveCell.Offset(2, 2).Value
    ElseIf ActiveCell.Offset(-2, -3).Value = "Control" Then
        currCell.Offset(0, 8) = ActiveCell.Offset(2, 0).Value
        currCell.Offset(0, 9) = ActiveCell.Offset(2, 1).Value
        currCell.Offset(0, 10) = ActiveCell.Offset(2, 2).Value
    End If
        
    'Keep looking for keyword endTime to find other output results
    'if the first instance of endTime was not for "Control" Room dose and there is LPZ dose available
    If ActiveCell.Offset(-2, -3).Value <> "Control" And ActiveCell.Offset(4, -3).Value = "Low" Then
        Cells.FindNext(After:=ActiveCell).Select
        'Keep looking for endTime until all EAB, LPZ and Control Room dose are extracted
        Do While ActiveCell.Offset(0, -3).Value = "Time"
            If ActiveCell.Offset(-2, -3).Value = "Exclusion" Then
                currCell.Offset(0, 2) = ActiveCell.Offset(2, 0).Value
                currCell.Offset(0, 3) = ActiveCell.Offset(2, 1).Value
                currCell.Offset(0, 4) = ActiveCell.Offset(2, 2).Value
            ElseIf ActiveCell.Offset(-2, -3).Value = "Low" Then
                currCell.Offset(0, 5) = ActiveCell.Offset(2, 0).Value
                currCell.Offset(0, 6) = ActiveCell.Offset(2, 1).Value
                currCell.Offset(0, 7) = ActiveCell.Offset(2, 2).Value
                'Exit Do loop if there isn't a Control Room dose result to come
                If ActiveCell.Offset(4, -3).Value <> "Control" Then
                    Exit Do
                End If
            ElseIf ActiveCell.Offset(-2, -3).Value = "Control" Then
                currCell.Offset(0, 8) = ActiveCell.Offset(2, 0).Value
                currCell.Offset(0, 9) = ActiveCell.Offset(2, 1).Value
                currCell.Offset(0, 10) = ActiveCell.Offset(2, 2).Value
                'Exit Do loop once the final Control Room dose results have been
                'obtained.
                Exit Do
            End If
            Cells.FindNext(After:=ActiveCell).Select
        Loop
    End If
    
    '================================================================
    'PART 2: GATHER ACTIVITY RESULTS FOR XENON
    '================================================================
    
    'May not exist: Xenon data are available in the Nuclide Inventory table
    Set rngFound1 = Cells.Find(What:="Inventory:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
    If Not rngFound1 Is Nothing Then
        rngFound1.Select
        If ActiveCell.Offset(0, -4).Value = "Control" Then
            'Xe-135 in Control Room Activity
            Set rngFound2 = Cells.Find(What:="Xe-135", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=False)
            If Not rngFound2 Is Nothing Then
                rngFound2.Select
                currCell.Offset(0, 12) = ActiveCell.Offset(0, 1).Value
            End If
        ElseIf ActiveCell.Offset(0, -3).Value = "Containment" Then
            'Xe-131m in Containment Activity.  If not available, use Cs-137 result.
            Set rngFound2 = Cells.Find(What:="Xe-131m", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=False)
            If Not rngFound2 Is Nothing Then
                rngFound2.Select
                currCell.Offset(0, 15) = ActiveCell.Offset(0, 1).Value
            Else
                'Cs-137 in Containment Activity
                Set rngFound3 = Cells.Find(What:="Cs-137", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                    :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                    False, SearchFormat:=False)
                If Not rngFound3 Is Nothing Then
                    rngFound3.Select
                    currCell.Offset(0, 15) = ActiveCell.Offset(0, 1).Value
                End If
            End If
            'May not be available: Xe-135 in Containment Activity
            Set rngFound4 = Cells.Find(What:="Xe-135", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=False)
            If Not rngFound4 Is Nothing Then
                rngFound4.Select
                currCell.Offset(0, 14) = ActiveCell.Offset(0, 1).Value
            End If
        End If
    End If
    
    '================================================================
    'PART 3: GATHER ACTIVITY RESULTS FOR IODINE
    '================================================================
    
    'May not exist: Iodine data are available in the I-131 Summary
    Set rngFound5 = Cells.Find(What:="Summary", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
    If Not rngFound5 Is Nothing And rngFound5.Offset(0, -1) = "I-131" Then
        'Initialize rngFound5 to location of title header for the I-131 Summary table
        Set rngFound5 = rngFound5.Offset(3, -1)
        rngFound5.Select
        'Look for keyword endTime to find output results
        Cells.Find(What:=endTime, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
            :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False, SearchFormat:=False).Select
        'Loop through each column in I-131 Summary table to find Iodine results
        i = 0
        Do While rngFound5.Offset(0, i) <> ""
            If rngFound5.Offset(0, i).Value = "Containment" Then
                currCell.Offset(0, 13) = ActiveCell.Offset(0, 1 + i).Value
            ElseIf rngFound5.Offset(0, i).Value = "Control" Then
                currCell.Offset(0, 11) = ActiveCell.Offset(0, 1 + i).Value
            End If
            'Increment rngFound4 pointer to the next column in the title header
            i = i + 1
        Loop
        
        'A possibility exist that the table is separated if it has more than 3 title headers, causing
        'the control column to shift downwards.  Check for this instance two rows below.
        If (ActiveCell.Offset(2, 0) = "Control") Then
            'Look for keyword endTime to find output results
            Cells.Find(What:=endTime, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=False).Select
            currCell.Offset(0, 11) = ActiveCell.Offset(0, 1).Value
        End If
    End If
    
End Sub
