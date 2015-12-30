Sub Extract_Test_Results()
'
'
    Dim rowNum As Integer, colNum As Integer, currCell As Range
    
    'Initialize pointer to cell C4 in the results table
    rowNum = 4
    colNum = 3
    Set currCell = Workbooks("Series One Test Results v3.xlsm").Sheets(1).Cells(rowNum, colNum)
    
    'Clear previous results
    Range(currCell, currCell.Offset(100, 13)).ClearContents
    
    'Look for keyword "output" from the bottom to enter the OUTPUT section
    Cells.Find(What:="output", After:=Range("A1"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:= _
        False, SearchFormat:=False).Select
    
    '===============================================================
    'PART 1: FIRST GATHER RESULTS FOR EAB, LPZ AND CONTROL ROOM DOSE
    '===============================================================
    
    'Look for keyword "720" to find output results at Time = 720 hour
    Cells.Find(What:="720", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Select
    'Keep looking for "720" until it is a time value in the OUTPUT section
    Do While ActiveCell.Offset(0, -3).Value <> "Time"
        Cells.FindNext(After:=ActiveCell).Select
    Loop
    
    'Once the first instance of "720" is found, check what value it is and store data accordingly
    If ActiveCell.Offset(-2, -3).Value = "Exclusion" Then
        currCell = ActiveCell.Offset(2, 0).Value
        currCell.Offset(0, 1) = ActiveCell.Offset(2, 1).Value
        currCell.Offset(0, 2) = ActiveCell.Offset(2, 2).Value
    ElseIf ActiveCell.Offset(-2, -3).Value = "Low" Then
        currCell.Offset(0, 3) = ActiveCell.Offset(2, 0).Value
        currCell.Offset(0, 4) = ActiveCell.Offset(2, 1).Value
        currCell.Offset(0, 5) = ActiveCell.Offset(2, 2).Value
    ElseIf ActiveCell.Offset(-2, -3).Value = "Control" Then
        currCell.Offset(0, 6) = ActiveCell.Offset(2, 0).Value
        currCell.Offset(0, 7) = ActiveCell.Offset(2, 1).Value
        currCell.Offset(0, 8) = ActiveCell.Offset(2, 2).Value
    End If
        
    'Keep looking for keyword "720" to find other output results at Time = 720 hour
    'If the first instance of "720" is not for "Control" Room dose
    If ActiveCell.Offset(-2, -3).Value <> "Control" Then
        Cells.FindNext(After:=ActiveCell).Select
        'Keep looking for "720" until all EAB, LPZ and Control Room dose are extracted
        Do While ActiveCell.Offset(0, -3).Value = "Time"
            If ActiveCell.Offset(-2, -3).Value = "Exclusion" Then
                currCell = ActiveCell.Offset(2, 0).Value
                currCell.Offset(0, 1) = ActiveCell.Offset(2, 1).Value
                currCell.Offset(0, 2) = ActiveCell.Offset(2, 2).Value
            ElseIf ActiveCell.Offset(-2, -3).Value = "Low" Then
                currCell.Offset(0, 3) = ActiveCell.Offset(2, 0).Value
                currCell.Offset(0, 4) = ActiveCell.Offset(2, 1).Value
                currCell.Offset(0, 5) = ActiveCell.Offset(2, 2).Value
                'Exit Do loop if there isn't a Control Room dose result to come
                If ActiveCell.Offset(4, -3).Value <> "Control" Then
                    Exit Do
                End If
            ElseIf ActiveCell.Offset(-2, -3).Value = "Control" Then
                currCell.Offset(0, 6) = ActiveCell.Offset(2, 0).Value
                currCell.Offset(0, 7) = ActiveCell.Offset(2, 1).Value
                currCell.Offset(0, 8) = ActiveCell.Offset(2, 2).Value
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
            Cells.Find(What:="Xe-135", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=False).Select
            currCell.Offset(0, 10) = ActiveCell.Offset(0, 1).Value
        ElseIf ActiveCell.Offset(0, -3).Value = "Containment" Then
            'Xe-131m in Containment Activity.  If not available, use Cs-137 result.
            Set rngFound2 = Cells.Find(What:="Xe-131m", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=False)
            If Not rngFound2 Is Nothing Then
                rngFound2.Select
                currCell.Offset(0, 13) = ActiveCell.Offset(0, 1).Value
            Else
                'Cs-137 in Containment Activity
                Cells.Find(What:="Cs-137", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                    :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                    False, SearchFormat:=False).Select
                    currCell.Offset(0, 13) = ActiveCell.Offset(0, 1).Value
            End If
            'May not be available: Xe-135 in Containment Activity
            Set rngFound3 = Cells.Find(What:="Xe-135", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=False)
            If Not rngFound3 Is Nothing Then
                rngFound3.Select
                currCell.Offset(0, 12) = ActiveCell.Offset(0, 1).Value
            End If
        End If
    End If
    
    '================================================================
    'PART 3: GATHER ACTIVITY RESULTS FOR IODINE
    '================================================================
    
    'May not exist: Iodine data are available in the I-131 Summary
    Set rngFound4 = Cells.Find(What:="Summary", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
    If Not rngFound4 Is Nothing And rngFound4.Offset(0, -1) = "I-131" Then
        'Initialize rngFound4 to location of title header for the I-131 Summary table
        Set rngFound4 = rngFound4.Offset(3, -1)
        rngFound4.Select
        'Look for keyword "720" to find output results at Time = 720 hour
        Cells.Find(What:="720", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
            :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False, SearchFormat:=False).Select
        'Loop through each column in I-131 Summary table to find Iodine results
        i = 0
        Do While rngFound4.Offset(0, i) <> ""
            If rngFound4.Offset(0, i).Value = "Containment" Then
                currCell.Offset(0, 11) = ActiveCell.Offset(0, 1 + i).Value
            ElseIf rngFound4.Offset(0, i).Value = "Control" Then
                currCell.Offset(0, 9) = ActiveCell.Offset(0, 1 + i).Value
            End If
            'Increment rngFound4 pointer to the next column in the title header
            i = i + 1
        Loop
    End If
    
End Sub
