' this precedure must be revise by refering date/time+act.work instead only act.work for to check deleted time confirm
Private Sub ConvertRowtoSheet()
' This procedure is for edit/revise before release to main procedure
' Revise on Feb 12, 2024
' ////////////___________READ_ME___________////////////////
' Every activity must undergo time confirmation before its deletion,
' ensuring that the deletion time counter follows confirmation.
' Therefore,the negative value boolean for actual work time
' or execution code always functions correctly.
' ///////////////////////////////////////////////////////////////
On Error GoTo ErrorHandler
    Dim ws As Worksheet, pq As Worksheet
    Dim firstOccurrence As Variant, lastOccurrence As Variant
    Dim queryRange As Range, queryValue As Range
    Dim conbinationText As String
    Dim personelID As String
    Dim next_Array As Long
    Dim TableRowCount As Long
    Dim lookupValue As Variant
    Dim timeAct As Object
    
    Set ws = ThisWorkbook.Worksheets("MainPage")
    Set pq = ThisWorkbook.Worksheets("PQ")
    Set timeAct = ws.ListObjects("timeAct")
    
    lookupValue = ws.Range("WO_select").value
    
    ' Set lookupvalue for each order
    If Not IsNumeric(lookupValue) Then
        Exit Sub
    End If

    ' Clear Table Data Contents and re-size table to default
    timeAct.DataBodyRange.ClearContents
    TableRowCount = timeAct.ListRows.Count
    If TableRowCount > 10 Then
        ws.Rows("15:" & TableRowCount + 4).Delete Shift:=xlUp
    End If

    ' Find the first, last occurrence and set to new range
    firstOccurrence = Application.Match(lookupValue, pq.Columns(1), 0)
    lastOccurrence = Application.Match(lookupValue, pq.Columns(1), 1)
    Set queryRange = pq.Range("A" & firstOccurrence & ":A" & lastOccurrence)
    
    ' Initializing variable for combineText_array array
    Dim cnfText As String
    Dim CNFcounter As String
    Dim saveName As String
    Dim timeDeleteOccured As Boolean, foundTimeDeletion As Boolean
    Dim foundIndex As Long
    Dim fCount As Long
    Dim combineText_array() As String, ID_array() As String, Counter_array() As String
    ReDim combineText_array(0 To 0)
    ReDim ID_array(0 To 0)
    ReDim Counter_array(0 To 0)
    
    ' #Part1# createing array from each activity row in selected range
    For Each queryValue In queryRange
        ' Set negative actual work time to positive
        ' To use in remove duplicate ID staff in next code block
        personelID = queryValue.Offset(0, 6).value
        cnfText = queryValue.Offset(0, 8).value
        CNFcounter = queryValue.Offset(0, 7).value
        timeDeleteOccured = queryValue.Offset(0, 4).value < 0
        
        ' Re-assign value when personelID is empty
        If personelID = "" Then
        
            personelID = Split(cnfText, ",")(0)
            If InStr(1, CStr(personelID), "Reason", vbTextCompare) > 0 Or _
            InStr(1, CStr(personelID), "Cancel", vbTextCompare) > 0 Then
                personelID = saveName
            Else
                saveName = personelID
            End If
                
        End If
        
        ' Re-assign personel id to negative when found negative working time
        If timeDeleteOccured Then
            personelID = "-" & personelID
            foundTimeDeletion = timeDeleteOccured + False
        End If
        
        ' Combine text in each column in the same row.
        ' Offset=1:Act,2:staDate,3:finDate, 4:actWork, 5:PM_actType, 6:ID, 8:cnfText
        conbinationText = queryValue.Offset(0, 1).value & "_" & queryValue.Offset(0, 2).value & "_" & _
                        queryValue.Offset(0, 3).value & "_" & Abs(queryValue.Offset(0, 4).value) & "_" & _
                        queryValue.Offset(0, 5).value
                                
        ' Combine the text and set the initiate array value;
        If combineText_array(0) = "" Then
            combineText_array(0) = conbinationText
            ID_array(0) = personelID
            Counter_array(0) = CNFcounter
            
        Else

            foundIndex = -1
            For fCount = 0 To next_Array
                If combineText_array(fCount) = conbinationText Then
                    foundIndex = fCount
                    Exit For
                End If
            Next fCount
    
            ' When matching a combined text, collect the ID; If not found, get a new array index and value.
            If foundIndex >= 0 Then
                ID_array(foundIndex) = ID_array(foundIndex) & "," & personelID
                Counter_array(foundIndex) = Counter_array(foundIndex) & "," & CNFcounter
            Else
                next_Array = next_Array + 1
                ReDim Preserve combineText_array(0 To next_Array)
                ReDim Preserve ID_array(0 To next_Array)
                ReDim Preserve Counter_array(0 To next_Array)
                combineText_array(next_Array) = conbinationText
                ID_array(next_Array) = personelID
                Counter_array(next_Array) = CNFcounter
            End If
        End If
    Next queryValue

    ' Set initiate row value to use in row index
    Dim splitedCombine() As String
    Dim splitedID() As String
    Dim splitedCounter() As String
    Dim newID_array() As String
    Dim newCounter_array() As String
    Dim priID As Long
    Dim secID As Long
    Dim idIA As Long
    Dim actType As Long

    ' #Part2# filtering and deleting array when found confirmation deleted
    ' when time deletion is detected, remove incorrect confirmation with deletion itself
    If foundTimeDeletion Then
    
        ReDim newID_array(LBound(ID_array) To UBound(ID_array))
        ReDim newCounter_array(LBound(ID_array) To UBound(ID_array))
        ' Remove both opposite value on Personel ID array
        For idIA = LBound(ID_array) To UBound(ID_array)
            '  Split text to array for further execution
            splitedID = Split(ID_array(idIA), ",")
            splitedCounter = Split(Counter_array(idIA), ",")
                ' Loop through each element in personalIDarray
                For priID = LBound(splitedID) To UBound(splitedID)
                ' Check if the current value has an opposite value in the array
                    For secID = LBound(splitedID) To UBound(splitedID)
                        If priID <> secID And "-" & splitedID(priID) = splitedID(secID) And splitedID(priID) <> "" And splitedID(secID) <> "" Then
                            splitedID(priID) = ""
                            splitedID(secID) = ""
                            splitedCounter(priID) = ""
                            splitedCounter(secID) = ""
                            Exit For
                        End If
                    Next secID
                Next priID
            
            ' Create new array from splitedID without empty value
            For priID = LBound(splitedID) To UBound(splitedID)
                If splitedID(priID) <> "" Then
                    If newID_array(idIA) = "" Then
                        newID_array(idIA) = splitedID(priID)
                        newCounter_array(idIA) = splitedCounter(priID)
                    Else
                        newID_array(idIA) = newID_array(idIA) & "," & splitedID(priID)
                        newCounter_array(idIA) = newCounter_array(idIA) & "," & splitedCounter(priID)
                    End If
                End If
            Next priID
            ' Debug.Print "idIA:" & idIA
        Next idIA
        
    Else
        
        ' If no Time-Deletion, copy array value to new array name
        newID_array = ID_array
        newCounter_array = Counter_array
    End If

    Dim textToCell As Long
    Dim replaceCommaID As String
    Dim replaceCommaCounter As String
    Dim cellRowCount As Integer

    ' #Part3# filling combined text of array to cells
     For textToCell = LBound(combineText_array) To UBound(combineText_array)
        
        ' Skip empty array
        If newID_array(textToCell) <> "" Then
        
            If textToCell > 9 Then
                ws.Rows(textToCell + 5 & ":" & textToCell + 5).Insert Shift:=xlDown
                ' timeAct.Resize ws.Range("B5:J" & textToCell + 5)
            End If
    
            ' Create the array values to fill into cells
            splitedCombine = Split(combineText_array(textToCell), "_")
    
            Select Case Mid(splitedCombine(4), 2, 2)
                ' Travelling Time "2TRCT1"
                Case "TR"
                    actType = 6
                ' Waiting Time "2WACT1":
                Case "WA"
                    actType = 7
                ' Working Time "2WOCT1"
                Case Else
                    actType = 5
            End Select
    
            replaceCommaID = Replace(newID_array(textToCell), ",", ", ", 1)
            replaceCommaCounter = Replace(newCounter_array(textToCell), ",", ", ", 1)
            ' Fill array value in to each column in row
            ws.Cells(cellRowCount + 5, 1).value = textToCell + 1
            ws.Cells(cellRowCount + 5, 2).NumberFormat = "@"
            ws.Cells(cellRowCount + 5, 2).value = splitedCombine(0)
            ws.Cells(cellRowCount + 5, 3).value = splitedCombine(1)
            ws.Cells(cellRowCount + 5, 3).NumberFormat = "dd/mm/yyyy"
            ws.Cells(cellRowCount + 5, 4).value = splitedCombine(1)
            ws.Cells(cellRowCount + 5, 4).NumberFormat = "hh:mm"
            ws.Cells(cellRowCount + 5, actType).value = splitedCombine(3)
            ws.Cells(cellRowCount + 5, 8).value = replaceCommaID
            ws.Cells(cellRowCount + 5, 9).value = replaceCommaCounter
            ws.Cells(cellRowCount + 5, 9).NumberFormat = "#"
            ws.Cells(cellRowCount + 5, 8).WrapText = True
            cellRowCount = cellRowCount + 1
        End If
     Next textToCell
     
    ' Set numberformat for actual Start/End of WO
    Range("StartEndCNF").NumberFormat = "dd-mm-yyyy   hh:mm"

ErrorHandler:
    Debug.Print Err.Description
    Exit Sub
End Sub
