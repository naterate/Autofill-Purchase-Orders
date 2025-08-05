Sub AutoFillItemDetails()
    Dim dbWs As Worksheet
    Dim poWs As Worksheet
    Dim itemNo As String
    Dim foundCell As Range
    Dim rng As Range
    Dim lastRow As Long

    ' use 1st sheet as PO and 2nd sheet as DB
    Set poWs = ThisWorkbook.Sheets(1)
    Set dbWs = ThisWorkbook.Sheets(2)

    ' assumes headers in row 1
    Set rng = dbWs.Range("A2:A" & dbWs.Cells(Rows.Count, 1).End(xlUp).Row)
    
    ' loop through DB sheet rows starting from row 2
    lastRow = poWs.Cells(Rows.Count, 2).End(xlUp).Row
    
    ' fill Name, Description, Custom, U/M (Cols 3, 4, 5, 6 in PO)
    Dim i As Long
    For i = 23 To lastRow
        itemNo = Trim(poWs.Cells(i, 2).Value)
        If itemNo <> "" Then
            Set foundCell = rng.Find(What:=itemNo, LookIn:=xlValues, LookAt:=xlWhole)
            If Not foundCell Is Nothing Then
                If poWs.Cells(i, 3).Value = "" Then poWs.Cells(i, 3).Value = foundCell.Offset(0, 1).Value ' Name
                If poWs.Cells(i, 4).Value = "" Then poWs.Cells(i, 4).Value = foundCell.Offset(0, 2).Value ' Description
                If poWs.Cells(i, 5).Value = "" Then poWs.Cells(i, 5).Value = foundCell.Offset(0, 3).Value ' Custom
                If poWs.Cells(i, 6).Value = "" Then poWs.Cells(i, 6).Value = foundCell.Offset(0, 4).Value ' U/M
            End If
        End If
    Next i

    MsgBox "Details updated based on Item No.", vbInformation
End Sub
