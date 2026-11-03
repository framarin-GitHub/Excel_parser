Option Explicit

' Port of excel_parser.ts (Office Script).
' Reads the first table on the active sheet and draws one 4x2 block per plot on a new sheet.
' Block position: row = (maxY - Y) * 4, col = (X - minX) * 2.
Public Sub BuildSchema()
    If ActiveSheet.ListObjects.Count = 0 Then
        MsgBox "Il foglio attivo non ha una tabella.", vbExclamation
        Exit Sub
    End If

    Dim texts As Variant
    texts = ActiveSheet.ListObjects(1).Range.Value

    Dim headers() As String, data() As String
    Dim rowCount As Long
    rowCount = ReadRecords(texts, headers, data)
    If rowCount = 0 Then Exit Sub

    Dim cX As Long, cY As Long, cProg As Long, cRep As Long, cId As Long, cVar As Long
    cX = ColIdx(headers, "X-base")
    cY = ColIdx(headers, "Y-altezza")
    cProg = ColIdx(headers, "bl_N_prog")
    cRep = ColIdx(headers, "REP")
    cId = ColIdx(headers, "ID_plot")
    cVar = ColIdx(headers, "VARIETY_NA")

    Dim i As Long, v As Long
    Dim maxRows As Long, minCols As Long
    minCols = 1000000
    For i = 1 To rowCount
        v = CLng(data(i, cX))
        If v < minCols Then minCols = v
        v = CLng(data(i, cY))
        If v > maxRows Then maxRows = v
    Next i

    Dim anchor As Range
    ActiveWorkbook.Worksheets.Add.Activate
    Set anchor = ActiveCell

    Dim target As Range
    For i = 1 To rowCount
        Set target = anchor.Offset((maxRows - CLng(data(i, cY))) * 4, (CLng(data(i, cX)) - minCols) * 2).Resize(4, 2)
        target.Value = BuildCellBlock(data(i, cProg), data(i, cRep), data(i, cId), data(i, cVar), data(i, cY), data(i, cX))
        ApplyBorders target
    Next i
End Sub

' Splits table values into headers and string records. Stops at the first blank first column, like the TS version.
Private Function ReadRecords(texts As Variant, headers() As String, data() As String) As Long
    Dim nRows As Long, nCols As Long, r As Long, c As Long, n As Long
    nRows = UBound(texts, 1)
    nCols = UBound(texts, 2)

    ReDim headers(1 To nCols)
    For c = 1 To nCols
        headers(c) = CStr(texts(1, c))
    Next c

    ReDim data(1 To nRows, 1 To nCols)
    For r = 2 To nRows
        If CStr(texts(r, 1)) = "" Then Exit For
        n = n + 1
        For c = 1 To nCols
            data(n, c) = CStr(texts(r, c))
        Next c
    Next r
    ReadRecords = n
End Function

' 1-based index of a header name; raises if missing.
Private Function ColIdx(headers() As String, colName As String) As Long
    Dim c As Long
    For c = LBound(headers) To UBound(headers)
        If headers(c) = colName Then
            ColIdx = c
            Exit Function
        End If
    Next c
    Err.Raise vbObjectError + 513, "ColIdx", "Column not found: " & colName
End Function

Private Function BuildCellBlock(prog As String, rep As String, id As String, variety As String, y As String, x As String) As Variant
    Dim block(1 To 4, 1 To 2) As Variant
    block(1, 1) = prog:    block(1, 2) = rep
    block(2, 1) = id:      block(2, 2) = ""
    block(3, 1) = variety: block(3, 2) = ""
    block(4, 1) = "Y: " & y: block(4, 2) = "X: " & x
    BuildCellBlock = block
End Function

' Thin outer edges, no inner lines, no diagonal.
Private Sub ApplyBorders(rng As Range)
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlInsideVertical).LineStyle = xlNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlNone
    Dim edge As Variant
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With rng.Borders(edge)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next edge
End Sub
