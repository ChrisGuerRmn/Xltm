Attribute VB_Name = "Module1"
Sub DetectHygrometryPeriodsWithDuration()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Modifie si ta feuille a un autre nom

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    Dim startRow As Long
    Dim countAbove As Long
    Dim outputRow As Long
    outputRow = 2 ' Ligne de d�part pour les r�sultats

    ' Cr�e une nouvelle feuille pour les r�sultats
    Dim resultWs As Worksheet
    Set resultWs = ThisWorkbook.Sheets.Add
    resultWs.Name = "P�riodes D�tection"
    resultWs.Range("A1").Value = "D�but"
    resultWs.Range("B1").Value = "Fin"
    resultWs.Range("C1").Value = "Dur�e (hh:mm:ss)"

    i = 2 ' Supposons que les donn�es commencent � la ligne 2
    Do While i <= lastRow
        If ws.Cells(i, 2).Value > 50 Then
            startRow = i
            countAbove = 1

            Do While ws.Cells(i + countAbove, 2).Value > 50 And (i + countAbove) <= lastRow
                countAbove = countAbove + 1
            Loop

            If countAbove >= 2 Then
                Dim startTime As Variant
                Dim endTime As Variant
                startTime = ws.Cells(startRow, 1).Value
'                endTime = ws.Cells(startRow + countAbove - 1, 1).Value
                endTime = ws.Cells(startRow + countAbove, 1).Value

                resultWs.Cells(outputRow, 1).Value = startTime
                resultWs.Cells(outputRow, 2).Value = endTime
                resultWs.Cells(outputRow, 3).Value = endTime - startTime
                resultWs.Cells(outputRow, 3).NumberFormat = "[hh]:mm:ss"

                outputRow = outputRow + 1
            End If

            i = i + countAbove
        Else
            i = i + 1
        End If
    Loop

    MsgBox "Analyse termin�e. P�riodes d�tect�es : " & outputRow - 2
End Sub

