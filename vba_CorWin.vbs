Option Explicit

Dim CWX As String
Dim colIndex As String
Dim rowIndex As Integer
Dim inputPPN As String
Dim folderPath As String
Dim fileName As String
Dim mainWorkBook As Workbook
Sub read_WinIBW_Data()
    'original : https://stackoverflow.com/questions/20390397/reading-entire-text-file-using-vba#answer-20390880
    Dim data As String
    
    Open folderPath & "\export_WinIBW.txt" For Input As #1
    rowIndex = 1
    Do Until EOF(1)
        Line Input #1, data
        rowIndex = rowIndex + 1
        If Not data = Empty Then
'Lance le bon script
            Select Case CWX
                Case "[CW1]"
                    ctrlUA103format data
                Case Else
                    MsgBox "La valeur entrée en H2 ne devrait pas être possible"
            End Select
        Else
            rowIndex = rowIndex - 1
        End If
    Loop
    Close #1
    
    'Formattage cellules
    With mainWorkBook.Sheets("Résultats").Range("A2:" & colIndex & rowIndex)
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        .Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
End Sub
Sub formatEnTetes()
    'https://www.automateexcel.com/vba/format-cells/
    'Crée les en-têtes pour la feuille "Résultats"
    
    mainWorkBook.Worksheets("Résultats").Activate
    Select Case CWX
        Case "[CW1]"
            Range("A1").Value = "PPN"
            Range("B1").Value = "UA103$a"
            Range("C1").Value = "UA103$b"
            Range("D1").Value = "Résultats"
            colIndex = "D"
        Case Else
            MsgBox "La valeur entrée en H2 ne devrait pas être possible"
    End Select
    With mainWorkBook.Worksheets("Résultats").Range("A1:" & colIndex & "1")
        .Interior.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(255, 255, 255)
    End With
    
    'Pour éviter que les PPN deviennent des nombres
    Range("A:" & colIndex).NumberFormat = "@"
    
End Sub
Sub ctrlUA103format(data)
    'le get des dollars se fait dans WinIBW. ici je compare juste
    Dim UA103a As String, UA103b As String, PPNval As String, output As String, dataSplit
    
    dataSplit = Split(data, ";_;")
    PPNval = dataSplit(0)
    UA103a = dataSplit(1)
    UA103b = dataSplit(2)
    
    'Vérification du format des données
    If UA103a <> Replace(Replace(UA103a, ".", "X"), "?", "X") Then
        output = appendNote(output, "PB en UA103a")
        mainWorkBook.Sheets("Résultats").Range("B" & rowIndex).Interior.Color = RGB(255, 192, 0)
    End If
    If UA103b <> Replace(Replace(UA103b, ".", "X"), "?", "X") Then
        output = appendNote(output, "PB en UA103b")
        mainWorkBook.Sheets("Résultats").Range("C" & rowIndex).Interior.Color = RGB(255, 192, 0)
    End If

    If output = Empty Then
        If UA103a = "00000000" And UA103b = "00000000" Then
            output = "PAS DE 103 $a NI $b"
            mainWorkBook.Sheets("Résultats").Range("A" & rowIndex & ":D" & rowIndex).Interior.Color = RGB(0, 176, 240)
        Else
            output = "OK"
            mainWorkBook.Sheets("Résultats").Range("A" & rowIndex & ":D" & rowIndex).Interior.Color = RGB(146, 208, 80)
        End If
    Else
        mainWorkBook.Sheets("Résultats").Range("A" & rowIndex & ":D" & rowIndex).Interior.Color = RGB(255, 0, 0)
        If InStr(output, "UA103a") > 0 Then
            mainWorkBook.Sheets("Résultats").Range("B" & rowIndex).Interior.Color = RGB(255, 192, 0)
        End If
        If InStr(output, "UA103b") > 0 Then
            mainWorkBook.Sheets("Résultats").Range("C" & rowIndex).Interior.Color = RGB(255, 192, 0)
        End If
    End If

    mainWorkBook.Sheets("Résultats").Range("A" & rowIndex).Value = PPNval
    mainWorkBook.Sheets("Résultats").Range("B" & rowIndex).Value = UA103a
    mainWorkBook.Sheets("Résultats").Range("C" & rowIndex).Value = UA103b
    mainWorkBook.Sheets("Résultats").Range("D" & rowIndex).Value = output
    
End Sub

Function appendNote(var As String, text As String)
    If var = "" Then
        var = text
    Else
        var = var & Chr(10) & text
    End If
    appendNote = var
End Function
Sub cleanData()
    Worksheets("Résultats").Activate
    Range("A:ZZ").Delete
    Worksheets("Introduction").Activate
    Range("I2:J999999").ClearContents
    Range("I2").Select
End Sub
Sub imp_PPN_Alma()
    
    
Dim nbRow As Integer
Dim exportAlma As Workbook

Set mainWorkBook = ActiveWorkbook
folderPath = Application.ActiveWorkbook.Path
Workbooks.Open fileName:=folderPath & "\export_alma_CorWin.xlsx"
Set exportAlma = Workbooks("export_alma_CorWin.xlsx")

nbRow = Cells(Rows.count, "K").End(xlUp).Row

'Récupère les données
Dim PPN
    
For rowIndex = 2 To nbRow
    PPN = exportAlma.Worksheets("Results").Cells(rowIndex, 10).Value
    PPN = Right(Mid(PPN, InStr(PPN, "(PPN)"), 14), 9)
    mainWorkBook.Worksheets("Introduction").Cells(rowIndex, 9).Value = PPN
Next
Workbooks("export_alma_CorWin.xlsx").Close
mainWorkBook.Worksheets("Introduction").Activate
Range("A2").Select
    
End Sub
Sub copyPPNlist()
    Dim nbPPN As Integer

    folderPath = Application.ActiveWorkbook.Path
    
    ActiveWorkbook.Worksheets("Introduction").Activate
    Range("I:I").Sort key1:=Cells(2, 9), order1:=xlAscending, Header:=xlYes
    Range("I2").Insert xlShiftDown
    Range("I3").Copy
    Range("I2").PasteSpecial Paste:=xlPasteFormats
    Range("I2").Value = folderPath
    nbPPN = Application.WorksheetFunction.CountA(Range("I:I"))
    
    Range("I2:I" & nbPPN).Copy
    
    
End Sub
Sub Main()
'Timer : https://www.thespreadsheetguru.com/the-code-vault/2015/1/28/vba-calculate-macro-run-time

'Timer : début
Dim StartTime As Double
Dim MinutesElapsed As String
StartTime = Timer

Set mainWorkBook = ActiveWorkbook
folderPath = Application.ActiveWorkbook.Path

mainWorkBook.Worksheets("Introduction").Activate
CWX = Right(Range("H2").Value, 5)
Range("I:I").Sort key1:=Cells(2, 9), order1:=xlAscending, Header:=xlYes

formatEnTetes

read_WinIBW_Data

'Lance un script additionnel si nécessaire
Select Case CWX
    Case "[CW1]"
        'fct
End Select

mainWorkBook.Worksheets("Résultats").Activate
Columns("A:" & colIndex).AutoFit
Rows("1:" & rowIndex).AutoFit

'Formattage spéciaux pour un script
Select Case CWX
    Case "[CW1]"
        'format
End Select

Range("A1").Select

mainWorkBook.Worksheets("Introduction").Range("I2") = "Ø"

'Timer suite & fin
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
MsgBox "Exécution terminée en " & MinutesElapsed & "."

End Sub
