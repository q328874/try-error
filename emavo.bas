Option Explicit
Sub EMAVO()
'Evaluation der neuen EMAVO ggü. der vorherigen Pauschale
'Betrachtungszeitraum: monatlich 2020
'Bezugsgruppe: alle Datensätze zum Stichtag 10.08.2020
Dim lz As Integer, zSpalte As Integer, i As Integer, Spalte As Integer, Zeile As Integer
Dim Zelle As Range, Bereich As Range
Dim myWS As Worksheet

If WorksheetExists("2019") Then
    MsgBox "Tabelle exitiert bereits ... Funktion wird abgebrochen!"
    Exit Sub
End If


Application.ScreenUpdating = False
Worksheets.Select
ActiveWindow.DisplayZeros = False

For Each myWS In ThisWorkbook.Worksheets
    'weiter nur bei Tabellen mit Rohdaten
    If Left(myWS.Name, 3) = "Tab" Then
        'letzte Zeile
        lz = myWS.Cells(Rows.Count, 1).End(xlUp).Row

        'Spalte Betrag in Zahlenformat wandeln
        Set Bereich = myWS.Range("L2:L" & lz)
        For Each Zelle In Bereich
            If Zelle.Value = "" Then Zelle.Offset(0, -1).Value = Zelle.Offset(0, -10).Value
            Zelle.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
            Zelle.Value = Zelle.Value * 1
        Next Zelle

        'Monatsnamen auslesen und Tabellenblatt danach benennen
        myWS.Range("C2").Value = CDate(myWS.Range("C2").Value)
        myWS.Range("C2").NumberFormat = "dd/mm/yy"
        myWS.Name = Format(myWS.Range("C2").Value, "mmmm")

        'Spaltenköpfe Bemerkungen und Betrag um Monatsnamen ergänzen
        myWS.Range("K1").Value = "Bemerkungen " & myWS.Name
        myWS.Columns("K").HorizontalAlignment = xlLeft
        myWS.Range("L1").Value = "Betrag " & myWS.Name
        myWS.Range("L1").HorizontalAlignment = xlLeft

        'Text in Spalte Bemerkung ersetzen
        myWS.Range("K2:K" & lz).Replace "keine 4 Dienstpaare", "k4Dp", xlPart
        myWS.Range("K2:K" & lz).Replace "Tatbestandsmerkmal ", "", xlPart
        
        'überflüssige Spalten löschen
        myWS.Columns("B:J").Delete
    End If
Next myWS

Application.ScreenUpdating = True
End Sub

Public Function WorksheetExists(ByVal WorksheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = (Sheets(WorksheetName).Name <> "")
    On Error GoTo 0
End Function
