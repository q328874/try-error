Option Explicit
Sub EMAVO()
'Evaluation der neuen EMAVO ggü. der vorherigen Pauschale
'Betrachtungszeitraum: Kalenderjahr 2019
'Vergleichswert: Dez. 2018
'Bezugsgruppe: alle DPNA-Datensätze zum Stichtag 05.02.2020
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

For Each myWS In Worksheets
    'letzte Zeile
    lz = myWS.Cells(Rows.Count, 1).End(xlUp).Row

    'Spalte Bemerkungen füllen
    'Spalte Betrag in Zahlenformat wandeln
    Set Bereich = myWS.Range("L2:L" & lz)
    For Each Zelle In Bereich
        If Zelle.Value = "" Then Zelle.Offset(0, -1).Value = Zelle.Offset(0, -10).Value
        Zelle.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        Zelle.Value = Zelle.Value * 1
    Next Zelle

    'Tabellenblatt benennen
    myWS.Range("C2").Value = CDate(myWS.Range("C2").Value)
    myWS.Range("C2").NumberFormat = "dd/mm/yy"
    myWS.Name = Format(myWS.Range("C2").Value, "mmmm")

    'Spaltenköpfe Bemerkungen und Betrag um Monatsnamen ergänzen
    myWS.Range("K1").Value = "Bemerkungen " & myWS.Name
    myWS.Columns("K").HorizontalAlignment = xlLeft
    myWS.Range("L1").Value = "Betrag " & myWS.Name
    myWS.Range("L1").HorizontalAlignment = xlLeft

    'Text ersetzen
    myWS.Range("K2:K" & lz).Replace "keine 4 Dienstpaare", "k4Dp", xlPart
    myWS.Range("K2:K" & lz).Replace "Tatbestandsmerkmal ", "", xlPart
Next myWS

'Tabellenblatt 2019 als letztes Blatt anfügen
With ActiveWorkbook
    .Worksheets.Add after:=Worksheets(Worksheets.Count)
    .ActiveSheet.Name = "2019"
    '.ActiveWindow.DisplayZeros = False
End With

With Worksheets("2019")
    'Spalte A mit Namen nach Blatt 2019 kopieren
    Tabelle1.Range("A1:A" & lz).Copy .Range("A1:A" & lz)

    'Spalten "OE/Funktion" und "Zulage alt" in Tabelle 2019 einfügen
    .Columns("B:C").Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromLeftOrAbove
    .Range("B1:C" & lz).Borders.LineStyle = xlContinuous
    .Range("B1:C1").HorizontalAlignment = xlLeft
    .Range("B1").Value = "OE/Funktion"
    .Range("C1").Value = "Zulage alt"

    'Spalten Bemerkung und Betrag aus jedem Monatsblatt nach 2019 kopieren
    zSpalte = 4
    For i = 1 To 12
        Sheets(i).Range("K1:L" & lz).Copy .Cells(1, zSpalte)
        zSpalte = zSpalte + 2
    Next i
    
    'Summenspalte anfügen
    .Columns("AB").Insert Shift:=xlToRight
    .Columns("AB").HorizontalAlignment = xlCenter
    .Range("AB1").Value = "Summe"
    .Range("AB1").HorizontalAlignment = xlLeft
    .Range("AB1:AB" & lz).Borders.LineStyle = xlContinuous
    For Zeile = 2 To lz
        For Spalte = 5 To 27 Step 2
            .Cells(Zeile, 28).Value = .Cells(Zeile, 28).Value + Cells(Zeile, Spalte).Value
        Next Spalte
    Next Zeile
    
    'Autofilter aktivieren
    .Rows("1:1").AutoFilter

    'Spaltenbreite in Tabelle 2019 automatisch anpassen
    .Columns("A:AA").EntireColumn.AutoFit

    'Fenster fixieren
    .Range("B2").Select
    ActiveWindow.FreezePanes = True

    'Nullwerte ausblenden
    ActiveWindow.DisplayZeros = False
    
    'nach vorn verschieben
    .Move before:=Worksheets(1)
End With

Application.ScreenUpdating = True

End Sub

Public Function WorksheetExists(ByVal WorksheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = (Sheets(WorksheetName).Name <> "")
    On Error GoTo 0
End Function
