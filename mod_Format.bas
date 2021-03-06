Option Explicit
Sub Formatierung()
' das Textformat der Tabelle1 in echte Datum- und Zeitformate wandeln
' nicht benötigte Spalten entfernen

' Vorspiel
Worksheets("Tabelle1").Activate
ActiveWindow.DisplayZeros = False
Application.ScreenUpdating = False
Dim Laufzeit As Double
Laufzeit = Timer
Dim d As Integer, lz As Integer
Dim Zelle As Range, Bereich As Range
Dim Jahr As Date, Neujahr As Date, Karfreitag As Date
Dim Ostersonntag As Date, Ostermontag As Date, Maifeiertag As Date
Dim Himmelfahrt As Date, Pfingstsonntag As Date, Pfingstmontag As Date
Dim TDE As Date, Reformationstag As Date, BuBTag As Date
Dim ErsterWeihnachtstag As Date, ZweiterWeihnachtstag As Date

' letzte Zeile ermitteln
lz = Cells(Rows.Count, 1).End(xlUp).Row

' nicht benötigte Spalten löschen
Columns("W:AT").Delete
Columns("C").Delete

' neue Spalte für abgesetzte Stunden (negatives Zeitvolumen) einfügen
Columns("K").Insert

' Spaltenüberschriften ändern
Range("K1") = "Minus"
Range("J1") = "Plus"
Range("Q1") = "Nacht"
Range("R1") = "Nacht06"
Range("S1") = "Betrag"
Range("T1:V1").ClearContents

' Tabelle entfärben
Rows("1:120").Interior.ColorIndex = xlNone
Rows("1").Interior.ColorIndex = 15

' Spalte B in echtes Datumsformat wandeln
' Wochenenden und Feiertage markieren
' http://www.office-loesung.de/ftopic472647_0_0_asc.php

Jahr = Year(CDate(Range("B2")))
Neujahr = DateSerial(Jahr, 1, 1)
d = (((255 - 11 * (Jahr Mod 19)) - 21) Mod 30) + 21
Ostersonntag = DateSerial(Jahr, 3, 1) + d + (d > 48) + 6 - ((Jahr + Jahr \ 4 + d + (d > 48) + 1) Mod 7)
Maifeiertag = DateSerial(Jahr, 5, 1)
Karfreitag = Ostersonntag - 2
Ostermontag = Ostersonntag + 1
Himmelfahrt = Ostersonntag + 39
Pfingstsonntag = Ostersonntag + 49
Pfingstmontag = Ostersonntag + 50
TDE = DateSerial(Jahr, 10, 3)
Reformationstag = DateSerial(Jahr, 10, 31)
BuBTag = DateSerial(Jahr, 12, 25) - Weekday(DateSerial(Jahr, 12, 25), vbMonday) - 32
ErsterWeihnachtstag = DateSerial(Jahr, 12, 25)
ZweiterWeihnachtstag = DateSerial(Jahr, 12, 26)

Set Bereich = Range("B2:B" & lz)
For Each Zelle In Bereich
    Zelle.Value = CDate(Zelle.Value)
    Zelle.NumberFormat = "ddd, dd/mm/yy"
    Zelle.HorizontalAlignment = xlRight
    Select Case Weekday(Zelle)
      Case 1
        Zelle.Interior.ColorIndex = 3
      Case 7
        Zelle.Interior.ColorIndex = 45
      Case Else
        Select Case Zelle.Value
            Case Is = Neujahr, Ostersonntag, Maifeiertag, Karfreitag, Ostermontag, Himmelfahrt, Pfingstsonntag, TDE, Reformationstag, BuBTag, ErsterWeihnachtstag, ZweiterWeihnachtstag
                Zelle.Interior.ColorIndex = 3
        End Select
    End Select
Next

' Spalten D-G in echtes Zeitformat wandeln (von, bis)
Set Bereich = Range("D2:G" & lz)
For Each Zelle In Bereich
    Zelle.NumberFormat = "h:mm"
    Zelle.Value = Replace(Zelle.Value, ",", ":")
Next

' Spalten H-J in echtes Zeitformat wandeln (Plan, Ist, +/-)
Set Bereich = Range("H2:J" & lz)
For Each Zelle In Bereich
    Zelle.NumberFormat = "[hh]:mm"
    Zelle.Value = Replace(Zelle.Value, ",", ":")
Next

'negative Zeitvolumen aus Spalte J nach rechts verschieben, Minuszeichen entfernen und in echtes Zeitfomat wandeln
For Each Zelle In Range("J2:J" & lz)
    If Zelle.Value Like "-*" Then
        Zelle.Offset(0, 1).Value = Mid(Zelle.Value, 2)
        Zelle.ClearContents
    End If
Next

' Spalten M-V in echtes Zeitformat wandeln
Set Bereich = Range("M2:V" & lz)
For Each Zelle In Bereich
    Zelle.NumberFormat = "[hh]:mm"
    Zelle.Value = Replace(Zelle.Value, ",", ":")
Next

' Schrift, Rahmen, Spaltenbreite
With Columns("A:V")
    .Font.Name = "Calibri"
    .Font.Size = 11
    .Borders.LineStyle = xlContinuous
    .EntireColumn.AutoFit
End With

ActiveWindow.FreezePanes = False
ActiveWindow.SplitRow = 1
ActiveWindow.FreezePanes = True
Range("C2").Select
Application.ScreenUpdating = True

MsgBox Format(Timer - Laufzeit, "#0.00") & " Sekunden gerödelt!"
End Sub
