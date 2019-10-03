Option Explicit
Sub Nachtstunden()
' Berechnung der Nachtstunden 20:00 bis 6:00 Uhr

' Vorspiel
Dim Von1 As Integer, Bis1 As Integer
Dim Von2 As Integer, Bis2 As Integer
Dim Zeile As Integer, lz As Integer
Dim DB1 As Date, DB2 As Date, DE1 As Date, DE2 As Date
Dim Ist As Date, Nacht As Date, NzB As Date, NzE As Date
NzB = "20:00"
NzE = "06:00"
Von1 = 4
Bis1 = 5
Von2 = 6
Bis2 = 7
Range("Q1") = "Nacht"
Range("R1") = "Nacht4"

' letzte Zeile ermitteln
lz = Cells(Rows.Count, 1).End(xlUp).Row

For Zeile = 2 To lz
    If Cells(Zeile, Von1) <> Empty Then
        DB1 = Cells(Zeile, Von1)
        DE1 = Cells(Zeile, Bis1)
        Cells(Zeile, 17).Formula = Max(0, Min(NzE + (NzB > NzE), MOD(DE1, 1) + (MOD(DB1, 1) > MOD(DE1, 1))) - Max(NzB, MOD(DB1, 1))) + Max(0, (Min(NzE, MOD(DE1, 1) + (MOD(DB1, 1) > MOD(DE1, 1))) - MOD(DB1, 1)) * (NzB > NzE)) + Max(0, Min(NzE + (NzB > NzE), MOD(DE1, 1) + 0) - NzB) * (MOD(DB1, 1) > MOD(DE1, 1))
        'Nacht=MAX(;MIN(NzE+(NzB>NzE);REST(DE1;1)+(REST(DB1;1)>REST(DE1;1)))-MAX(NzB;REST(DB1;1)))+MAX(;(MIN(NzE;REST(DE1;1)+(REST(DB1;1)>REST(DE1;1)))-REST(DB1;1))*(NzB>NzE))+MAX(;MIN(NzE+(NzB>NzE);REST(DE1;1)+0)-NzB)*(REST(DB1;1)>REST(DE1;1))
    End If
    If Cells(Zeile, Von2).Value <> Empty Then
        DB2 = Cells(Zeile, Von2)
        DE2 = Cells(Zeile, Bis2)
        Cells(Zeile, 17) = Cells(Zeile, 17) + Cells(Zeile, 17).Formula = Max(0, Min(NzE + (NzB > NzE), REST(DE1, 1) + (REST(DB1, 1) > REST(DE1, 1))) - Max(NzB, REST(DB1, 1))) + Max(0, (Min(NzE, REST(DE1, 1) + (REST(DB1, 1) > REST(DE1, 1))) - REST(DB1, 1)) * (NzB > NzE)) + Max(0, Min(NzE + (NzB > NzE), REST(DE1, 1) + 0) - NzB) * (REST(DB1, 1) > REST(DE1, 1))

    End If

Next

End Sub
