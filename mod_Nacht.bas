Option Explicit
Sub Nachtstunden()
' Berechnung der allgemeinen Nachtstunden 20:00 bis 6:00 Uhr sowie
' der besonderen Nachtstunden nach 0:00 Uhr
' http://www.excelformeln.de/formeln.html?welcher=9
' =MAX(;MIN(B1+(A1>B1);B2+(A2>B2))-MAX(A1;A2))+MAX(;(MIN(B1;B2+(A2>B2))-A2)*(A1>B1))+MAX(;MIN(B1+(A1>B1);B2+0)-A1)*(A2>B2)

' Vorspiel
Dim Von1 As Integer, Bis1 As Integer
Dim Von2 As Integer, Bis2 As Integer
Dim Zeile As Integer, lz As Integer
Dim DB1 As Date, DB2 As Date, DE1 As Date, DE2 As Date
Dim Ist As Date, Nacht As Date, NzB20 As Date, NzB00 as Date, NzE As Date
NzB20 = "20:00"
NzB00= "00:00"
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
        Cells(Zeile, 17).Formula = MAX(;MIN(NzE+(NzB20>NzE);DE1+(DB1>DE1))-MAX(NzB20;DB1))+MAX(;(MIN(NzE;DE1+(DB1>DE1))-DB1)*(NzB20>NzE))+MAX(;MIN(NzE+(NzB20>NzE);DE1+0)-NzB20)*(DB1>DE1)
        Cells(Zeile, 18).Formula = MAX(;MIN(NzE+(NzB00>NzE);DE1+(DB1>DE1))-MAX(NzB00;DB1))+MAX(;(MIN(NzE;DE1+(DB1>DE1))-DB1)*(NzB00>NzE))+MAX(;MIN(NzE+(NzB00>NzE);DE1+0)-NzB00)*(DB1>DE1)
    End If
    If Cells(Zeile, Von2).Value <> Empty Then
        DB2 = Cells(Zeile, Von2)
        DE2 = Cells(Zeile, Bis2)
        Cells(Zeile, 17) = Cells(Zeile, 17) + Cells(Zeile, 17).Formula = MAX(;MIN(NzE+(NzB20>NzE);DE2+(DB2>DE2))-MAX(NzB20;DB2))+MAX(;(MIN(NzE;DE2+(DB2>DE2))-DB2)*(NzB20>NzE))+MAX(;MIN(NzE+(NzB20>NzE);DE2+0)-NzB20)*(DB2>DE2)
        Cells(Zeile, 18) = Cells(Zeile, 18) + Cells(Zeile, 18).Formula = MAX(;MIN(NzE+(NzB00>NzE);DE2+(DB2>DE2))-MAX(NzB00;DB2))+MAX(;(MIN(NzE;DE2+(DB2>DE2))-DB2)*(NzB00>NzE))+MAX(;MIN(NzE+(NzB00>NzE);DE2+0)-NzB00)*(DB2>DE2)
    End If
Next
End Sub
