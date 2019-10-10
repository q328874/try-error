Option Explicit
Sub Nachtstunden()
' Berechnung der allgemeinen Nachtstunden 20:00 bis 6:00 Uhr sowie
' der besonderen Nachtstunden nach 0:00 Uhr
' http://www.excelformeln.de/formeln.html?welcher=9
' =MAX(;MIN(B1+(A1>B1);B2+(A2>B2))-MAX(A1;A2))+MAX(;(MIN(B1;B2+(A2>B2))-A2)*(A1>B1))+MAX(;MIN(B1+(A1>B1);B2+0)-A1)*(A2>B2)

' Vorspiel
Dim Zeile As Integer, lz As Integer
Dim DB1 As Date, DB2 As Date, DE1 As Date, DE2 As Date
Dim Nacht20 As Date, Nacht00 As Date, NzB20 As Date, NzB00 As Date, NzE As Date
Dim WsF As WorksheetFunction
Set WsF = Application.WorksheetFunction
NzB20 = "20:00"
NzB00 = "00:00"
NzE = "06:00"

' letzte Zeile ermitteln
lz = Cells(Rows.Count, 1).End(xlUp).Row

For Zeile = 2 To lz
    If Cells(Zeile, 4).Value <> Empty Then
        DB1 = Cells(Zeile, 4).Value
        DE1 = Cells(Zeile, 5).Value
        Nacht20 = WsF.Max(0, WsF.Min(NzE + (NzB20 > NzE), DE1 + (DB1 > DE1)) - WsF.Max(NzB20, DB1)) + WsF.Max(0, (WsF.Min(NzE, DE1 + (DB1 > DE1)) - DB1) * (NzB20 > NzE)) + WsF.Max(0, WsF.Min(NzE + (NzB20 > NzE), DE1 + 0) - NzB20) * (DB1 > DE1)
        Nacht00 = WsF.Max(0, WsF.Min(NzE + (NzB00 > NzE), DE1 + (DB1 > DE1)) - WsF.Max(NzB00, DB1)) + WsF.Max(0, (WsF.Min(NzE, DE1 + (DB1 > DE1)) - DB1) * (NzB00 > NzE)) + WsF.Max(0, WsF.Min(NzE + (NzB00 > NzE), DE1 + 0) - NzB00) * (DB1 > DE1)
    End If
    If Cells(Zeile, 6).Value <> Empty Then
        DB2 = Cells(Zeile, 6).Value
        DE2 = Cells(Zeile, 7).Value
        Nacht20 = Nacht20 + WsF.Max(0, WsF.Min(NzE + (NzB20 > NzE), DE2 + (DB2 > DE2)) - WsF.Max(NzB20, DB2)) + WsF.Max(0, (WsF.Min(NzE, DE2 + (DB2 > DE2)) - DB2) * (NzB20 > NzE)) + WsF.Max(0, WsF.Min(NzE + (NzB20 > NzE), DE2 + 0) - NzB20) * (DB2 > DE2)
        Nacht00 = Nacht00 + WsF.Max(0, WsF.Min(NzE + (NzB00 > NzE), DE2 + (DB2 > DE2)) - WsF.Max(NzB00, DB2)) + WsF.Max(0, (WsF.Min(NzE, DE2 + (DB2 > DE2)) - DB2) * (NzB00 > NzE)) + WsF.Max(0, WsF.Min(NzE + (NzB00 > NzE), DE2 + 0) - NzB00) * (DB2 > DE2)
    End If
    Cells(Zeile, 17)=Nacht20
    Cells(Zeile, 18)=Nacht00
Next
End Sub
