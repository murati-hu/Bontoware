Attribute VB_Name = "egyeb"
Option Explicit

Public Fid As Integer
Public Const MotorID = 243
Public Const ValtoFcs = 7
Public Const GumiID = 359
Public Const AkksiID = 338
Public Const ProgramNeve = "BontoWare Beta 1"


'Egy�b f�ggv�nyek

'T�rtek vessz�tlen�t�se
Public Function Vesszotlenito(Mit As String) As String
    Dim i As Long
    i = InStr(1, Mit, ",")
    If i = 0 Then
        Vesszotlenito = Mit
    Else
        Vesszotlenito = Mid(Mit, 1, i - 1) & "." & Mid(Mit, i + 1)
    End If
    If Trim(Vesszotlenito) = "" Then Vesszotlenito = 0
End Function

'Logikai �rt�k �ltal v�laszt sz�veget
Public Function Alakit(Mit As Boolean, Optional Igaz As String, Optional Hamis As String) As String
    'If Igaz = "" Then Igaz = "TRUE"
    'If Hamis = "" Then Hamis = "FALSE"
    If Mit Then Alakit = Igaz Else Alakit = Hamis
End Function

'El�re be�r� motor
Public Sub Kiegeszit(Mibe As ComboBox)
Dim i As Long, h As Byte
    h = Len(Mibe.Text)
    For i = 0 To Mibe.ListCount - 1
        If Mid(Mibe.List(i), 1, Len(Mibe.Text)) = Mibe.Text Then
            Mibe.Text = Mibe.List(i)
            Mibe.SelStart = h
            Mibe.SelLength = Len(Mibe.Text) - h
        End If
    Next i
End Sub

'Visszajelz�si rendszer
Public Sub Visszajelez(Kinek As Byte, Optional Mit As Long)
    Select Case Kinek
        Case 10 'Partner lista
                partner_lista.Frissit (partner_lista.Szures)
        Case 20 'Aut� felv�tele
                auto.BeszurPartner Mit
        Case 21 'M�rka lista
                auto.Marka_Frissit
        Case 22 'Nyomtat�si fel�let
                
        Case 30 'Aut� lista frissit
                auto_lista.Frissit
        Case 40 'Sz�ml�n l�v� partner friss�t�se
                szamlazo.BeszurPartner Mit
        Case 50 '�j vev� besz�r�sa
                hulladek_szamla.BeszurPartner 0, Mit
        Case 51 '�j sz�ll�t� besz�r�sa
                hulladek_szamla.BeszurPartner 1, Mit
        Case 60 '�j elad� besz�r�sa �j alkatr�szn�l
                alkatresz_lap.BeszurPartner Mit
        Case 70 '�j elad� besz�r�sa import aut�hoz
                felauto_lap.BeszurPartner Mit
        Case 80 'Kimeneti havi
    End Select
End Sub

'NULL stringek kezel�se
Public Function Nstr(Mit) As String
On Error GoTo Hiba
    Nstr = CStr(Mit)
Exit Function
Hiba:
    Nstr = ""
End Function

'Saj�t hibakezel� f�ggv�ny
Public Sub Hiba(Kod As Long, Optional cim As String)
    Select Case Kod
        Case 53
            MsgBox "Az adatbazis.ini nem tal�lhat�. K�rem hozzon l�tre egy �jat az adatb�zis el�r�s�vel." & vbCrLf & "A program fut�sa most megszakad.", vbCritical, cim
            End
        Case -2147467259
            MsgBox "Az adatb�zis nem tal�lhat� az adatbazis.ini-ben megadott helyen." & vbCrLf & "A program fut�sa most megszakad.", vbCritical, cim
            End
        Case Else
            Err.Number = Kod
            MsgBox Err.Description
    End Select
    
    Open "hiba.log" For Append As 3
        Print #3, Date & "-" & Time() & " - " & Kod
    Close 3
End Sub

'List�k legels� elem�t jel�li
Public Sub ElsotJelol(Miben As ComboBox)
On Error Resume Next
    Miben.ListIndex = 0
End Sub
'Ideiglenes f�jl
Public Function TmpGeneral(Optional Sablon As String) As String
    TmpGeneral = Hour(Time) & Minute(Time) & Second(Time) & Int(Rnd(Time) * 100) & ".tmp"
End Function
'ADott tartalm� lista jel�l�se
Public Sub Jelol(Miben As ComboBox, Id As Long)
On Error GoTo Hiba
    Dim i As Long
    i = 0
    Do While Miben.ItemData(i) <> Id And i < Miben.ListCount - 1
        i = i + 1
    Loop
    Miben.Text = Miben.List(i)
Hiba:
End Sub
'D�tum �t�r�
Public Function DatumAtir(datum As Date) As String
    DatumAtir = "#" & Month(datum) & "/" & Day(datum) & "/" & Year(datum) & "#"
End Function

'�vj�rat felt�lt�se
Public Sub EvjaratListaba(Hova As ComboBox)
    Dim i As Integer
    Hova.Clear
    Hova.Visible = False
    For i = 1950 To Year(Date)
        Hova.AddItem i
    Next i
    Hova.Visible = True
    Hova.ListIndex = Hova.ListCount - 1
End Sub

'Nyilv�ntart�si sz�m aj�nl�
Public Function NySzamAjanlo(tipus As String) As String
    NySzamAjanlo = (Ertek("bonto", "id", "1", "szam_a") + 1) & "/" & Year(Date)
End Function

'Sz�vegdoboz lokkol�sa
Public Sub Lokkol(Szoveg As TextBox, Optional Engedelyez As Boolean, Optional Enable As Boolean)
    If Enable Then Szoveg.Enabled = Engedelyez
    If Not Engedelyez Then
        Szoveg.BackColor = vbButtonFace
    Else
        Szoveg.BackColor = vbWhite
    End If
End Sub

'�j sor hozz�f�z�se VBhez
Public Function Ujsor(Mihez As String, Mit As String) As String
    Ujsor = Mihez & vbCrLf & Mit
End Function

'Sz�nek bet�lt�se
Public Sub Szin_Betolt(szin As ComboBox, Optional Jelol As Byte)
    szin.Clear
    szin.List(0) = "piros"
    szin.List(1) = "k�k"
    szin.ListIndex = Jelol
End Sub

'Felhaszn�l� bel�ptet�se
Public Sub Beleptet()
    belepes.Show vbModal
End Sub

'Jogosults�g ellen�rz�se
Public Function Jogos(Folyamat As Integer, Optional uzenet As Boolean) As Boolean
    Dim Sor As New ADODB.Recordset
    SQL_p "SELECT * FROM jogok WHERE uid=" & Fid & " and ablak=" & Folyamat, Sor
    If Sor.RecordCount = 0 Then
        If uzenet = True Then MsgBox "�nnek nincs jogosults�ga ehhez az folyamathoz!", vbCritical, Ertek("felhasznalok", "id", CStr(Fid), "nev")
        'Unload Ablak
        Jogos = False
    Else
        Jogos = True
    End If
    Sor.Close
End Function

'Logikai �rt�kek konvert�l�sa
Public Function KonvertalLogikai(Ertek) As String
    If Ertek = True Then
        KonvertalLogikai = "TRUE"
        Exit Function
    End If
    
    If Ertek = False Then
        KonvertalLogikai = "FALSE"
        Exit Function
    End If
    
    KonvertalLogikai = Nstr(Ertek)
End Function

'Jelenleg bejelentkezett felhaszn�l�
Public Function Bejelentkezve() As String
    Bejelentkezve = Ertek("felhasznalok", "id", CStr(Fid), "nev")
End Function

'Ablak neve
Public Function AblakNeve(Ablak As Form)
    AblakNeve = Ertek("ablakok", "id", CStr(Ablak.Tag), "nev")
End Function
'Ablakok megnevez�se
Public Sub ElnevezAblak(Ablak As Form, Optional Szoveg As String)
    Ablak.Caption = AblakNeve(Ablak) & " " & Szoveg & " - " & Bejelentkezve
End Sub

'Soronk�nti besz�r�s lablekbe
Public Sub LUzenet(Hova As Label, Mit As String)
    Hova.Caption = Hova.Caption & Mit & vbCrLf
End Sub

'�llapotok lek�rdez�se
Public Function Allapota(Mi As Byte) As String
    Select Case Mi
        Case 0
            Allapota = "nincs"
        Case 1
            Allapota = "�p"
        Case 2
            Allapota = "s�r�lt"
    End Select
End Function
'SZ�n�llapotok
Public Function SzinAllapot(Mi As Byte) As ColorConstants
    Select Case Mi
        Case 0
            SzinAllapot = vbRed
        Case 1
            SzinAllapot = &H4000&
        Case 2
            SzinAllapot = &H796EB
    End Select
End Function
'Combofelt�telszab�s
Public Function ComboFeltetel(FeltetelSor As String, Kombo As ComboBox, Mezo As String, Optional Adatos As Boolean)
    If Kombo.ListIndex > 0 Then
        If FeltetelSor <> "" Then FeltetelSor = FeltetelSor & " AND "
        If Adatos Then
            FeltetelSor = FeltetelSor & "(" & Mezo & "=" & Kombo.ItemData(Kombo.ListIndex) & ")"
        Else
            FeltetelSor = FeltetelSor & "(" & Mezo & "=" & Alakit(Kombo.ListIndex - 1, "TRUE", "FALSE") & ")"
        End If
    End If
    ComboFeltetel = FeltetelSor
End Function
'Min�s�t�s
Public Function MinositesTipus(Mi As Integer) As String
    Select Case Mi
        Case 0
            MinositesTipus = "piros"
        Case 1
            MinositesTipus = "k�k"
    End Select
End Function
'Aut�k t�pusa
Public Function AutoTipus(Mi As Integer) As String
    Select Case Mi
        Case 0
            AutoTipus = "Csak 1 alkatr�sz"
        Case 1
            AutoTipus = "Bont�sos"
        Case 2
            AutoTipus = "Import bont�sos"
    End Select
End Function

'Sz�m eg�sz�t�se �s �rt�kel�se ixel�s
Public Function IxElem(Ix As Byte, Hanyadik As Byte) As Boolean
    Dim p As String
    p = CStr(Ix)
    If Len(p) = 1 Then
        p = "00" & p
    Else
        If Len(p) = 2 Then p = "0" & p
    End If
    
    If Mid(p, Hanyadik, 1) <> 0 Then IxElem = True Else IxElem = False
End Function
Public Sub RowColor(Color As ColorConstants, Item As ListItem, Optional Bold As Boolean)
    Dim i As Byte
    For i = 1 To Item.ListSubItems.Count
        Item.ListSubItems(i).ForeColor = Color
        Item.ListSubItems(i).Bold = Bold
    Next i
End Sub
'Fejleszt�s
Public Sub Fejl()
    MsgBox "Ez a funci� jelenleg fejleszt�s alatt �ll!", vbInformation, "Fejleszt�s alatt"
End Sub

Public Sub Kozepre(Mit, Minek)
    Mit.Move (Minek.Width - Mit.Width) / 2
End Sub
Public Function NKieg(Szoveg As String, Optional Ures As String, Optional hely As Byte) As String
    If hely = 0 Then hely = 2
    If Ures = "" Then Ures = "0"
    If Len(Szoveg) <= hely Then
        Dim i As Byte
        For i = 1 To hely - Len(Szoveg)
            NKieg = NKieg & Ures
        Next i
        NKieg = NKieg & Szoveg
    Else
        NKieg = Szoveg
    End If
End Function

'Alkatr�szek el�k�sz�t�se elad�s el�tt
'Public Function KomplettEladas(Melyik As Long)
'    Dim sor As New ADODB.Recordset
'    Dim Auto As Long, Alkatresz As Long
'    Dim Alcsop As Long, Cikkszam As Long
''
'    Alkatresz = Ertek("raktarkeszlet", "id", CStr(Melyik), "alkatresz")
'    Auto = Ertek("raktarkeszlet", "id", CStr(Melyik), "auto")
'
'    Cikkszam = Ertek("alkatresznevek", "id", CStr(Alkatresz), "cikkszam")
'
'    If Cikkszam = 1 Then
'        'Komplett --> �sszes t�bbi 4-es t�pus
'        Alcsop = Ertek("raktarkeszlet", "id", CStr(Melyik), "alkatresz")
'    Else
'        'Nem komplett -> Komplett t�rl�se
'        FSQL "UPDATE raktarkeszlet SET tipus=4 WHERE auto= and "
'    End If
'End Function
