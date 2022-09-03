Attribute VB_Name = "autok"
Option Explicit
'Autóhoz tartozó eljárások

'Kaszni tömegének összeszámolása
Public Function KaszniTomege(Melyik As Long) As Double
    Dim sor As New ADODB.Recordset
    Dim Teljes_tomeg As Double
    
    'Autó tömegének lekérdezése
    Teljes_tomeg = Ertek("autok", "id", CStr(Melyik), "tomeg")
    SQL_p "SELECT tipus, ewc, irany, suly, elkelt, sztorno, selejt FROM raktarkeszlet WHERE (tipus=1 or tipus=0 or tipus=3) and auto=" & Melyik, sor
    If Not sor.EOF Then
        sor.MoveFirst
        Do While Not sor.EOF
            Select Case sor!tipus
                Case 0 'Alkatrész
                    If (sor!irany = -1 Or sor!elkelt) And CBool(sor!sztorno) = False Then
                        Teljes_tomeg = Teljes_tomeg - sor!suly
                    Else
                        'Még nem kelt el, valahol itt van a telepen
                        If sor!ewc <> 1 Then
                            Teljes_tomeg = Teljes_tomeg - sor!suly
                        End If
                    End If
                Case 1 'Hulladék
                    'Iránya mindig 1
                    Teljes_tomeg = Teljes_tomeg - sor!suly
                Case 2 'Autó
                    'Az autó régi hulladékrekordja
                Case 3 'Hulladékként értékesített alkatrész
                    Teljes_tomeg = Teljes_tomeg - sor!suly
            End Select
            sor.MoveNext
        Loop
    End If
    KaszniTomege = Teljes_tomeg
End Function
'Fenti mentése
Public Sub MentKaszniTomege(Melyik As Long)
    FSQL "DELETE * FROM raktarkeszlet WHERE tipus=2 and auto=" & Melyik
    FSQL "INSERT into raktarkeszlet (tipus, auto, ewc, suly) VALUES ( 2, " & Melyik & ", " & Alakit(Ertek("autok", "id", CStr(Melyik), "bontva"), "1", "0") & ", " & Vesszotlenito(CStr(KaszniTomege(Melyik))) & ")"
End Sub

'Alkatreszt a kasznibl
Public Sub AlkMentKasznitomege(Melyik As Long)
    Dim auto As Long
    auto = Ertek("raktarkeszlet", "id", CStr(Melyik), "auto")
    If auto > 0 Then
        MentKaszniTomege auto
    End If
End Sub

'Autó leselejtezése
Public Sub SelejtezAuto(Melyiket As Long, Hogy As Boolean)
    FSQL "UPDATE autok SET selejt=" & Alakit(Hogy, "TRUE", "FALSE") & " where id=" & Melyiket
    
    Dim kerd As Byte
    If Hogy Then
        kerd = MsgBox("Kívánja az autó még meglévõ összes alkatrészét is leselejtezni?", vbYesNo, "Alkatrészek selejtezése")
    
        If kerd = vbYes Then
            SelejtezAutoAlkatresz Melyiket, True
        Else
            alkatresz_selejt.Mutasd Melyiket
        End If
    Else
        kerd = MsgBox("Kívánja az autó összes telepen lévõ alkatrészének selejtezését visszavonni?", vbYesNo, "Alkatrészek selejtezése")
    
        If kerd = vbYes Then
            SelejtezAutoAlkatresz Melyiket, False
        End If
    End If
End Sub

'Autó-Alkatrész selejtezése
Public Sub SelejtezAutoAlkatresz(auto As Long, Hogy As Boolean)
    If Hogy Then
        FSQL "UPDATE raktarkeszlet SET selejt=TRUE WHERE tipus=0 and (elkelt=FALSE or sztorno=TRUE) and auto=" & auto
    Else
        FSQL "UPDATE raktarkeszlet SET selejt=FALSE WHERE selejt=TRUE and tipus=0 and (elkelt=FALSE or sztorno=TRUE) and auto=" & auto
    End If
End Sub

'Alkatrész selejtezése
Public Sub Selejtez(Melyiket As Long, Hogy As Boolean)
     FSQL "UPDATE raktarkeszlet SET selejt=" & Alakit(Hogy, "TRUE", "FALSE") & " where id=" & Melyiket
End Sub
