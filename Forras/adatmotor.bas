Attribute VB_Name = "adatmotor"
Option Explicit
Const INIFajl = "adatbazis.ini"
Public DB As String
Public cnn As New ADODB.Connection
Public sql As New ADODB.Command
Public Rekord As New ADODB.Recordset

'Munkak�nyvt�r
Public Function Konyvtar() As String
On Error GoTo Hiba
    Open INIFajl For Input As 100
        Line Input #100, Konyvtar
    Close 100
    
Exit Function
Hiba:
    'Hiba Err.Number
    'MsgBox Konyvtar
    Konyvtar = App.Path & "\"
End Function

'Adatb�tis forr�sa
Public Function Forras() As String
On Error GoTo Hiba
    DB = Konyvtar & "adatok.mdb" 'db
    Forras = DB
Exit Function
Hiba:
    Hiba Err.Number
End Function

Public Sub BezarR(Mit As ADODB.Recordset)
On Error Resume Next
    Mit.Close
End Sub
'Adatb�tis megnyit�sa
Public Sub Megnyitas()
On Error GoTo Hiba
    'jelentesdb.jelenteskapcs.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Forras & ";"
    cnn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Forras & ";"
    Set sql.ActiveConnection = cnn
    'Frissites
Exit Sub
Hiba:
    Hiba Err.Number
End Sub

'R�gi Glob�lis Sql parancsfuttat�
Public Sub sql_parancs(Parancs As String)
    SQL_p Parancs, Rekord
End Sub


Public Function Tagol(Mit As String) As String
On Error GoTo Hiba
    Tagol = Mid(Mit, 1, Len(Mit) - 7) & " " & Mid(Right(Mit, 7), 1, 3) & " " & Right(Mit, 4)
Exit Function
Hiba:
    Tagol = Mit
End Function

'Kombolista felt�lt�se adatt�bl�b�l
Public Sub Betolt(Kombo As ComboBox, Tabla As String, nev As String, Rendezes As String, Optional Elso As String, Optional Where As String, Optional Megjelol As Long)
'On Error GoTo hiba
    Dim Id As Long
    Dim Sor As New ADODB.Recordset
    
    Kombo.Clear
    Kombo.Visible = False
    
    Id = 0
    If Elso <> "" Then
        Kombo.List(0) = Elso
        Kombo.ItemData(0) = -1
        Id = Id + 1
    End If
    
    'SQL_p "SELECT * FROM " & Tabla & " " & Where & " order by " & Rendezes, Sor
    SQL_p "SELECT " & nev & ", id FROM " & Tabla & " " & Where & " order by " & Rendezes, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        Do While Not Sor.EOF
            Kombo.List(Id) = Sor.Fields(0).Value 'Oszlop(Sor, nev) 'Sor!Nev
            Kombo.ItemData(Id) = Sor!Id
            Sor.MoveNext
            Id = Id + 1
        Loop
    End If
    Kombo.Visible = True
    Sor.Close
    If Megjelol = 0 Then
            ElsotJelol Kombo
        Else
            Jelol Kombo, Megjelol
    End If
Exit Sub
Hiba:
    On Error Resume Next
    Sor.Close
    Kombo.Visible = True
    Hiba Err.Number, "Szin Frissit�si hiba"
End Sub

Public Function MelyikGyartmany(Id As Long)
    MelyikGyartmany = Ertek("markak", "id", CStr(Id), "marka")
End Function


Public Function MelyikTipus(Id As Long)
    MelyikTipus = Ertek("tipusok", "id", CStr(Id), "tipus")
End Function


Public Sub SQL_p(Parancs As String, R As Recordset)
On Error Resume Next
    Dim sql2 As New ADODB.Command
    Set sql2.ActiveConnection = cnn
    
    R.CursorLocation = adUseClient
    sql2.CommandText = Parancs
    
    R.Open sql2
    If Err.Number <> 0 Then
        'Call hiba(Err.Number, "SQL_parancsfuttat�")
        MsgBox "Hibajelent�s: " & Err.Description, vbInformation, "SQL2 parancsfuttat� - " & Err.Number
    End If
End Sub
'Meghat�rozott �rt�ket ad vissza az oszlopn�vnek megfelel�en
Public Function Oszlop(Sor As Recordset, Melyiket As String)
    Dim i As Long
    i = Sor.Fields.Count - 1
    Do While i > 0 And Sor.Fields.Item(i).Name <> Melyiket
        i = i - 1
    Loop
    If IsNumeric(Sor.Fields.Item(i).Value) Then
        Oszlop = Sor.Fields.Item(i).Value
    Else
        Oszlop = Nstr(Sor.Fields.Item(i).Value)
    End If
End Function
'Konkr�t �rt�k lek�rdez�se adatt�bl�b�l
Public Function Ertek(Tabla As String, Mi As String, Mivel As String, Vissza As String, Optional egyeb As String)
    Dim Sor As New ADODB.Recordset
    
    SQL_p "SELECT " & Vissza & " from " & Tabla & " where " & Mi & "=" & Mivel, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        'Ertek = Nstr(Oszlop(Sor, Vissza))
        Ertek = Sor.Fields(0).Value
    Else
        Ertek = -1
    End If
    Sor.Close
End Function
'Adatt�bla n�vel�se
Public Sub Novel(Tabla As String, Mi As String, Mivel As String, Mezo As String)
    Dim i As Integer
    Dim Sor As New ADODB.Recordset
    i = (Ertek(Tabla, Mi, Mivel, Mezo)) + 1
    SQL_p "UPDATE " & Tabla & " SET " & Mezo & "=" & i & " WHERE " & Mi & "=" & Mivel, Sor
End Sub

'�j elem felv�tele a tanul� list�kba
Public Sub TanuldMeg(Tabla As String, Mit As String, Optional Mezo As String)
    Dim Sor As New ADODB.Recordset
    If Mezo = "" Then Mezo = "nev"
    
    If Not LetezikIlyen(Tabla, Mezo, "'" & Mit & "'") Then
        SQL_p "INSERT INTO " & Tabla & " (" & Mezo & ") VALUES ('" & Mit & "')", Sor
    End If
End Sub
Public Sub VarosTanul(Varos As String, Telazon As String)
    Dim Sor As New ADODB.Recordset
    
    If Not LetezikIlyen("telepulesek", "id", Telazon) Then
        SQL_p "INSERT INTO telepulesek (id, telepules) VALUES (" & Telazon & ", '" & Varos & "')", Sor
    End If
End Sub
'L�tezik-e m�r a megadott �rt�kkel rekord?0
Public Function LetezikIlyen(Tabla As String, Mezo As String, Ertek As String, Optional PWhere As String)
    Dim Sor As New ADODB.Recordset
    
    SQL_p "SELECT * FROM " & Tabla & " WHERE " & Mezo & "=" & Ertek, Sor
    If Sor.RecordCount = 0 Then LetezikIlyen = False Else LetezikIlyen = True
End Function

'Csak utas�t�s v�grehajt�s
Public Sub FSQL(Parancs As String)
    Dim Fut As New ADODB.Recordset
    SQL_p Parancs, Fut
End Sub

'�ruhullad�kok kivezet�se, visszav�telez�se
Public Sub EladAruHulladek(ewc As Long, DB As Long, megse As Boolean, Optional Kinek As Long)
    Dim Sor As New ADODB.Recordset
    Dim suly As Double
    
    If DB > 0 Then
        If megse Then
            SQL_p "SELECT TOP " & DB & " id FROM raktarkeszlet WHERE ewc=" & ewc & " and tipus=0 and elkelt=FALSE and irany=1 and sztorno=FALSE", Sor
        Else
            SQL_p "SELECT TOP " & DB & " id FROM raktarkeszlet WHERE ewc=" & ewc & " and tipus=0 and elkelt=FALSE and irany=1 and sztorno=FALSE", Sor
        End If
    End If
    
    If Not Sor.EOF Then Sor.MoveFirst
    Do While Not Sor.EOF
        If megse Then
            FSQL "UPDATE raktarkeszlet SET tipus=0 WHERE id=" & Sor!Id
        Else
            FSQL "UPDATE raktarkeszlet SET tipus=3 WHERE id=" & Sor!Id
        End If
        'MsgBox sor!Id
        Debug.Print Sor!Id
        Sor.MoveNext
    Loop
    Sor.Close
End Sub


'�ruhullad�kok darabkil�v�lt�sa
Public Function HulladekDarab(ewc As String, tomeg As Double) As Long
    If Ertek("ewc", "id", ewc, "termek") Then
        HulladekDarab = Round(tomeg / Ertek("ewc", "id", ewc, "szorzo"))
    Else
        HulladekDarab = 0
    End If
End Function
