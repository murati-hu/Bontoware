Attribute VB_Name = "partner"
Option Explicit

'Partner tábla moduljai

'Partner adataival feltölti az azonos nevû szovegdobozokat
Public Sub Partner_Adatai(Ablak As Form, Rekord As Recordset)
    On Error Resume Next
    With Ablak
        .vnev.Text = Nstr(Rekord!vnev)
        .knev.Text = Nstr(Rekord!knev)
        .orszag.Text = Nstr(Rekord!orszag)
        .telazon.Text = Nstr(Rekord!telazon)
        .varos.Text = Nstr(Rekord!varos)
        .irszam.Text = Nstr(Rekord!irszam)
        .cim.Text = Nstr(Rekord!cim)
        .ado.Text = Nstr(Rekord!ado)
        .tel.Text = Nstr(Rekord!tel)
        .fax.Text = Nstr(Rekord!fax)
        .email.Text = Nstr(Rekord!email)
        .szemelyi.Text = Nstr(Rekord!szemelyi)
        .kuj.Text = Nstr(Rekord!kuj)
        .ktj.Text = Nstr(Rekord!ktj)
        .allampolg.Text = Nstr(Rekord!allampolg)
        .megj.Text = Nstr(Rekord!megj)
        .vhk.Text = Nstr(Rekord!vhk)
    End With
End Sub

'Betötli a megadott partnert és meghívja annak kitöltõjét
Public Function Partner_Load(Kit As Long, Ablak As Form) As Boolean
    Dim Sor As New ADODB.Recordset
    
    SQL_p "SELECT * from partnerek where id=" & Kit, Sor
    
    If Not Sor.EOF Then
        Sor.MoveFirst
        Partner_Adatai Ablak, Sor
        Partner_Load = True
    Else
        Partner_Load = False
    End If
    
    Sor.Close
End Function

'Betölti egy listába az összes partnert
Public Function Partner_Listaba(Hova As ComboBox, Optional Elso As String)
On Error GoTo Hiba
    Dim i As Long
    Dim Sor As New ADODB.Recordset
    
    Hova.Clear
    i = 0
    If Elso <> "" Then
        i = 1
        Hova.AddItem Elso
    End If
    SQL_p "SELECT * FROM partnerek order by vnev", Sor
    
    If Not Sor.EOF Then Sor.MoveFirst
    Hova.Visible = False
    
    Do While Not Sor.EOF
        Hova.List(i) = Nstr(Sor!vnev & " " & Sor!knev)
        '& " (" & sor!Id & ")")
        Hova.ItemData(i) = Sor!Id
        Sor.MoveNext
        i = i + 1
    Loop
    Hova.Visible = True
    ElsotJelol Hova
    Sor.Close
Exit Function
Hiba:
    Sor.Close
    Hova.Visible = True
    Hiba Err.Number, "Partner Frissitési hiba"
End Function

'Partner törlése
Public Sub Partner_Torol(Id As Long, Hova As Byte)
    sql_parancs ("DELETE FROM partnerek WHERE id=" & Id)
    Visszajelez Hova, Id
End Sub

