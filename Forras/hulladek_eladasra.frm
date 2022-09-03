VERSION 5.00
Begin VB.Form hulladek_eladasra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hulladékok eladásra"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton auto_selejtezes 
      Caption         =   "Autók selejtezése"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox hasznalat 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Text            =   "hasznalat"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6225
      TabIndex        =   9
      Top             =   0
      Width           =   6255
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hulladék eladása"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   2355
      End
   End
   Begin VB.CommandButton mentes 
      Caption         =   "Hozzáad"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton bezar 
      Caption         =   "Mégse"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ComboBox ewc_lista 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   5295
   End
   Begin VB.ComboBox afa 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox suly 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "0"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox ar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      TabIndex        =   4
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "További hasznosítás módja:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   1980
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EWC:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   420
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tömeg:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Egységár kg-ként:"
      Height          =   195
      Index           =   0
      Left            =   3960
      TabIndex        =   11
      Top             =   1200
      Width           =   1290
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Áfakulcs:"
      Height          =   195
      Index           =   12
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   660
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   5640
      X2              =   5520
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Label osszsuly 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kg"
      Height          =   195
      Left            =   5760
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "hulladek_eladasra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim all As Byte
Dim SID As Long
Dim SzamlaSzam As Long

Private Sub auto_selejtezes_Click()
    auto_selejt.Show vbModal
End Sub

Private Sub bezar_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
    felso = bontoware.zold
    all = 0
    SID = 0
    Hulladek_Frissit
    Betolt Me.afa, "afa", "afa", "afa"
    
    hasznalat.Clear
    hasznalat.AddItem "Bálázó"
    hasznalat.AddItem "Shredder"
    hasznalat.AddItem "Újrahasznosítás"
    hasznalat.AddItem "Megsemmisítés"
End Sub

Private Sub Hulladek_Frissit()
    Dim Sor As New ADODB.Recordset
    Dim p As String, i As Integer
    
    p = "SELECT * FROM ewc where id>0 order by ewc"
    
    SQL_p p, Sor
    If Not Sor.EOF Then
        i = 0
        Sor.MoveFirst
        Do While Not Sor.EOF
            ewc_lista.List(i) = Sor!ewc & Alakit(Sor!veszelyes, "*", "") & " - " & Sor!nev
            ewc_lista.ItemData(i) = Sor!Id
            i = i + 1
            Sor.MoveNext
        Loop
        Sor.Close
    End If
    ElsotJelol Me.ewc_lista
    
End Sub

Private Sub ewc_lista_Change()
    Dim Sor As New ADODB.Recordset
    Dim p As String
    
    p = "SELECT ewc.id, Sum([raktarkeszlet]![suly]*[raktarkeszlet]![irany]) AS SumOfsuly " & _
        "FROM ewc INNER JOIN raktarkeszlet ON ewc.id = raktarkeszlet.ewc " & _
        "Where ((((raktarkeszlet.tipus) = 0 and (raktarkeszlet.elkelt=FALSE and raktarkeszlet.sztorno=FALSE)) Or (raktarkeszlet.tipus) = 1 Or (raktarkeszlet.tipus) = 2)) " & _
        "GROUP BY ewc.id " & _
        "HAVING (((ewc.id)=" & ewc_lista.ItemData(ewc_lista.ListIndex) & "))"
    
    'MsgBox p
    SQL_p p, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        osszsuly.Caption = Sor.Fields(1).Value
        
    Else
        osszsuly.Caption = 0
    End If
    Sor.Close
    
    'If ewc_lista.ItemData(ewc_lista.ListIndex) = 1 Then
        auto_selejtezes.Visible = CBool(ewc_lista.ItemData(ewc_lista.ListIndex) = 1)
End Sub

Private Sub ewc_lista_Click()
    ewc_lista_Change
End Sub

Public Sub uj(szamla As Long)
    Form_Initialize
    SzamlaSzam = szamla
    Me.Show vbModal, hulladek_szamla
End Sub

Private Sub mentes_Click()
    Dim Sor As New ADODB.Recordset
    Dim HanyDB As Long
    Dim Aru As Boolean
    
    Aru = Ertek("ewc", "id", ewc_lista.ItemData(ewc_lista.ListIndex), "termek")
    HanyDB = HulladekDarab(ewc_lista.ItemData(ewc_lista.ListIndex), suly.Text)
    'MsgBox HanyDB
    If all = 0 Then
    
        FSQL "INSERT INTO raktarkeszlet (tipus, auto, ewc, gyszam, irany, suly, ar, afa, megj) VALUES (1, -1, " & ewc_lista.ItemData(ewc_lista.ListIndex) & ", '" & hasznalat.Text & "',-1, " & Vesszotlenito(suly.Text) & ", " & Vesszotlenito(ar.Text) & ", " & afa.Text & ", '" & HanyDB & "')"
        SQL_p "SELECT id FROM raktarkeszlet WHERE auto=-1 and ewc=" & ewc_lista.ItemData(ewc_lista.ListIndex) & " and irany=-1 and suly=" & Vesszotlenito(suly.Text) & " and ar=" & Vesszotlenito(ar.Text) & " and afa=" & afa.Text & " ORDER BY id desc", Sor
        Sor.MoveFirst
        SID = Sor!Id
        Sor.Close
        FSQL "INSERT INTO elkelt (tetel, szamla) VALUES (" & SID & ", " & SzamlaSzam & " )"
        
        'Áruk lefoglalása
        If Aru Then EladAruHulladek ewc_lista.ItemData(ewc_lista.ListIndex), HanyDB, False, SID
    Else
        'Régi Áruk felszabadítása
        '- Ha a régi áru volt
        If Ertek("ewc", "id", Ertek("raktarkeszlet", "id", CStr(SID), "ewc"), "termek") = True Then
            EladAruHulladek ewc_lista.ItemData(ewc_lista.ListIndex), Ertek("raktarkeszlet", "id", CStr(SID), "megj"), True, SID
        End If
        
        'Újak lefoglalása
        If Aru Then EladAruHulladek ewc_lista.ItemData(ewc_lista.ListIndex), HanyDB, False, SID
        
        SQL_p "UPDATE raktarkeszlet SET megj='" & HanyDB & "', ewc=" & ewc_lista.ItemData(ewc_lista.ListIndex) & ", gyszam='" & hasznalat.Text & "', suly=" & Vesszotlenito(suly.Text) & " , ar=" & Vesszotlenito(ar.Text) & ", afa=" & afa.Text & " WHERE id=" & SID, Sor
    End If
    'Áruhulladéknál elkelt teszi láthatatlanná a darabszámokat
    
    hulladek_szamla.Frissit
    Unload Me
End Sub

Private Sub suly_Change()
    'If suly.Text > osszsuly.Caption Then
    '    suly.Text = osszsuly.Caption
    'End If
End Sub
Public Sub modosit(hulladek As Long, szamla As Long)
    Dim Sor As New ADODB.Recordset
    Form_Initialize
    all = 1
    SzamlaSzam = szamla
    SID = hulladek
    SQL_p "SELECT * FROM raktarkeszlet WHERE id=" & SID, Sor
    
    Jelol Me.ewc_lista, Sor!ewc
    'suly.Text = Abs(Nstr(Sor!suly))
    suly.Text = Abs(Sor!suly)
    afa.Text = Nstr(Sor!afa)
    ar.Text = Nstr(Sor!ar)
    hasznalat.Text = Nstr(Sor!gyszam)
    Me.Show vbModal, hulladek_szamla
End Sub
