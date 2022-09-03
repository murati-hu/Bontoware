VERSION 5.00
Begin VB.Form kintrol 
   Caption         =   "Alkatresz eladása közvetlen autóból"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Állapotlap bõvítés"
      Height          =   375
      Left            =   4800
      TabIndex        =   22
      Top             =   720
      Width           =   1935
   End
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6945
      TabIndex        =   20
      Top             =   0
      Width           =   6975
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alkatrész eladása autóból"
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
         Left            =   3240
         TabIndex        =   21
         Top             =   240
         Width           =   3570
      End
   End
   Begin VB.CommandButton bezar 
      Caption         =   "Bezár"
      Height          =   375
      Left            =   5160
      TabIndex        =   19
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton hozzaad 
      Caption         =   "Hozzáad"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox megj 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   2760
      Width           =   6495
   End
   Begin VB.ComboBox minosites 
      Height          =   315
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox allapot 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox afa 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox ar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox tomeg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CheckBox hianyos 
      Alignment       =   1  'Right Justify
      Caption         =   "H"
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox alkatresz 
      Height          =   315
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.ComboBox alcsop 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ComboBox focsop 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ComboBox nyszam 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label tipus 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2040
      TabIndex        =   11
      Top             =   720
      Width           =   45
   End
   Begin VB.Label Label6 
      Caption         =   "Megjegyzés:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Minõsítés:"
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Állapot"
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Áfa:"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Nettó ár:"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Tömeg:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
End
Attribute VB_Name = "kintrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bezar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    bovites.Show vbModal
End Sub

Private Sub focsop_Click()
    focsop_Change
End Sub
Private Sub focsop_Change()
    Alcsop_Betolt
End Sub

Private Sub alcsop_Click()
    alcsop_Change
End Sub

Private Sub alcsop_Change()
    Alkatresz_Betolt
End Sub

Private Sub hozzaad_Click()
    If nyszam.ItemData(nyszam.ListIndex) = 0 Then
        MsgBox "Csak nyilvántartási számhoz rendelhet alkatrészt!"
        Exit Sub
    End If
    Dim Sor As New ADODB.Recordset, p As String, Sor2 As New ADODB.Recordset
    'Rektárkészletbe veszem
    SQL_p "SELECT * FROM alkatresznevek where id=" & alkatresz.ItemData(alkatresz.ListIndex), Sor2
    If Not Sor2.EOF Then
        Sor2.MoveFirst
        p = "INSERT into raktarkeszlet (tipus, alkatresz, auto, allapot, hianyos, suly, elkelt, megj, ar, afa, minosites, ewc) VALUES (0, " & alkatresz.ItemData(alkatresz.ListIndex) & ", " & nyszam.ItemData(nyszam.ListIndex) & " , " & allapot.ListIndex & ", " & KonvertalLogikai(hianyos.Value) & ", " & Vesszotlenito(tomeg.Text) & " , FALSE, '" & megj.Text & "', " & Vesszotlenito(ar.Text) & ", " & afa.List(afa.ListIndex) & ", " & minosites.ListIndex & ", " & Sor2!ewc & ")"
        SQL_p p, Sor
    End If
    Sor2.Close
    'Beszúrom a számlára
    p = "SELECT TOP 1 raktarkeszlet.id From raktarkeszlet ORDER BY raktarkeszlet.id DESC"
    SQL_p p, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        szamlazo.Beszur Sor!Id
    End If
    Sor.Close
    'Törlöm az ablakot
    ElsotJelol Me.nyszam
    ElsotJelol Me.focsop
    ElsotJelol Me.afa
    ElsotJelol Me.minosites
    allapot.Text = allapot.List(1)
    ar.Text = ""
    tomeg.Text = ""
    megj.Text = ""
    hianyos.Value = False
    
    Unload Me
End Sub

Private Sub nyszam_Click()
    nyszam_Change
End Sub

Private Sub nyszam_Change()
    Dim Sor As New ADODB.Recordset, p As String
    p = "SELECT markak.marka, tipusok.tipus, autok.id "
    p = p & "FROM (markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN autok ON tipusok.id = autok.tipus "
    p = p & "WHERE (((autok.id)=" & nyszam.ItemData(nyszam.ListIndex) & "))"
    SQL_p p, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        tipus.Caption = Sor.Fields(0) & " " & Sor.Fields(1)
    Else
        tipus.Caption = ""
    End If
    Sor.Close
End Sub

Private Sub Focsop_Betolt()
    Dim Sor As New ADODB.Recordset, i As Integer
    SQL_p "SELECT * FROM focsop order by id", Sor
    If Not Sor.EOF Then Sor.MoveFirst
    focsop.Clear
    focsop.List(0) = "Fõcsoport"
    focsop.ItemData(0) = 0
    i = 1
    Do While Not Sor.EOF
        focsop.List(i) = Sor!nev
        focsop.ItemData(i) = Sor!Id
        i = i + 1
        Sor.MoveNext
    Loop
    'focsop.ListIndex = 0
    ElsotJelol Me.focsop
    Sor.Close
End Sub

Private Sub Alcsop_Betolt()
    Dim Sor As New ADODB.Recordset, i As Integer
    SQL_p "SELECT * FROM alcsop where focsop=" & focsop.ItemData(focsop.ListIndex) & " order by id", Sor
    If Not Sor.EOF Then Sor.MoveFirst
    alcsop.Clear
    alcsop.List(0) = "Minden Alcsoport"
    alcsop.ItemData(0) = 0
    i = 1
    Do While Not Sor.EOF
        alcsop.List(i) = Sor!nev
        alcsop.ItemData(i) = Sor!Id
        i = i + 1
        Sor.MoveNext
    Loop
    'alkatresz.ListIndex = 0
    ElsotJelol Me.alcsop
    Sor.Close
End Sub

Private Sub Alkatresz_Betolt()
    Dim Sor As New ADODB.Recordset, i As Integer
    SQL_p "SELECT * FROM alkatresznevek where alcsop=" & alcsop.ItemData(alcsop.ListIndex) & " order by id", Sor
    If Not Sor.EOF Then Sor.MoveFirst
    alkatresz.Clear
    alkatresz.List(0) = "Minden Alkatrész"
    alkatresz.ItemData(0) = 0
    i = 1
    Do While Not Sor.EOF
        alkatresz.List(i) = Sor!nev
        alkatresz.ItemData(i) = Sor!Id
        i = i + 1
        Sor.MoveNext
    Loop
    'alkatresz.ListIndex = 0
    ElsotJelol Me.alkatresz
    Sor.Close
End Sub

Private Sub Nyszam_betolt()
    Dim Sor As New ADODB.Recordset, i As Integer
    SQL_p "SELECT autok.nyszam, autok.id, autok.selejt, autok.bontva From autok WHERE ((autok.selejt)=False) order by nyszam", Sor

    If Not Sor.EOF Then Sor.MoveFirst
    nyszam.Clear
    nyszam.List(0) = "Nincs szárazrafektetett autó"
    nyszam.ItemData(0) = 0
    i = 0
    Do While Not Sor.EOF
        nyszam.List(i) = Sor!nyszam
        nyszam.ItemData(i) = Sor!Id
        i = i + 1
        Sor.MoveNext
    Loop
    ElsotJelol Me.nyszam
    Sor.Close
End Sub

Private Sub Allapot_Kitolt()
    allapot.List(0) = "nincs"
    allapot.ItemData(0) = 0
    allapot.List(1) = "ép"
    allapot.ItemData(1) = 1
    allapot.List(2) = "sérült"
    allapot.ItemData(2) = 2
    allapot.Text = allapot.List(1)
End Sub

Private Sub Form_Load()
    felso = bontoware.zold
    Focsop_Betolt
    Nyszam_betolt
    Betolt Me.afa, "afa", "afa", "afa"
    Szin_Betolt Me.minosites
    Allapot_Kitolt
End Sub
