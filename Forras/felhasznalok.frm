VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form felhasznalok 
   Caption         =   "Felhasználók kezelése"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7065
      TabIndex        =   14
      Top             =   0
      Width           =   7095
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Felhasználók kezelése"
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
         Left            =   3840
         TabIndex        =   15
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton stat 
      Caption         =   "Statisztikák mutatása"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton kilep 
      Caption         =   "Kilépés"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton mentes 
      Caption         =   "Adatok mentése"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton uj 
      Caption         =   "Új felhasználó létrehozása"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox jelszo1 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox jelszo 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox nev 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   720
      Width           =   3855
   End
   Begin MSComctlLib.ListView felhasznalok 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Azonosító"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Felhasználónév"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView jogok 
      Height          =   3255
      Left            =   2520
      TabIndex        =   3
      Top             =   2040
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5741
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Biztonsági szint"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Folyamat neve"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Felhasználó hozzáférései:"
      Height          =   195
      Index           =   3
      Left            =   2520
      TabIndex        =   12
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nyilvántartott felhasználók:"
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   1905
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jelszó megerõsítése:"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   1440
      Width           =   1470
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jelszó:"
      Height          =   195
      Index           =   12
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Név:"
      Height          =   195
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   345
   End
End
Attribute VB_Name = "felhasznalok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim uid As Integer


Private Sub felhasznalok_Click()
    Betolt felhasznalok.SelectedItem.Text
End Sub

Private Sub Form_Initialize()
    
    Dim Sor As New ADODB.Recordset
    felso = bontoware.narancs
    uid = 0
    Frissit
End Sub

Private Sub Betolt(Kit As Integer)
    Dim Sor As New ADODB.Recordset
    Dim elem As ListItem
    Dim i As Integer
    
    uid = Kit
    
    For i = 1 To jogok.ListItems.Count
            jogok.ListItems(i).Checked = False
    Next i
            
    SQL_p "SELECT ablak FROM jogok where uid=" & Kit, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        Do While Not Sor.EOF
            For i = 1 To jogok.ListItems.Count
                If jogok.ListItems(i).Text = Sor!Ablak Then
                    jogok.ListItems(i).Checked = True
                End If
            Next i
            Sor.MoveNext
        Loop
    End If
    Sor.Close
    
    SQL_p "SELECT * FROM felhasznalok where id=" & Kit, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        nev.Text = Nstr(Sor!nev)
        jelszo.Text = Nstr(Sor!jelszo)
        jelszo1.Text = Nstr(Sor!jelszo)
    End If
    Sor.Close
    
End Sub

Private Sub Frissit()
    Dim Sor As New ADODB.Recordset
    Dim elem As ListItem
    Dim i As Integer
    
    jogok.ListItems.Clear
    SQL_p "SELECT * FROM ablakok", Sor
    Sor.MoveFirst
    Do While Not Sor.EOF
        Set elem = jogok.ListItems.Add(, , Sor!Id)
            elem.ListSubItems.Add , , Nstr(Sor!nev)
        Sor.MoveNext
    Loop
    Sor.Close
    
    felhasznalok.ListItems.Clear
    SQL_p "SELECT * FROM felhasznalok order by nev", Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        Betolt (Sor!Id)
        Do While Not Sor.EOF
            Set elem = felhasznalok.ListItems.Add(, , Sor!Id)
                elem.ListSubItems.Add , , Nstr(Sor!nev)
            Sor.MoveNext
        Loop
    End If
    Sor.Close
    
    'felhasznalok.SetFocus
End Sub

Private Sub Form_Load()
    Form_Initialize
End Sub

Private Sub jogok_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim Sor As New ADODB.Recordset
    SQL_p "Delete * FROM jogok where uid=" & uid & " and ablak=" & Item.Text, Sor
    If Item.Checked Then SQL_p "insert into jogok (uid, ablak) VALUES (" & uid & "," & Item.Text & ")", Sor
End Sub

Private Sub kilep_Click()
    Unload Me
End Sub

Private Sub mentes_Click()
    Dim Sor As New ADODB.Recordset
    
    If jelszo.Text = jelszo1.Text Then
        SQL_p "UPDATE felhasznalok SET nev='" & nev.Text & "', jelszo='" & jelszo.Text & "' where id=" & uid, Sor
    Else
        MsgBox "A két jelszó nem egyezik meg!"
    End If
End Sub

Private Sub stat_Click()
    felhasznalo_stat.Mutasd felhasznalok.SelectedItem.Text
End Sub

Private Sub uj_Click()
    Dim p As String, Ujnev As String
    Dim Sor As New ADODB.Recordset
    
    Ujnev = InputBox("Kérem adja meg az új felhasználó nevét!", "Új felhasználó")
    Ujnev = Trim(Ujnev)
    
    If Ujnev = "" Then Exit Sub
    
    p = "INSERT INTO felhasznalok (nev) VALUES ('" & Ujnev & "')"
    SQL_p p, Sor
    SQL_p "SELECT id FROM felhasznalok where nev='" & Ujnev & "' order by id desc", Sor
    Sor.MoveFirst
    Betolt (Sor!Id)
    Sor.Close
    
    Frissit
End Sub
