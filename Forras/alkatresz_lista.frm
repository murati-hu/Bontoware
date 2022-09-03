VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form alkatresz_lista 
   Caption         =   "Alkatrész lista"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   14895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame vezerlok 
      Caption         =   "Szûrés"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   14655
      Begin VB.TextBox keres 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   1
         Top             =   480
         Width           =   4455
      End
      Begin VB.CommandButton alkhozza 
         Caption         =   "Alkatrész gyors felvétele raktárkészletbe"
         Height          =   615
         Left            =   11040
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox marka 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox focsop 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox tipus 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox alkatresz 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox alcsop 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton frissites 
         Caption         =   "Frissít"
         Default         =   -1  'True
         Height          =   375
         Left            =   9000
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gyorskeresés:"
         Height          =   195
         Index           =   5
         Left            =   4080
         TabIndex        =   17
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Alkatrész:"
         Height          =   195
         Index           =   4
         Left            =   4080
         TabIndex        =   16
         Top             =   840
         Width           =   690
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Alcsoport:"
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   15
         Top             =   840
         Width           =   705
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Fõcsoport:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   750
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Típus:"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   13
         Top             =   240
         Width           =   450
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Gyártmány:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   795
      End
      Begin VB.Label osszesen 
         AutoSize        =   -1  'True
         Caption         =   "Kérem válasszon egy szûrési feltételt!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6720
         TabIndex        =   10
         Top             =   1080
         Width           =   3210
      End
   End
   Begin VB.CommandButton bezar 
      Caption         =   "Bezár"
      Height          =   495
      Left            =   11040
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ListView raktarkeszlet 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   8705
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Azonosító"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cikkszám"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Gyáriszám"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nyszám"
         Object.Width           =   1677
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Gyártmány"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Típus"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fõcsoport"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Alcsoport"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Alkatrész"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Állapot"
         Object.Width           =   884
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Hiányos"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Árkategória"
         Object.Width           =   441
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Ár"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Áfa"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Súly"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Helye"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "alkatresz_lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim jelzo As Integer
Dim all As Byte  '0 ha sima alkatresztabla, 1 ha szamlabol hivva
Dim elem As ListItem

'Dim betoltve As Boolean

Public Sub meghiv(Optional allapot As Byte)
    all = allapot '0 - csak listázás
    If allapot = 1 Then alkhozza.Visible = True
    Form_Initialize
    Me.Show
End Sub

Public Sub Frissit(Optional KSzuro As String)
    
If Not Me.Visible Then Exit Sub
    
    Dim Sor As New ADODB.Recordset
    Dim p As String, Szures1 As String, Szures2 As String, Szures3 As String
    Dim kover As Boolean
    
    raktarkeszlet.Visible = False
    raktarkeszlet.ListItems.Clear
    
    '                   0                       1                    2                  3           4           5              6          7             8                       9                   10                         11                        12                 13              14              15              16              17                  18                     19               20
    p = "SELECT raktarkeszlet.id, raktarkeszlet.cikkszam, raktarkeszlet.gyszam, autok.nyszam, markak.marka, tipusok.tipus, focsop.nev, alcsop.nev, alkatresznevek.nev, raktarkeszlet.allapot, raktarkeszlet.hianyos, autok.arkategoria, raktarkeszlet.ar, raktarkeszlet.afa, raktarkeszlet.suly, raktarkeszlet.megj, focsop.cikkszam, alcsop.cikkszam, alkatresznevek.cikkszam, autok.id, raktarkeszlet.alkatresz, focsop.id, raktarkeszlet.hely " & _
        "FROM ((markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN autok ON tipusok.id = autok.tipus) INNER JOIN (raktarkeszlet INNER JOIN (vtsz INNER JOIN (focsop INNER JOIN (alcsop INNER JOIN alkatresznevek ON alcsop.id = alkatresznevek.alcsop) ON focsop.id = alcsop.focsop) ON vtsz.id = alkatresznevek.vtsz) ON raktarkeszlet.alkatresz = alkatresznevek.id) ON autok.id = raktarkeszlet.auto " & _
        "WHERE (((raktarkeszlet.tipus)=0) AND ((raktarkeszlet.selejt)=False) AND ((raktarkeszlet.sztorno)=False) AND ((raktarkeszlet.elkelt)=False)) "

    'NIncs szûrési feltétel
    osszesen.ForeColor = vbBlack
    If (marka.ListIndex = 0 Or tipus.ListIndex = 0) And (focsop.ListIndex = 0 Or alcsop.ListIndex = 0) And Len(keres.Text) < 4 Then
        osszesen.ForeColor = vbRed
        osszesen.Caption = "Nem adta meg a teljes szûrési feltételt!"
        
        raktarkeszlet.ListItems.Clear
        Exit Sub
    End If
    
    'SZûrés gyártmány és típus alapján
    If tipus.ListIndex = 0 Then
        'Exit Sub
        If marka.ListIndex = 0 Then
            Szures1 = ""
        Else
            Szures1 = " and tipusok.marka=" & marka.ItemData(marka.ListIndex) & " "
        End If
    Else
        Szures1 = " and tipusok.id=" & tipus.ItemData(tipus.ListIndex) & " "
    End If
    
    'Szûrés alkatrészek alapján
    If alkatresz.ListIndex = 0 Then
        If alcsop.ListIndex = 0 Then
            If focsop.ListIndex = 0 Then
                Szures2 = ""
            Else
                Szures2 = " and focsop.id=" & focsop.ItemData(focsop.ListIndex) & " "
            End If
        Else
            Szures2 = " and alcsop.id=" & alcsop.ItemData(alcsop.ListIndex) & " "
        End If
    Else
        Szures2 = " and raktarkeszlet.alkatresz=" & alkatresz.ItemData(alkatresz.ListIndex) & " "
    End If
    
    KSzuro = Trim(KSzuro)
    If KSzuro <> "" Then
        Szures3 = " and (alkatresznevek.nev LIKE '%" & KSzuro & "%' or alcsop.nev LIKE '%" & KSzuro & "%' or autok.evjarat  LIKE '%" & KSzuro & "%' or autok.motorkod  LIKE '%" & KSzuro & "%' or autok.valto  LIKE '%" & KSzuro & "%')"
    End If
    
    
    
    p = p & Szures1 & Szures2 & Szures3
    
    SQL_p p, Sor
    
    osszesen.Caption = "Listázva: " & Sor.RecordCount & " tétel"
    
    If Not Sor.EOF Then Sor.MoveFirst
    Do While Not Sor.EOF
            Set elem = raktarkeszlet.ListItems.Add(, , Sor.Fields(0).Value) 'Azonosito
                elem.ListSubItems.Add , , Nstr(NKieg(Sor.Fields(16).Value) & NKieg(Sor.Fields(17).Value) & NKieg(Sor.Fields(18).Value)) 'Cikkszám
                elem.ListSubItems.Add , , Nstr(Sor.Fields(2).Value) 'Marka
                elem.ListSubItems.Add , , Nstr(Sor.Fields(3).Value) 'Tipus
                elem.ListSubItems.Add , , Nstr(Sor.Fields(4).Value) 'Focsoport
                elem.ListSubItems.Add , , Nstr(Sor.Fields(5).Value) 'Alcsoport
                elem.ListSubItems.Add , , Nstr(Sor.Fields(6).Value) 'Alkatresz
                elem.ListSubItems.Add , , Nstr(Sor.Fields(7).Value) 'Alcsoport
                
                
                Szures1 = Nstr(Sor.Fields(8).Value)
                If (Sor.Fields(20).Value = egyeb.MotorID) Then
                    'Motorkód és szám
                    Szures1 = Szures1 & " - motorkód:" & Ertek("autok", "id", CStr(Sor.Fields(19).Value), "motorkod") & "; motorszám: " & Ertek("autok", "id", CStr(Sor.Fields(19).Value), "motor")
                Else
                If Sor.Fields(18).Value = 1 And Sor.Fields(21).Value = ValtoFcs Then
                    'Váltókód és szám
                    Szures1 = Szures1 & " - váltókód:" & Ertek("autok", "id", CStr(Sor.Fields(19).Value), "valto")
                End If
                End If
                elem.ListSubItems.Add , , Szures1 'Alkatrész
                
                
                
                elem.ListSubItems.Add , , Allapota(Sor.Fields(9).Value) 'Állapot
                
                elem.ListSubItems.Add , , Alakit(Sor.Fields(10).Value, "hiányos", "") 'Hianyos
                elem.ListSubItems.Add , , Chr(Sor.Fields(11).Value) 'Ar
                elem.ListSubItems.Add , , Sor.Fields(12).Value 'Afa
                elem.ListSubItems.Add , , Sor.Fields(13).Value 'megj
                elem.ListSubItems.Add , , Sor.Fields(14).Value 'Afa
                elem.ListSubItems.Add , , Nstr(Sor.Fields(22).Value) 'Hely
                
                
                '15'elem.ListSubItems.Add , , Nstr(Sor.Fields(15).Value) 'megj
                
                'elem.ForeColor = SzinAllapot(Sor.Fields(9).Value - 1)
                
                If Sor.Fields(18).Value = 1 Then kover = True Else kover = False
                If Sor.Fields(9).Value = 1 Then
                    RowColor vbBlack, elem, kover
                Else
                    RowColor &H796EB, elem, kover
                End If
                
        Sor.MoveNext
    Loop
    Sor.Close
    raktarkeszlet.Visible = True
End Sub

Private Sub alkhozza_Click()
    kintrol.Show vbModal
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vezerlok.Move 0, 0, Me.ScaleWidth
    raktarkeszlet.Move 120, vezerlok.Height, Me.ScaleWidth - 2 * raktarkeszlet.Left, Me.ScaleHeight - raktarkeszlet.Top - 100
End Sub

Private Sub frissites_Click()
    Frissit
    Fokusz
End Sub

Private Sub bezar_Click()
    Unload Me
End Sub

'Private Sub Form_Paint()
'    If Not betoltve And Visible Then
'        betoltve = True
'        Frissit
'    End If
'End Sub

Private Sub Form_Initialize()
    'betoltve = False
    jelzo = 0
    Marka_Betolt
    Focsop_Betolt
    'Frissit
    jelzo = 1
End Sub

Private Sub keres_Change()
If Len(keres.Text) > 2 Then
    Frissit keres.Text
Else
    If keres.Text = "" Then Frissit
End If

End Sub

Private Sub raktarkeszlet_DblClick()
    alkatresz_eladasra.Bearaz raktarkeszlet.SelectedItem.Text, 255, all '666 lista
    'Frissit
End Sub

Private Sub alkatresz_Change()
    If jelzo = 1 And alkatresz.ListIndex > 0 Then Frissit
    Fokusz
End Sub

Private Sub alkatresz_Click()
    alkatresz_Change
End Sub

Private Sub marka_Click()
    marka_change
End Sub

Private Sub marka_change()
    Tipus_Betolt
    Fokusz
End Sub

Private Sub focsop_Click()
    focsop_Change
End Sub
Private Sub focsop_Change()
    Alcsop_Betolt
    Fokusz
End Sub

Private Sub alcsop_Click()
    alcsop_Change
End Sub

Private Sub alcsop_Change()
    Alkatresz_Betolt
    Fokusz
End Sub

Private Sub tipus_Click()
    tipus_change
End Sub

Private Sub tipus_change()
    If jelzo = 1 And tipus.ListIndex > 0 Then Frissit
    Fokusz
End Sub

Private Sub Marka_Betolt()
    Betolt Me.marka, "markak", "marka", "marka", "Minden gyártmány"
End Sub

Private Sub Tipus_Betolt()
   Betolt Me.tipus, "tipusok", "tipus", "tipus", "Minden típusa", "where marka=" & marka.ItemData(marka.ListIndex)
End Sub

Private Sub Focsop_Betolt()
    Betolt Me.focsop, "focsop", "nev", "id", "Minden fõcsoport"
End Sub

Private Sub Alcsop_Betolt()
    'If Not Me.Visible Then Exit Sub
    
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
Private Sub Fokusz()
On Error Resume Next
    raktarkeszlet.SetFocus
End Sub

