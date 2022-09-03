VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form auto_lista 
   Caption         =   "Nyilvántartott jármûvek"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   13035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame szures_panel 
      Caption         =   "Szûrés:"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12735
      Begin VB.ComboBox allapotlap 
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox ewc 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox selejtezve 
         Height          =   315
         Left            =   9000
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox hely 
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker ig 
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   45023233
         CurrentDate     =   38580
      End
      Begin MSComCtl2.DTPicker tol 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   45023233
         CurrentDate     =   38580
      End
      Begin VB.CommandButton uj_auto_felv 
         Caption         =   "Új autó"
         Height          =   375
         Left            =   9720
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox keres 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5640
         TabIndex        =   4
         Top             =   480
         Width           =   3975
      End
      Begin VB.ComboBox tipus_lista 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
      Begin VB.ComboBox marka_lista 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label ossz 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selejtezés:"
         Height          =   195
         Left            =   10920
         TabIndex        =   21
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Típus:"
         Height          =   195
         Index           =   8
         Left            =   2880
         TabIndex        =   20
         Top             =   240
         Width           =   450
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gyártmány:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   795
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dátumig"
         Height          =   195
         Index           =   6
         Left            =   1800
         TabIndex        =   18
         Top             =   960
         Width           =   585
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dátumtól:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   675
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Állapotlap:"
         Height          =   195
         Index           =   4
         Left            =   7320
         TabIndex        =   16
         Top             =   960
         Width           =   735
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EWC:"
         Height          =   195
         Index           =   3
         Left            =   5640
         TabIndex        =   14
         Top             =   960
         Width           =   420
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selejtezés:"
         Height          =   195
         Index           =   1
         Left            =   9000
         TabIndex        =   12
         Top             =   960
         Width           =   765
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hely:"
         Height          =   195
         Index           =   2
         Left            =   3720
         TabIndex        =   10
         Top             =   960
         Width           =   360
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gyorskeresés:"
         Height          =   195
         Index           =   0
         Left            =   5640
         TabIndex        =   6
         Top             =   240
         Width           =   1005
      End
   End
   Begin MSComctlLib.ListView auto_lista 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9975
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      NumItems        =   20
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Azonosító"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nyilvántartási szám"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Eredet"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Telepen van-e"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Selejtezve"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "EWC"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Állapotlap"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Gyártmány"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Típus"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Évjárat"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Alvázszám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Motorszám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Motorkód"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Rendszám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Henger"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Hajtóanyag"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Sebváltó"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Eladó"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Bevétel"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Típus"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "auto_lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim elem As ListItem
Public Szures As String
Private Seged As New ADODB.Recordset
Private betoltve As Boolean

Enum ListaModok
    Teljes_lista
    Alvaz_leltar
    Motor_leltar
    Sebvalto_leltar
    NemSzarazrafektetett_lista
End Enum

Private Sub allapotlap_Change()
    Frissit
End Sub

Private Sub allapotlap_Click()
    Frissit
End Sub

Private Sub auto_lista_DblClick()
    adatlap.Megnyit auto_lista.SelectedItem.Text, 30
End Sub

Private Sub Form_Paint()
    If Not betoltve And Visible Then
        betoltve = True
        Frissit
    End If
End Sub

Private Sub uj_auto_felv_Click()
    auto.uj 30
End Sub

Public Sub Frissit()

If Not Me.Visible Then Exit Sub
If Not betoltve Then Exit Sub

'On Error GoTo hiba
    Dim Id As Long, p As String
    Dim Sz As String
    Dim Sor As New ADODB.Recordset
    
    auto_lista.Visible = False
    auto_lista.ListItems.Clear
    
    
    '               0           1           2               3               4               5               6           7               8               9           10          11              12              13              14              15
    p = "SELECT autok.id, autok.nyszam, autok.telepen, autok.selejt, autok.bontva, autok.allapotlap, markak.marka, tipusok.tipus, autok.evjarat, autok.alvaz, autok.motor, autok.motorkod, autok.rendszam, autok.henger, autok.hajtoanyag, [vnev]+' '+[knev] AS Kif1, autok.datum, autok.allapot, autok.eredet, autok.valto " & _
        "FROM partnerek INNER JOIN ((markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN autok ON tipusok.id = autok.tipus) ON partnerek.id = autok.elado "

    'SZûrési feltételek
    Sz = "WHERE (autok.allapot>0) and (autok.datum>=" & DatumAtir(tol.Value) & " and autok.datum<=" & DatumAtir(ig.Value) & ")"
    If keres.Text <> "" Then
        Sz = Sz & " and (alvaz LIKE '%" & keres.Text & "%' or  motor LIKE '%" & keres.Text & "%'  or  motorkod LIKE '%" & keres.Text & "%'  or  nyszam LIKE '%" & keres.Text & "%'  or  datum LIKE '%" & keres.Text & "%' or  hely LIKE '%" & keres.Text & "%' or  bon_szam LIKE '%" & keres.Text & "%' or  rendszam LIKE '%" & keres.Text & "%' or  evjarat LIKE '%" & keres.Text & "%' or  torzskonyv LIKE '%" & keres.Text & "%' or  forgalmi LIKE '%" & keres.Text & "%' or  bon_forg LIKE '%" & keres.Text & "%' or  valto LIKE '%" & keres.Text & "%')"
    End If
    
    Sz = ComboFeltetel(Sz, Me.hely, "autok.telepen")
    Sz = ComboFeltetel(Sz, Me.ewc, "autok.bontva")
    Sz = ComboFeltetel(Sz, Me.selejtezve, "autok.selejt")
    Sz = ComboFeltetel(Sz, Me.allapotlap, "autok.allapotlap")
    
    If tipus_lista.ListIndex = 0 Then
        If marka_lista.ListIndex > 0 Then
            Sz = ComboFeltetel(Sz, Me.marka_lista, "markak.id", True)
        End If
    Else
        Sz = ComboFeltetel(Sz, Me.tipus_lista, "tipusok.id", True)
    End If
    
    
    p = p & Sz
    Debug.Print p
    SQL_p p, Sor
    ossz.Caption = "Összesen " & Sor.RecordCount & " tétel."
    If Not Sor.EOF Then
        Sor.MoveFirst
        Do While Not Sor.EOF
            Set elem = auto_lista.ListItems.Add(, , Sor.Fields(0).Value)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(1).Value)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(18).Value)
                elem.ListSubItems.Add , , Alakit(Nstr(Sor.Fields(2).Value), "Telepen van", "Nincs a telepen")
                elem.ListSubItems.Add , , Alakit(Nstr(Sor.Fields(3).Value), "Selejtezve", "")
                elem.ListSubItems.Add , , Alakit(Nstr(Sor.Fields(4).Value), "160106", "160104*")
                elem.ListSubItems.Add , , Alakit(Nstr(Sor.Fields(5).Value), "Van", "Nincs kitöltve")
                elem.ListSubItems.Add , , Nstr(Sor.Fields(6).Value)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(7).Value)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(8).Value)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(9).Value)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(10).Value)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(11).Value)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(12).Value)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(13).Value)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(14).Value)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(19).Value)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(15).Value)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(16).Value)
                elem.ListSubItems.Add , , AutoTipus(Nstr(Sor.Fields(17).Value))
                
                
                If Sor.Fields(4).Value Then
                
                    If Sor.Fields(3).Value Then
                        RowColor &H808080, elem
                    Else
                        RowColor &H8000&, elem
                    End If
                Else
                    If Sor.Fields(3).Value Then
                        RowColor &H80FF&, elem
                    Else
                        RowColor vbRed, elem
                    End If
                End If
                
                Sor.MoveNext
        Loop
    End If
    auto_lista.Visible = True
    Sor.Close
Exit Sub
Hiba:
    Hiba Err.Number, "Frissitési hiba"
    auto_lista.Visible = True
    Sor.Close
End Sub

Private Sub ewc_Change()
    Frissit
End Sub

Private Sub ewc_Click()
    Frissit
End Sub

Private Sub Form_Initialize()
    betoltve = False
    'tol.Value = "1996.01.01" 'Year(Date) & ".01.31"
    tol.Value = Year(Date) - 1 & ".01.01"
    ig.Value = Year(Date) & ".12.31"
    
    hely.Clear
    hely.AddItem "Mindegy"
    hely.AddItem "Telepen kívül"
    hely.AddItem "Telepen"
    hely.ListIndex = 2
    
    selejtezve.Clear
    selejtezve.AddItem "Mindegy"
    selejtezve.AddItem "Nem selejtezett"
    selejtezve.AddItem "Selejtettek"
    selejtezve.ListIndex = 0
    
    ewc.Clear
    ewc.AddItem "Mindegy"
    ewc.AddItem "160104*"
    ewc.AddItem "160106"
    ewc.ListIndex = 0
    
    allapotlap.Clear
    allapotlap.AddItem "Mindegy"
    allapotlap.AddItem "Nincs kitöltve"
    allapotlap.AddItem "Kitöltve"
    allapotlap.ListIndex = 0
    
    Betolt Me.marka_lista, "markak", "marka", "marka", "Minden Gyártmány"
    'betoltve = True
    Frissit
    
    
End Sub

Private Sub Form_Load()
    Form_Initialize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    szures_panel.Move 120, 120, Me.ScaleWidth - 2 * szures_panel.Left
    auto_lista.Move 0, 1920, Me.ScaleWidth - 2 * auto_lista.Left, Me.ScaleHeight - auto_lista.Top - 120
    
    
End Sub

Private Sub hely_Change()
    Frissit
End Sub

Private Sub hely_Click()
    Frissit
End Sub

Private Sub ig_Change()
    Frissit
End Sub

Private Sub keres_Change()
  Frissit
End Sub

Private Sub marka_lista_Change()
    Betolt Me.tipus_lista, "tipusok", "tipus", "tipus", "Minden típus", "where marka=" & marka_lista.ItemData(marka_lista.ListIndex)
    Frissit
End Sub
Private Sub marka_lista_Click()
    marka_lista_Change
End Sub

Private Sub selejtezve_Change()
    Frissit
End Sub
Public Sub Mutasd(Hogyan As ListaModok)
    Unload Me
    Select Case Hogyan
        Case 1
                '23456
            Me.Caption = "Alvázleltár"
            auto_lista.ColumnHeaders(3).Width = 0.001
            auto_lista.ColumnHeaders(4).Width = 0.001
            auto_lista.ColumnHeaders(5).Width = 0.001
            auto_lista.ColumnHeaders(6).Width = 0.001
            auto_lista.ColumnHeaders(7).Width = 0.001
            auto_lista.ColumnHeaders(12).Width = 0.001
            auto_lista.ColumnHeaders(13).Width = 0.001
            auto_lista.ColumnHeaders(14).Width = 0.001
            auto_lista.ColumnHeaders(15).Width = 0.001
            auto_lista.ColumnHeaders(16).Width = 0.001
            auto_lista.ColumnHeaders(17).Width = 0.001
            auto_lista.ColumnHeaders(18).Width = 0.001
            auto_lista.ColumnHeaders(19).Width = 0.001
            auto_lista.ColumnHeaders(20).Width = 0.001
         Case 2
                '23456
            Me.Caption = "Motorleltár"
            auto_lista.ColumnHeaders(3).Width = 0.001
            auto_lista.ColumnHeaders(4).Width = 0.001
            auto_lista.ColumnHeaders(5).Width = 0.001
            auto_lista.ColumnHeaders(6).Width = 0.001
            auto_lista.ColumnHeaders(7).Width = 0.001
            auto_lista.ColumnHeaders(11).Width = 0.001
            auto_lista.ColumnHeaders(14).Width = 0.001
            auto_lista.ColumnHeaders(15).Width = 0.001
            auto_lista.ColumnHeaders(16).Width = 0.001
            auto_lista.ColumnHeaders(17).Width = 0.001
            auto_lista.ColumnHeaders(18).Width = 0.001
            auto_lista.ColumnHeaders(19).Width = 0.001
            auto_lista.ColumnHeaders(20).Width = 0.001
        Case 3
            Me.Caption = "Sebességváltó leltár"
            auto_lista.ColumnHeaders(3).Width = 0.001
            auto_lista.ColumnHeaders(4).Width = 0.001
            auto_lista.ColumnHeaders(5).Width = 0.001
            auto_lista.ColumnHeaders(6).Width = 0.001
            auto_lista.ColumnHeaders(7).Width = 0.001
            auto_lista.ColumnHeaders(11).Width = 0.001
            auto_lista.ColumnHeaders(12).Width = 0.001
            auto_lista.ColumnHeaders(13).Width = 0.001
            auto_lista.ColumnHeaders(14).Width = 0.001
            auto_lista.ColumnHeaders(15).Width = 0.001
            auto_lista.ColumnHeaders(16).Width = 0.001
            auto_lista.ColumnHeaders(18).Width = 0.001
            auto_lista.ColumnHeaders(19).Width = 0.001
            auto_lista.ColumnHeaders(20).Width = 0.001
        Case 4
            Me.Caption = "Még nem szárazrafektetett gépjármûvek"
            cimke(1).Visible = False
            cimke(3).Visible = False
            cimke(4).Visible = False
            ewc.Visible = False
            selejtezve.Visible = False
            allapotlap.Visible = False
            ewc.ListIndex = 1
    End Select
    Me.Show
End Sub
Private Sub selejtezve_Click()
    Frissit
End Sub

Private Sub tipus_lista_Change()
    Frissit
End Sub

Private Sub tipus_lista_Click()
    Frissit
End Sub

Private Sub tol_Change()
    Frissit
End Sub
