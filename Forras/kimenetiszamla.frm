VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form hulladek_szamla 
   Caption         =   "Számlázás hulladékra"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   12840
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton felhasznalovaltas 
      Caption         =   "Felhasználó váltás"
      Height          =   615
      Left            =   10440
      TabIndex        =   46
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton uj_szamla 
      Caption         =   "Új Számla létrehozása"
      Height          =   615
      Left            =   10440
      TabIndex        =   45
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton mentes 
      Caption         =   "Számla lezárása"
      Height          =   375
      Left            =   4920
      TabIndex        =   44
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton nyomtat 
      Caption         =   "Nyomtatás"
      Height          =   375
      Left            =   4920
      TabIndex        =   43
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame ossz 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   7800
      Width           =   12975
      Begin VB.Label brutto_l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fizetendõ mindösszesen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8880
         TabIndex        =   28
         Top             =   120
         Width           =   2640
      End
      Begin VB.Label afa_l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Áfa érték mindösszesen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5040
         TabIndex        =   27
         Top             =   120
         Width           =   2550
      End
      Begin VB.Label netto_l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nettó érték mindösszesen: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   2820
      End
   End
   Begin VB.CommandButton mind 
      Caption         =   "Összes tétel törlése"
      Height          =   375
      Left            =   8520
      TabIndex        =   24
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton frissit_btn 
      Caption         =   "Frissit"
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton torol_btn 
      Caption         =   "Törlés"
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton hozzaad 
      Caption         =   "Tétel felvétele"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Frame keret 
      Caption         =   "Számla fejléce"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.TextBox ktj 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox kuj 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox ado 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2400
         Width           =   2055
      End
      Begin VB.ComboBox vevo_lista 
         Height          =   315
         Index           =   1
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1920
         Width           =   6135
      End
      Begin VB.CommandButton uj_vevo 
         Caption         =   "Új partner"
         Height          =   375
         Index           =   1
         Left            =   7080
         TabIndex        =   32
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox cim 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2400
         Width           =   6135
      End
      Begin VB.CommandButton vevo_mod 
         Caption         =   "Adatok Módosítása"
         Height          =   375
         Index           =   1
         Left            =   8280
         TabIndex        =   30
         Top             =   1920
         Width           =   1695
      End
      Begin VB.ComboBox peldany 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3240
         Width           =   1215
      End
      Begin VB.ComboBox fizmod 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton vevo_mod 
         Caption         =   "Adatok Módosítása"
         Height          =   375
         Index           =   0
         Left            =   8280
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox cim 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   6135
      End
      Begin VB.CommandButton uj_vevo 
         Caption         =   "Új partner"
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox vevo_lista 
         Height          =   315
         Index           =   0
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   6135
      End
      Begin VB.TextBox ado 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker teljesites 
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   45088769
         CurrentDate     =   38555
      End
      Begin MSComCtl2.DTPicker kelt 
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   45088769
         CurrentDate     =   38555
      End
      Begin MSComCtl2.DTPicker hatarido 
         Height          =   375
         Left            =   7080
         TabIndex        =   15
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   45088769
         CurrentDate     =   38555
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KTJ:"
         Height          =   195
         Index           =   13
         Left            =   5280
         TabIndex        =   41
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KUJ:"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   345
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   9960
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Szállító:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   37
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cím:"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   36
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adószám:"
         Height          =   195
         Index           =   9
         Left            =   7080
         TabIndex        =   35
         Top             =   2400
         Width           =   690
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Példányszám:"
         Height          =   195
         Index           =   8
         Left            =   3000
         TabIndex        =   23
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fizetés módja:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   3240
         Width           =   1005
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fizetés határideje:"
         Height          =   195
         Index           =   6
         Left            =   5640
         TabIndex        =   20
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teljesítés határideje:"
         Height          =   195
         Index           =   5
         Left            =   2520
         TabIndex        =   19
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Számla kelte:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   2880
         Width           =   945
      End
      Begin VB.Label szamlaszam_l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Számla még nincs elmentve"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   3105
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Számlaszám:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   915
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adószám:"
         Height          =   195
         Index           =   1
         Left            =   7080
         TabIndex        =   8
         Top             =   840
         Width           =   690
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cím:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   315
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vevõ:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   420
      End
   End
   Begin MSComctlLib.ListView hulladek 
      Height          =   3255
      Left            =   120
      TabIndex        =   29
      Top             =   4440
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5741
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
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Azonosító"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "EWC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Megnevezés"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "További hasznosítás módja"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Súly"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Kg-kénti egységára (Áfa nélkül)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Értéke (Áfa nélkül)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Áfa kulcsa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Az Áfa összege"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Értéke (Áfával együtt)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton sztorno 
      Caption         =   "Sztornózás"
      Height          =   375
      Left            =   8520
      TabIndex        =   42
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "hulladek_szamla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim all As Byte
Dim elem As ListItem
Dim Vevo As Long, Szallito As Long
Dim Kinek As Byte
Dim SID As Long
Dim SzamlaSzam As Long


Dim oAfa As Double 'összáfa
Dim oNetto As Double 'össznetto
Dim oBrutto As Double 'összbrutto

Const Terkoz = 100
'Const SID = 1




Private Sub felhasznalovaltas_Click()
    Beleptet
End Sub

Private Sub torol_btn_Click()
On Error Resume Next
    If all = 0 Then
        TetelTorol (hulladek.SelectedItem.Text)
    End If
    Frissit
End Sub

Private Sub frissit_btn_Click()
    Frissit
End Sub



Private Sub Form_Load()
    Form_Initialize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        hulladek.Move 120, 4440, Me.ScaleWidth - 2 * hulladek.Left, Me.ScaleHeight - hulladek.Top - 1000
        ossz.Move hulladek.Left, Me.ScaleHeight - ossz.Height - Terkoz, Me.ScaleWidth - (2 * ossz.Left)
End Sub
Public Sub Frissit()
        Dim Sor As New ADODB.Recordset
        Dim p As String
        
        Dim cAr As Double
        Dim cAfa As Byte
        Dim cSuly As Double
        
        '                   0           1           2           3           4                   5                   6
        p = "SELECT raktarkeszlet.id, ewc.ewc, ewc.veszelyes, ewc.nev, raktarkeszlet.suly, raktarkeszlet.afa, raktarkeszlet.ar, szamla.id, raktarkeszlet.gyszam " & _
            "FROM (szamla INNER JOIN elkelt AS elkelt_1 ON szamla.id = elkelt_1.szamla) INNER JOIN (ewc INNER JOIN raktarkeszlet ON ewc.id = raktarkeszlet.ewc) ON elkelt_1.tetel = raktarkeszlet.id " & _
            "WHERE (((szamla.id)=" & SID & ")); "

        SQL_p p, Sor
        
        'Találatok ellenõrzése
        If Sor.RecordCount = 0 Or all = 1 Then
            mentes.Enabled = False
            torol_btn.Enabled = False
            mind.Enabled = False
        Else
            mentes.Enabled = True
            torol_btn.Enabled = True
            mind.Enabled = True
        End If
        
        'Felvételek
        hulladek.Visible = False
        hulladek.ListItems.Clear
        
        oAfa = 0
        oNetto = 0
        oBrutto = 0
        
        If Not Sor.EOF Then Sor.MoveFirst
                
        Do While Not Sor.EOF
            cAr = Sor.Fields(6).Value
            cAfa = Sor.Fields(5).Value
            cSuly = Abs(Sor.Fields(4).Value)
            
            Set elem = hulladek.ListItems.Add(, , Sor.Fields(0).Value)
            elem.ListSubItems.Add , , Nstr(Sor.Fields(1).Value & Alakit(Sor.Fields(2).Value, "*", "")) 'EWC
            elem.ListSubItems.Add , , Nstr(Sor.Fields(3).Value) 'Megnevezés
            elem.ListSubItems.Add , , Nstr(Sor.Fields(8).Value)  'További hasznosítás
            elem.ListSubItems.Add , , Nstr(cSuly & " kg") 'SUly
            elem.ListSubItems.Add , , Nstr(cAr) 'Egységár
            elem.ListSubItems.Add , , Nstr(cSuly * cAr) 'Össz nettó ár
            elem.ListSubItems.Add , , Nstr(cAfa) 'Afakulcs
            elem.ListSubItems.Add , , Nstr((cAfa / 100) * cAr * cSuly) 'Áfa összeg
            elem.ListSubItems.Add , , Nstr((1 + (cAfa / 100)) * cAr * cSuly) 'Teljes Ar
            'elem.ListSubItems.Add , , Nstr((1 + (cAfa / 100)) * cAr) 'Teljes Ar
            
            oNetto = oNetto + (cAr * cSuly)
            oAfa = oAfa + ((cAfa / 100) * (cAr * cSuly))
            oBrutto = oBrutto + ((1 + (cAfa / 100)) * (cAr * cSuly))
            Sor.MoveNext
        Loop
        
        netto_l.Caption = "Nettó érték mindösszesen: " & oNetto & " Ft"
        brutto_l.Caption = "Fizetendõ (brutto) mindösszesen: " & oBrutto & " Ft"
        afa_l.Caption = "Áfa érték mindösszesen: " & oAfa & " Ft"
        hulladek.Visible = True
End Sub

Private Sub Form_Initialize()
    Kinek = 0
    all = 0
    SID = 0
    SzamlaSzam = 0
    
    fizmod.Clear
    fizmod.AddItem "Készpénz"
    fizmod.AddItem "Átutalás"
    fizmod.ListIndex = 0
    
    peldany.Clear
    peldany.List(0) = 2
    peldany.ItemData(0) = 2
    peldany.List(1) = 3
    peldany.ItemData(1) = 3
    peldany.List(2) = 4
    peldany.ItemData(2) = 4
    peldany.ListIndex = 1
    
    kelt.Value = Date
    hatarido.Value = kelt.Value
    teljesites.Value = kelt.Value
    
    szamlaszam_l.Caption = "A számla még nincs elmentve!" 'SID & "/" & Year(Date)
    
    
    
    PartnerFrissit
    ElsotJelol Me.vevo_lista(0)
    ElsotJelol Me.vevo_lista(1)
    vevo_lista_Change 0
    vevo_lista_Change 1
    
    'Frissit
End Sub
Public Sub PartnerFrissit()
    Partner_Listaba Me.vevo_lista(0)
    Partner_Listaba Me.vevo_lista(1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If all = 3 Then Exit Sub
    If vbYes = MsgBox("Biztos be akarja zárni a számlázót?", vbYesNo + vbQuestion, "Kilépés megerõsítése") Then
        Cancel = 0
        SzamlaTorol
    Else
        Cancel = 1
    End If
End Sub

Private Sub hozzaad_Click()
    hulladek_eladasra.uj SID
End Sub

Private Sub hulladek_DblClick()
On Error Resume Next
    If all = 0 Then
        hulladek_eladasra.modosit hulladek.SelectedItem.Text, SID
    End If
    Frissit
End Sub

Private Sub mentes_Click()
    If vbYes = MsgBox("Biztos le akarja zárni a számlát? Lezárás után már nem javíthat!", vbQuestion + vbYesNo, "Számla lezárása") Then
        Dim tsid As Long
        Dim Sor As New ADODB.Recordset
        
        If SzamlaSzam = 0 Then
            SzamlaSzam = CLng(Ertek("bonto", "id", "1", "szam_sz")) + 1
            SQL_p "UPDATE bonto SET szam_sz=" & SzamlaSzam, Sor
        End If
        
        SQL_p "UPDATE szamla SET szam=" & SzamlaSzam & ", ido='" & Time & "', kelt='" & kelt.Value & "', teljesites='" & teljesites.Value & "', hatarido='" & hatarido.Value & "', fizmod=" & fizmod.ListIndex & ", peldany=" & peldany.List(peldany.ListIndex) & ", vevo=" & Vevo & ", szallito=" & Szallito & ", uid=" & Fid & " WHERE id=" & SID, Sor
        
        Nyszam_Frissit
        
        'Módosított megnyit
        tsid = SID
        all = 3
        Unload Me
        hulladek_szamla.modosit tsid
        nyomtat_Click
    End If
End Sub

Private Sub mind_Click()
    MindenTorol
    Frissit
End Sub

Private Sub nyomtat_Click()
    Dim Sablon As String
    Dim Fajl As String
    
    Frissit
    
    Sablon = Konyvtar & "Sablonok\hullszamla.htm"
    Fajl = "c:\Windows\Temp\" & TmpGeneral(Sablon)
    
    TobbOldalasSzamla Fajl, Sablon
    nyomtatasikep.szamla Fajl
    Novel "szamla", "id", CStr(SID), "nyomtatva"
End Sub

Private Sub sztorno_Click()
 If vbYes = MsgBox("Biztosan sztornózni szeretné a számlát?", vbYesNo + vbQuestion, "Számla sztornózása") Then
    If all = 0 Then
        SzamlaTorol
    Else
        'Kész számla sztornózása
        Dim i As Integer
        Dim Sor As New ADODB.Recordset
        SQL_p "UPDATE szamla SET tipus=3 WHERE id=" & SID, Sor
        
        For i = 1 To hulladek.ListItems.Count
            ElkeltTorol hulladek.ListItems(i).Text
        Next i
    End If
    
    sztorno.Enabled = False
End If
End Sub



Private Sub uj_szamla_Click()
    Dim h As Byte
    h = Kinek
    Unload Me
    hulladek_szamla.uj h
End Sub

Private Sub uj_vevo_Click(Index As Integer)
    partner_lap.uj (50 + Index)
End Sub

Private Sub vevo_lista_Change(Index As Integer)
    BetoltVevo CByte(Index), vevo_lista(Index).ItemData(vevo_lista(Index).ListIndex)
End Sub

'Vevõ adatainak betöltése a szövegdobozokba
Private Sub BetoltVevo(Kinek As Byte, Id As Long)
On Error Resume Next
    Dim Sor As New ADODB.Recordset
    SQL_p "SELECT * from partnerek where id=" & Id, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        
        If Kinek = 0 Then Vevo = Id Else Szallito = Id
        
        cim(Kinek).Text = Nstr(Sor!irszam & " " & Nstr(Sor!varos) & " " & Sor!cim)
        ado(Kinek).Text = Nstr(Sor!ado)
    Else
        MsgBox "Nincs ilyen rekord!"
    End If
    If Kinek = 0 Then
        vevo_lista(1).ListIndex = vevo_lista(0).ListIndex
        kuj.Text = Nstr(Sor!kuj)
        ktj.Text = Nstr(Sor!ktj)
        BetoltVevo 1, Id
    End If
    Sor.Close
End Sub

Private Sub vevo_lista_Click(Index As Integer)
    vevo_lista_Change (Index)
End Sub

'Visszatérésnél partner beszúrása kívülrõl
Public Sub BeszurPartner(Kinek As Byte, Id As Long)
    PartnerFrissit
    Jelol Me.vevo_lista(Kinek), Id
    If Kinek = 0 Then
        Jelol Me.vevo_lista(1), Szallito
    Else
        Jelol Me.vevo_lista(0), Vevo
    End If
    
    vevo_lista_Change (Kinek)
End Sub
Private Sub vevo_mod_Click(Index As Integer)
     partner_lap.modosit vevo_lista(Index).ItemData(vevo_lista(Index).ListIndex), (50 + Index)
End Sub

Public Sub uj(Optional Hova As Byte)
    Form_Initialize
    Kinek = Hova
    
    Dim Sor As New ADODB.Recordset
   
    SQL_p "INSERT INTO szamla (tipus, kelt, hatarido, teljesites, uid) VALUES (2, '" & Date & "', '" & Date & "', '" & Date & "', " & Fid & ")", Sor
    SQL_p "SELECT id FROM szamla order by id desc", Sor
    SID = CLng(Sor!Id)
    'MsgBox SID
    
    sztorno.Visible = False
    
    Frissit
    Me.Show
End Sub
'Teljes számla törlése
Private Sub SzamlaTorol()
    If SzamlaSzam = 0 And all = 0 Then
        Dim Sor As New ADODB.Recordset
        
        MindenTorol
        SQL_p "DELETE * FROM szamla where id=" & SID, Sor
        'SQL_p "UPDATE bonto SET szam_sz=" & (CLng(Ertek("bonto", "id", "1", "szam_sz")) - 1), Sor
        'MsgBox "SZámla megsemmisítve"
    End If
End Sub
'MInden tétel torol
Private Sub MindenTorol()
On Error Resume Next
    Dim i As Integer
    For i = 1 To hulladek.ListItems.Count
            TetelTorol hulladek.ListItems(i).Text
    Next i
End Sub
'1 töröl
Private Sub TetelTorol(Melyiket As Long)
    Dim Sor As New ADODB.Recordset
    SQL_p "SELECT * FROM raktarkeszlet WHERE id=" & Melyiket, Sor
    Sor.MoveFirst
    
    EladAruHulladek Sor!ewc, HulladekDarab(Sor!ewc, Sor!suly), True
    
    FSQL "DELETE * FROM elkelt where szamla=" & SID & " and tetel=" & Melyiket
    FSQL "DELETE * FROM raktarkeszlet where id=" & Melyiket
    
    
    'ElkeltTorol Melyiket
End Sub
Private Sub ElkeltTorol(Melyiket As Long)
    Dim Sor As New ADODB.Recordset
    SQL_p "UPDATE raktarkeszlet SET sztorno=TRUE where id=" & Melyiket, Sor
End Sub
'Nyilvántartási szám eegj
Private Sub Nyszam_Frissit()
    SzamlaSzam = Ertek("szamla", "id", CStr(SID), "szam")
    szamlaszam_l.Caption = SzamlaSzam & "/" & Year(Date)
End Sub

Public Sub modosit(Melyiket As Long, Optional Hova As Byte)
    Form_Initialize
    Kinek = Hova
    SID = Melyiket
    all = 1
    
    Dim Sor As New ADODB.Recordset
    
    SQL_p "SELECT * FROM szamla WHERE id=" & SID, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        
        BeszurPartner 0, Sor!Vevo
        BeszurPartner 1, Sor!Szallito
        kelt.Value = Nstr(Sor!kelt)
        hatarido.Value = Nstr(Sor!hatarido)
        teljesites.Value = Nstr(Sor!teljesites)
        fizmod.ListIndex = Nstr(Sor!fizmod)
        Jelol Me.peldany, Nstr(Sor!peldany)
        
        Nyszam_Frissit
        Frissit
        
        'ELtûntetések
        vevo_lista(0).Enabled = False
        uj_vevo(0).Enabled = False
        vevo_mod(0).Enabled = False
        
        vevo_lista(1).Enabled = False
        uj_vevo(1).Enabled = False
        vevo_mod(1).Enabled = False
        
        kelt.Enabled = False
        teljesites.Enabled = False
        hatarido.Enabled = False
        
        fizmod.Enabled = False
        peldany.Enabled = False
        
        torol_btn.Visible = False
        mind.Visible = False
        hozzaad.Visible = False
        mentes.Visible = False
        
        
        'Mutatások
        sztorno.Visible = True
        nyomtat.Visible = True
        Sor.Close
        Me.Show
    Else
        MsgBox "HIBA"
        Sor.Close
        Unload Me
    End If
End Sub
'Több oldalas nyomtatás
Public Sub TobbOldalasSzamla(Kimenet As String, Forras As String, Optional DB As Integer)
    Dim bonto As New ADODB.Recordset
    Dim Vevo As New ADODB.Recordset
    Dim Sor As String
    Dim i As Integer, mutato As Integer

    SQL_p "SELECT * FROM bonto WHERE id=1", bonto
    SQL_p "SELECT * FROM partnerek where id=" & vevo_lista(0).ItemData(vevo_lista(0).ListIndex), Vevo
    mutato = 1
    
    Open Kimenet For Output As 1
        Open Forras For Input As 2
            Do While Not EOF(2)
                Line Input #2, Sor
                Select Case Trim(Sor)
                    Case "<#!SZAMLASZAM!#>"
                        Print #1, szamlaszam_l.Caption
                    Case "<#!TARTALOM!#>"
                        i = 1
                        'mutato = 1
                        Do While mutato <= hulladek.ListItems.Count ' And i <= DB
                            Print #1, "<TR>"
                            Print #1, "  <td>" & hulladek.ListItems(mutato).ListSubItems(1).Text & "<br>" & hulladek.ListItems(mutato).ListSubItems(2)
                            Print #1, "  </td>"
                            Print #1, "  <td>" & hulladek.ListItems(mutato).ListSubItems(3).Text & "</td>"
                            Print #1, "  <td>" & hulladek.ListItems(mutato).ListSubItems(4).Text & "</td>"
                            Print #1, "  <td>" & hulladek.ListItems(mutato).ListSubItems(5).Text & "</td>"
                            Print #1, "  <td>" & hulladek.ListItems(mutato).ListSubItems(6).Text & "</td>"
                            Print #1, "  <td>" & hulladek.ListItems(mutato).ListSubItems(7).Text & "</td>"
                            Print #1, "  <td>" & hulladek.ListItems(mutato).ListSubItems(8).Text & "</td>"
                            Print #1, "  <td>" & hulladek.ListItems(mutato).ListSubItems(9).Text & "</td>"
                            Print #1, "</TR>"
                            'Adatok.Fields(0).Value
                            
                            'i = i + 1
                            mutato = mutato + 1
                            'Adatok.MoveNext
                        Loop
                    Case "<#!ELADO_ADAT!#>"
                        Print #1, bonto!nev & "<br />"
                        Print #1, bonto!cg & "<br />"
                        Print #1, bonto!ado & "<br />"
                        Print #1, bonto!irszam & ", " & bonto!varos & " " & bonto!utca & " " & bonto!hazszam & "<br />"
                    
                    Case "<#!VEVO_ADAT!#>"
                        Print #1, Vevo!vnev & " " & Vevo!knev & "<br />"
                        Print #1, Vevo!ado & "<br />"
                        Print #1, Vevo!irszam & ", " & Vevo!varos & " " & Vevo!cim & "<br />"
                    Case "<#!FIZMOD!#>"
                        Print #1, fizmod.List(fizmod.ListIndex)
                    Case "<#!TELJDAT!#>"
                        Print #1, teljesites.Value
                    Case "<#!KELT!#>"
                        Print #1, kelt.Value
                    Case "<#!HATARIDO!#>"
                        Print #1, hatarido.Value
                    Case "<#!PELDANYSZAM!#>"
                        Print #1, peldany.List(peldany.ListIndex)
                    Case "<#!OSSZ!#>"
                        Print #1, oBrutto & "Ft"
                    Case "<#!OSSZ_SZOV!#>"
                        Dim Szam As Object
                        Set Szam = CreateObject("Szamok.Irasa")
                        Print #1, Szam.szamot_szovegge(oBrutto) & "forint. "
                    Case Else
                        Print #1, Sor
                End Select
            Loop
        Close 2
    Close 1
    
    'nyomtatasikep.gombsor.Visible = False
    'nyomtatasikep.Show
    'nyomtatasikep.bongeszo.Navigate2 Kimenet
End Sub
