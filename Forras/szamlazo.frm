VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form szamlazo 
   Caption         =   "Sz�ml�z�s Alkatr�szre"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton felhasznalovaltas 
      Caption         =   "Felhaszn�l� v�lt�s"
      Height          =   615
      Left            =   10320
      TabIndex        =   34
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton uj_szamla 
      Caption         =   "�j Sz�mla l�trehoz�sa"
      Height          =   615
      Left            =   10320
      TabIndex        =   33
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton mentes 
      Caption         =   "Sz�mla lez�r�sa"
      Height          =   375
      Left            =   4560
      TabIndex        =   30
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Timer ell 
      Interval        =   100
      Left            =   2040
      Top             =   2760
   End
   Begin VB.Frame ossz 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   7800
      Width           =   12975
      Begin VB.Label brutto_l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fizetend� mind�sszesen:"
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
         TabIndex        =   29
         Top             =   120
         Width           =   2640
      End
      Begin VB.Label afa_l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�fa �rt�k mind�sszesen:"
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
         TabIndex        =   28
         Top             =   120
         Width           =   2550
      End
      Begin VB.Label netto_l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nett� �rt�k mind�sszesen: "
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
         TabIndex        =   27
         Top             =   120
         Width           =   2820
      End
   End
   Begin VB.CommandButton mind 
      Caption         =   "�sszes t�tel t�rl�se"
      Height          =   375
      Left            =   8280
      TabIndex        =   25
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton torol_btn 
      Caption         =   "T�rl�s"
      Height          =   375
      Left            =   6360
      TabIndex        =   11
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton hozzaad 
      Caption         =   "T�tel felv�tele"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Frame keret 
      Caption         =   "Sz�mla fejl�ce"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9975
      Begin VB.ComboBox peldany 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox fizmod 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton vevo_mod 
         Caption         =   "Adatok M�dos�t�sa"
         Height          =   375
         Left            =   8040
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox cim 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   6135
      End
      Begin VB.CommandButton uj_vevo 
         Caption         =   "�j partner"
         Height          =   375
         Left            =   6840
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox vevo_lista 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox ado 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker teljesites 
         Height          =   375
         Left            =   4080
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   44892161
         CurrentDate     =   38555
      End
      Begin MSComCtl2.DTPicker kelt 
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   44892161
         CurrentDate     =   38555
      End
      Begin MSComCtl2.DTPicker hatarido 
         Height          =   375
         Left            =   6840
         TabIndex        =   16
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   44892161
         CurrentDate     =   38555
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P�ld�ny:"
         Height          =   195
         Index           =   8
         Left            =   3240
         TabIndex        =   24
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fizet�s m�dja:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fizet�s hat�rideje:"
         Height          =   195
         Index           =   6
         Left            =   5520
         TabIndex        =   21
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teljes�t�s hat�rideje:"
         Height          =   195
         Index           =   5
         Left            =   2520
         TabIndex        =   20
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sz�mla kelte:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label szamlaszam_l 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Sz�mla m�g nincs lez�rva"
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
         TabIndex        =   18
         Top             =   240
         Width           =   2925
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sz�mlasz�m:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   915
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ad�sz�m:"
         Height          =   195
         Index           =   1
         Left            =   6840
         TabIndex        =   9
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C�m:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vev�:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   420
      End
   End
   Begin MSComctlLib.ListView cikkek 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7858
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Azonos�t�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Term�k megnevez�se"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cikksz�m"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "VTSZ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Min�s�t�s"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "S�ly"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Mennyis�ge"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Egys�g�ra (�fa n�lk�l)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "�rt�ke (�fa n�lk�l)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "�fa kulcsa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Az �fa �sszege"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "�rt�ke (�f�val egy�tt)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton sztorno 
      Caption         =   "Sztorn�z�s"
      Height          =   375
      Left            =   8280
      TabIndex        =   32
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton nyomtat 
      Caption         =   "Nyomtat�s"
      Height          =   375
      Left            =   2760
      TabIndex        =   31
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton friss_btn 
      Caption         =   "Frissit"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
End
Attribute VB_Name = "szamlazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim all As Byte '0�j, 1m�d, 3ment, 4jog
Dim elem As ListItem
Dim Vevo As Long
Dim Kinek As Byte
Dim SID As Long
Dim SzamlaSzam As Long


Dim oAfa As Double '�ssz�fa
Dim oNetto As Double '�ssznetto
Dim oBrutto As Double '�sszbrutto

Const Terkoz = 100
'Const SID = 1


Private Sub cikkek_DblClick()
On Error Resume Next
    If all = 0 Then
        alkatresz_eladasra.Bearaz cikkek.SelectedItem.Text, , 2
    End If
    Frissit
End Sub

Private Sub torol_btn_Click()
    TetelTorol (cikkek.SelectedItem.Text)
    Frissit
End Sub

Private Sub friss_btn_Click()
    Frissit
End Sub

Private Sub ell_Timer()
    ElnevezAblak Me
    If Not Jogos(Me.Tag, True) Then Unload Me
End Sub

Private Sub felhasznalovaltas_Click()
    Beleptet
End Sub

Private Sub Form_Load()
    Form_Initialize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        cikkek.Move 120, 3240, Me.ScaleWidth - 2 * cikkek.Left, Me.ScaleHeight - cikkek.Top - 1000
        ossz.Move cikkek.Left, Me.ScaleHeight - ossz.Height - Terkoz, Me.ScaleWidth - (2 * ossz.Left)
End Sub
Public Sub Frissit()
        Dim Sor As New ADODB.Recordset
        Dim p As String
        
        SQL_p "SELECT szamla.id, raktarkeszlet.id FROM raktarkeszlet INNER JOIN (szamla INNER JOIN elkelt ON szamla.id = elkelt.szamla) ON raktarkeszlet.id = elkelt.tetel WHERE elkelt.szamla=" & SID, Sor
        
        cikkek.Visible = False
        cikkek.ListItems.Clear
        
        oAfa = 0
        oNetto = 0
        oBrutto = 0
        
        If Sor.RecordCount = 0 Or all = 1 Then
            nyomtat.Enabled = False
            mentes.Enabled = False
            torol_btn.Enabled = False
            mind.Enabled = False
        Else
            nyomtat.Enabled = True
            mentes.Enabled = True
            torol_btn.Enabled = True
            mind.Enabled = True
        End If
        
        If Not Sor.EOF Then
            Sor.MoveFirst
            Do While Not Sor.EOF
                FelveszElem CLng(Sor.Fields(1).Value)
                Sor.MoveNext
            Loop
        End If
        
        netto_l.Caption = "Nett� �rt�k mind�sszesen: " & oNetto & " Ft"
        brutto_l.Caption = "Fizetend� (brutto) mind�sszesen: " & oBrutto & " Ft"
        afa_l.Caption = "�fa �rt�k mind�sszesen: " & oAfa & " Ft"
        cikkek.Visible = True
        
End Sub
Public Sub Beszur(Melyiket As Long)
        Dim Sor As New ADODB.Recordset
        
        SQL_p "UPDATE raktarkeszlet SET elkelt=TRUE, irany=1 WHERE id=" & Melyiket, Sor
        SQL_p "INSERT INTO elkelt (szamla, tetel) VALUES (" & SID & ", " & Melyiket & ")", Sor
        'auto = Ertek("raktarkeszlet", "id", CStr(Melyiket), "auto")
        
        AlkMentKasznitomege Melyiket
        
        Frissit
        'alkatresz_lista.Frissit
        
End Sub

Private Sub FelveszElem(Melyiket As Long)
        Dim Sor As New ADODB.Recordset
        Dim p As String
        
        '                0                1           2              3              4            5               6              7                    8                 9              10            11
        p = "SELECT raktarkeszlet.id, autok.id, markak.marka, tipusok.tipus, focsop.nev, alcsop.nev, alkatresznevek.nev, raktarkeszlet.suly, raktarkeszlet.minosites, raktarkeszlet.ar, raktarkeszlet.afa, vtsz.vtsz, focsop.cikkszam, alcsop.cikkszam, alkatresznevek.cikkszam, raktarkeszlet.alkatresz, autok.motor, autok.valto, focsop.id"
        p = p & " FROM szamla INNER JOIN ((vtsz INNER JOIN (focsop INNER JOIN (alcsop INNER JOIN ((markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN (autok INNER JOIN (alkatresznevek INNER JOIN raktarkeszlet ON alkatresznevek.id = raktarkeszlet.alkatresz) ON autok.id = raktarkeszlet.auto) ON tipusok.id = autok.tipus) ON alcsop.id = alkatresznevek.alcsop) ON focsop.id = alcsop.focsop) ON vtsz.id = alkatresznevek.vtsz) INNER JOIN elkelt ON raktarkeszlet.id = elkelt.tetel) ON szamla.id = elkelt.szamla "
        p = p & " WHERE ((raktarkeszlet.id)=" & Melyiket & ")"
        
        'p = "SELECT alkatreszek.id, autok.nyszam, markak.marka, tipusok.tipus, focsop.nev, alcsop.nev, alkatresznevek.nev, alkatreszek.suly, alkatreszek.szine, alkatreszek.ar, alkatreszek.afa, vtsz.vtsz "
        'p = p & " FROM szamla INNER JOIN ((vtsz INNER JOIN (focsop INNER JOIN (alcsop INNER JOIN ((markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN (autok INNER JOIN (alkatresznevek INNER JOIN alkatreszek ON alkatresznevek.id = alkatreszek.alkatresz) ON autok.id = alkatreszek.auto) ON tipusok.id = autok.tipus) ON alcsop.id = alkatresznevek.alcsop) ON focsop.id = alcsop.focsop) ON vtsz.id = alkatresznevek.vtsz) INNER JOIN elkelt ON alkatreszek.id = elkelt.alkatresz) ON szamla.id = elkelt.szamla "
        'p = p & " WHERE ((alkatreszek.id)=" & Melyiket & ")"
        
        SQL_p p, Sor
        If Not Sor.EOF Then
            Dim cAr As Double
            Dim cAfa As Byte
            Dim cSuly As Double
            
            cAr = Sor.Fields(9).Value
            cAfa = Sor.Fields(10).Value
            cSuly = Sor.Fields(7).Value
            
            Set elem = cikkek.ListItems.Add(, , Melyiket)
            p = Nstr(Sor.Fields(2).Value & " " & Sor.Fields(3).Value & " " & Sor.Fields(5).Value & " " & Sor.Fields(6).Value)
            If Sor.Fields(15).Value = MotorID Then
                p = p & " - motorsz�m: " & Nstr(Sor.Fields(16).Value)
            Else
                If Sor.Fields(14).Value = 1 And Sor.Fields(18).Value = ValtoFcs Then
                    'V�lt�k�d �s sz�m
                    p = p & " - v�lt�k�d:" & Ertek("autok", "id", CStr(Sor.Fields(1).Value), "valto")
                End If
            End If
            elem.ListSubItems.Add , , p
            
            elem.ListSubItems.Add , , Nstr(NKieg(Sor.Fields(12).Value) & NKieg(Sor.Fields(13).Value) & NKieg(Sor.Fields(14).Value))
            elem.ListSubItems.Add , , Nstr(Sor.Fields(11).Value) 'VTSZ
            elem.ListSubItems.Add , , MinositesTipus(Sor.Fields(8).Value) 'Szin
            elem.ListSubItems.Add , , Nstr(cSuly & " kg") 'SUly
            elem.ListSubItems.Add , , "1 db"
            elem.ListSubItems.Add , , Nstr(cAr) 'AR
            elem.ListSubItems.Add , , Nstr(cAr) 'Ar
            elem.ListSubItems.Add , , Nstr(cAfa) 'Afa
            elem.ListSubItems.Add , , Nstr((cAfa / 100) * cAr) 'Ao
            elem.ListSubItems.Add , , Nstr((1 + (cAfa / 100)) * cAr) 'Teljes Ar
            'elem.ListSubItems.Add , , Nstr((1 + (cAfa / 100)) * cAr) 'Teljes Ar
            
            oNetto = oNetto + cAr
            oAfa = oAfa + ((cAfa / 100) * cAr)
            oBrutto = oBrutto + ((1 + (cAfa / 100)) * cAr)
        Else
            MsgBox "hiba"
        End If
        Sor.Close
        
End Sub
Private Sub Form_Initialize()
    Kinek = 0
    all = 0
    SID = 0
    SzamlaSzam = 0
    
    fizmod.Clear
    fizmod.AddItem "K�szp�nz"
    fizmod.AddItem "�tutal�s"
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
    
    Lokkol Me.cim, True, False
    Lokkol Me.ado, True, False
    
    szamlaszam_l.Caption = "A sz�mla m�g nincs elmentve!" 'SID & "/" & Year(Date)
    
    PartnerFrissit
    ElsotJelol Me.vevo_lista
    vevo_lista_Change
    
End Sub
Public Sub PartnerFrissit()
    Partner_Listaba Me.vevo_lista
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If all = 3 Then Exit Sub
    If vbYes = MsgBox("Biztos be akarja z�rni a sz�ml�z�t?", vbYesNo + vbQuestion, "Kil�p�s meger�s�t�se") Then
        Cancel = 0
        SzamlaTorol
    Else
        Cancel = 1
    End If
    alkatresz_lista.Hide
End Sub

Private Sub hozzaad_Click()
    'Beszur InputBox("sza")
    alkatresz_lista.meghiv 1
End Sub

Private Sub mentes_Click()
    If vbYes = MsgBox("Biztos le akarja z�rni a sz�ml�t? Lez�r�s ut�n m�r nem jav�that!", vbQuestion + vbYesNo, "Sz�mla lez�r�sa") Then
        Dim Sor As New ADODB.Recordset
        Dim tsid As Long
    
        If SzamlaSzam = 0 Then
            SzamlaSzam = CLng(Ertek("bonto", "id", "1", "szam_sz")) + 1
            SQL_p "UPDATE bonto SET szam_sz=" & SzamlaSzam, Sor
        End If
        
        SQL_p "UPDATE szamla SET szam=" & SzamlaSzam & ", ido='" & Time & "', kelt='" & kelt.Value & "', teljesites='" & teljesites.Value & "', hatarido='" & hatarido.Value & "', fizmod=" & fizmod.ListIndex & ", peldany=" & peldany.List(peldany.ListIndex) & ", vevo=" & Vevo & ", szallito=" & Vevo & ", uid=" & Fid & " WHERE id=" & SID, Sor
        
        Nyszam_Frissit
        
        'M�dos�tott megnyit
        tsid = SID
        all = 3
        Unload Me
        szamlazo.modosit tsid
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
    
    Sablon = Konyvtar & "Sablonok\alkszamla.htm"
    Fajl = "c:\Windows\Temp\" & TmpGeneral(Sablon)
    
    TobbOldalasSzamla Fajl, Sablon
    nyomtatasikep.szamla Fajl
    Novel "szamla", "id", CStr(SID), "nyomtatva"
    
    
End Sub

Private Sub sztorno_Click()
 If vbYes = MsgBox("Biztosan sztorn�zni szeretn� a sz�ml�t?", vbYesNo + vbQuestion, "Sz�mla sztorn�z�sa") Then
    If all = 0 And SID = 0 Then
        SzamlaTorol
    Else
        If Not Jogos(2, True) Then Exit Sub
        'K�sz sz�mla sztorn�z�sa
        Dim i As Integer
        Dim Sor As New ADODB.Recordset
        SQL_p "UPDATE szamla SET tipus=1 WHERE id=" & SID, Sor
        
        For i = 1 To cikkek.ListItems.Count
            'ElkeltTorol cikkek.ListItems(i).Text
            SztornoTetel cikkek.ListItems(i).Text
        Next i
    End If
    sztorno.Enabled = False
End If
End Sub

Private Sub uj_szamla_Click()
    Dim h As Byte
    h = Kinek
    Unload Me
    szamlazo.uj h
End Sub

Private Sub uj_vevo_Click()
    partner_lap.uj 40
End Sub

Private Sub vevo_lista_Change()
    BetoltVevo vevo_lista.ItemData(vevo_lista.ListIndex)
End Sub

'Vev� adatainak bet�lt�se a sz�vegdobozokba
Private Sub BetoltVevo(Id As Long)
    Dim Sor As New ADODB.Recordset
    SQL_p "SELECT * from partnerek where id=" & Id, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        Vevo = Id
        cim.Text = Nstr(Sor!irszam & " " & Nstr(Sor!varos) & " " & Sor!cim)
        'szemelyi.Text = Nstr(sor!szemelyi)
        ado.Text = Nstr(Sor!ado)
        'allampolg.Text = Nstr(sor!allampolg)
        Sor.Close
    Else
        MsgBox "Nincs ilyen rekord!"
    End If
End Sub

Private Sub vevo_lista_Click()
    vevo_lista_Change
End Sub

Private Sub vevo_lista_Validate(Cancel As Boolean)
    vevo_lista_Change
End Sub
'Visszat�r�sn�l partner besz�r�sa k�v�lr�l
Public Sub BeszurPartner(Id As Long)
    PartnerFrissit
    Jelol Me.vevo_lista, Id
    vevo_lista_Change
End Sub

Private Sub vevo_mod_Click()
     partner_lap.modosit vevo_lista.ItemData(vevo_lista.ListIndex), 40
End Sub

Public Sub uj(Optional Hova As Byte)
    Form_Initialize
    Kinek = Hova
    
    Dim Sor As New ADODB.Recordset
   
    SQL_p "INSERT INTO szamla (tipus, kelt, hatarido, teljesites, uid) VALUES (0, '" & Date & "', '" & Date & "', '" & Date & "', " & Fid & ")", Sor
    SQL_p "SELECT id FROM szamla order by id desc", Sor
    SID = CLng(Sor!Id)
    'MsgBox SID
    
    Frissit
    Me.Show
End Sub
'Teljes sz�mla t�rl�se
Private Sub SzamlaTorol()
        If SzamlaSzam = 0 And all = 0 Then
            Dim Sor As New ADODB.Recordset
        
            MindenTorol
            SQL_p "DELETE * FROM szamla where id=" & SID, Sor
            'SQL_p "UPDATE bonto SET szam_sz=" & (CLng(Ertek("bonto", "id", "1", "szam_sz")) - 1), Sor
            'MsgBox "SZ�mla megsemmis�tve"
        End If
End Sub
'MInden t�tel torol
Private Sub MindenTorol()
On Error Resume Next
    Dim i As Integer
    For i = 1 To cikkek.ListItems.Count
            TetelTorol cikkek.ListItems(i).Text
    Next i
End Sub
'1 t�r�l
Private Sub TetelTorol(Melyiket As Long)
    Dim Sor As New ADODB.Recordset
    SQL_p "DELETE * FROM elkelt where szamla=" & SID & " and tetel=" & Melyiket, Sor
    AlkMentKasznitomege Melyiket
    ElkeltTorol Melyiket
End Sub
Private Sub ElkeltTorol(Melyiket As Long)
    Dim Sor As New ADODB.Recordset
    SQL_p "UPDATE raktarkeszlet SET elkelt=FALSE, irany=1 where id=" & Melyiket, Sor
End Sub
'Nyilv�ntart�si sz�m eegj
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
        
        BeszurPartner Sor!Vevo
        kelt.Value = Nstr(Sor!kelt)
        hatarido.Value = Nstr(Sor!hatarido)
        teljesites.Value = Nstr(Sor!teljesites)
        fizmod.ListIndex = Nstr(Sor!fizmod)
        Jelol Me.peldany, Nstr(Sor!peldany)
        
        Nyszam_Frissit
        Frissit
        
        Sor.Close
        'ELt�ntet�sek
        vevo_lista.Enabled = False
        uj_vevo.Enabled = False
        vevo_mod.Enabled = False
        
        kelt.Enabled = False
        teljesites.Enabled = False
        hatarido.Enabled = False
        
        fizmod.Enabled = False
        peldany.Enabled = False
        
        torol_btn.Visible = False
        mind.Visible = False
        hozzaad.Visible = False
        mentes.Visible = False
        
        
        'Mutat�sok
        sztorno.Visible = True
        
        Me.Show
    Else
        MsgBox "HIBA"
        Sor.Close
        Unload Me
    End If
    
    
End Sub
'T�telek sztorn�z�sa
Public Sub SztornoTetel(Melyiket As Long)
    Dim Sor As New ADODB.Recordset
    
    Dim p As String, Seged As String
    Dim i As Integer
    SQL_p "SELECT * FROM raktarkeszlet WHERE ID=" & Melyiket, Sor
    Sor.MoveFirst
    p = "INSERT INTO raktarkeszlet ("
    For i = 1 To Sor.Fields.Count - 2
        If i > 1 Then p = p & ", "
        p = p & Sor.Fields(i).Name
    Next i
    p = p & ") VALUES ("
    For i = 1 To Sor.Fields.Count - 2
        If i > 1 Then p = p & ", "
        'MsgBox sor.Fields(i).Properties
        'Ha elkelt, invert�l
        If Sor.Fields(i).Name = "elkelt" Then
            p = p & KonvertalLogikai(Not Sor.Fields(i).Value)
        Else
            If KonvertalLogikai(Sor.Fields(i).Value) = "FALSE" Or KonvertalLogikai(Sor.Fields(i).Value) = "TRUE" Then
                p = p & KonvertalLogikai(Sor.Fields(i).Value)
            Else
                p = p & "'" & KonvertalLogikai(Sor.Fields(i).Value) & "'"
            End If
        End If
    Next i
    Sor.Close
    p = p & ")"
    
    Debug.Print p
    
    FSQL p
    'R�gi t�tel sztorn�ra �ll�t�sa
    FSQL "UPDATE raktarkeszlet SET sztorno=TRUE and irany=0 WHERE id=" & Melyiket
    AlkMentKasznitomege Melyiket
End Sub

'T�bb oldalas nyomtat�s
Public Sub TobbOldalasSzamla(Kimenet As String, Forras As String, Optional DB As Integer)
    Dim bonto As New ADODB.Recordset
    Dim Vevo As New ADODB.Recordset
    Dim Sor As String
    Dim i As Integer, mutato As Integer

    SQL_p "SELECT * FROM bonto WHERE id=1", bonto
    SQL_p "SELECT * FROM partnerek where id=" & vevo_lista.ItemData(vevo_lista.ListIndex), Vevo
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
                        Do While mutato <= cikkek.ListItems.Count ' And i <= DB
                            Print #1, "<TR>"
                            Print #1, "  <td>" & cikkek.ListItems(mutato).ListSubItems(1).Text & "<br>cikkszam: " & cikkek.ListItems(mutato).ListSubItems(2)
                            Print #1, "  </td>"
                            Print #1, "  <td>" & cikkek.ListItems(mutato).ListSubItems(3).Text & "</td>"
                            Print #1, "  <td>" & cikkek.ListItems(mutato).ListSubItems(4).Text & "</td>"
                            Print #1, "  <td>" & cikkek.ListItems(mutato).ListSubItems(9).Text & "</td>"
                            Print #1, "  <td>" & cikkek.ListItems(mutato).ListSubItems(6).Text & "</td>"
                            Print #1, "  <td>" & cikkek.ListItems(mutato).ListSubItems(7).Text & "</td>"
                            Print #1, "  <td>" & cikkek.ListItems(mutato).ListSubItems(8).Text & "</td>"
                            Print #1, "  <td>" & cikkek.ListItems(mutato).ListSubItems(10).Text & "</td>"
                            Print #1, "  <td>" & cikkek.ListItems(mutato).ListSubItems(11).Text & "</td>"
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
Public Sub UjGyorsTetel(Mit As Long, Optional Hova As Byte)
    uj Hova
    'Beszur Mit
    Me.Show
    hozzaad_Click
End Sub
