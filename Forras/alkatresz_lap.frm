VERSION 5.00
Begin VB.Form alkatresz_lap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alkatrész felvétel"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   8025
      TabIndex        =   52
      Top             =   0
      Width           =   8055
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Új alkatrész felvétele"
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
         Left            =   5040
         TabIndex        =   53
         Top             =   240
         Width           =   2835
      End
   End
   Begin VB.TextBox megj 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   21
      Top             =   5280
      Width           =   7815
   End
   Begin VB.ComboBox afa 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4680
      Width           =   855
   End
   Begin VB.ComboBox szin 
      Height          =   315
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CheckBox hianyos 
      Caption         =   "H"
      Height          =   255
      Left            =   3240
      TabIndex        =   47
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox allapot 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox bear 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Text            =   "0"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox ar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Text            =   "0"
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton elado_mod 
      Caption         =   "Módosít"
      Height          =   255
      Left            =   7200
      TabIndex        =   26
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton uj_elado 
      Caption         =   "Új"
      Height          =   255
      Left            =   6600
      TabIndex        =   25
      Top             =   960
      Width           =   495
   End
   Begin VB.ComboBox elado_lista 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   5775
   End
   Begin VB.CommandButton bezar 
      Caption         =   "Mégse"
      Height          =   375
      Left            =   6120
      TabIndex        =   24
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton felvesz_uj 
      Caption         =   "Felvesz és újat kezd"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton felvesz_zar 
      Caption         =   "Felvesz és bezár"
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   5880
      Width           =   2055
   End
   Begin VB.ComboBox hajtoanyag 
      Height          =   315
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox tomeg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Text            =   "0"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton ujmtip 
      Caption         =   "Új"
      Height          =   255
      Left            =   7320
      TabIndex        =   27
      Top             =   1680
      Width           =   615
   End
   Begin VB.ComboBox allam 
      Height          =   315
      ItemData        =   "alkatresz_lap.frx":0000
      Left            =   7080
      List            =   "alkatresz_lap.frx":0002
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox szamla 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   1335
   End
   Begin VB.ComboBox kategoria 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2280
      Width           =   4455
   End
   Begin VB.TextBox nyszam 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame komplettmotor 
      BorderStyle     =   0  'None
      Caption         =   "Motor adatai:"
      Height          =   1215
      Left            =   5040
      TabIndex        =   34
      Top             =   3840
      Width           =   2895
      Begin VB.TextBox henger 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox motor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   18
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox motorkod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   19
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hengerûrtartalom:"
         Height          =   195
         Index           =   10
         Left            =   1440
         TabIndex        =   43
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motorszám:"
         Height          =   195
         Index           =   11
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   810
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motorkód:"
         Height          =   195
         Index           =   9
         Left            =   0
         TabIndex        =   41
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.ComboBox evjarat 
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      Text            =   "evjarat"
      Top             =   2280
      Width           =   975
   End
   Begin VB.ComboBox alkatresz 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3480
      Width           =   7815
   End
   Begin VB.ComboBox alcsop 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2880
      Width           =   4455
   End
   Begin VB.ComboBox focsop 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2880
      Width           =   3255
   End
   Begin VB.ComboBox tipus_lista 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1680
      Width           =   3735
   End
   Begin VB.ComboBox marka_lista 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Megjegyzés:"
      Height          =   195
      Index           =   19
      Left            =   120
      TabIndex        =   51
      Top             =   5040
      Width           =   885
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Áfa:"
      Height          =   195
      Index           =   17
      Left            =   2640
      TabIndex        =   50
      Top             =   4440
      Width           =   285
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Színe:"
      Height          =   195
      Index           =   16
      Left            =   3840
      TabIndex        =   49
      Top             =   3840
      Width           =   450
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Állapota:"
      Height          =   195
      Index           =   15
      Left            =   2640
      TabIndex        =   48
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ár:"
      Height          =   195
      Index           =   14
      Left            =   1560
      TabIndex        =   46
      Top             =   4440
      Width           =   195
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Beszerzési Ár:"
      Height          =   195
      Index           =   13
      Left            =   1560
      TabIndex        =   45
      Top             =   3840
      Width           =   990
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Eladó:"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   44
      Top             =   960
      Width           =   450
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Államjelzés:"
      Height          =   195
      Index           =   12
      Left            =   7080
      TabIndex        =   40
      Top             =   2040
      Width           =   810
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hajtóanyag:"
      Height          =   195
      Index           =   28
      Left            =   5760
      TabIndex        =   39
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tömege:"
      Height          =   195
      Index           =   18
      Left            =   120
      TabIndex        =   38
      Top             =   3840
      Width           =   630
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Számla száma:"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   37
      Top             =   4440
      Width           =   1050
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kategória:"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   36
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nyilvántartási szám:"
      Height          =   195
      Index           =   23
      Left            =   120
      TabIndex        =   35
      Top             =   3840
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Évjárat:"
      Height          =   195
      Index           =   5
      Left            =   4680
      TabIndex        =   33
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alkatrész:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   32
      Top             =   3240
      Width           =   690
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Altipus:"
      Height          =   195
      Index           =   3
      Left            =   3480
      TabIndex        =   31
      Top             =   2640
      Width           =   510
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fõtípus:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   30
      Top             =   2640
      Width           =   570
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Típus:"
      Height          =   195
      Index           =   1
      Left            =   3480
      TabIndex        =   29
      Top             =   1440
      Width           =   450
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gyártmány:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   1440
      Width           =   795
   End
End
Attribute VB_Name = "alkatresz_lap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Elado As Long
Dim all As Byte
Dim Kinek As Byte
Dim SID As Long

Private Sub alcsop_Change()
    Betolt Me.alkatresz, "alkatresznevek", "nev", "id", , "where alcsop=" & alcsop.ItemData(alcsop.ListIndex)
    'If alcsop.ItemData(alcsop.ListIndex) = 35 Then
        'komplettmotor.Visible = CBool(alcsop.ItemData(alcsop.ListIndex) = MotorID)
    'Else
    '    komplettmotor.Enabled = False
    'End If
End Sub

Private Sub alcsop_Click()
    alcsop_Change
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub bezar_Click()
    Unload Me
End Sub

Private Sub elado_lista_Change()
On Error Resume Next
    Elado = elado_lista.ItemData(elado_lista.ListIndex)
End Sub

Private Sub elado_lista_Click()
    elado_lista_Change
End Sub

Private Sub elado_mod_Click()
    partner_lap.modosit elado_lista.ItemData(elado_lista.ListIndex), 60
End Sub

Private Sub felvesz_uj_Click()
    If felvesz Then uj Kinek
End Sub

Private Sub felvesz_zar_Click()
    If felvesz Then Unload Me
End Sub

Private Sub focsop_Change()
    Betolt Me.alcsop, "alcsop", "nev", "nev", , "where focsop=" & focsop.ItemData(focsop.ListIndex)
End Sub

Private Sub focsop_Click()
    focsop_Change
End Sub

Private Sub Form_Initialize()
    felso = bontoware.zold
    all = 0
    Kinek = 0
    
    ar.Text = 0
    motor.Text = ""
    motorkod.Text = ""
    henger.Text = 0
    szamla.Text = ""
    bear.Text = 0
    tomeg.Text = 0
    
    
    EvjaratListaba Me.evjarat
    Szin_Betolt Me.szin
    Betolt Me.kategoria, "kategoria", "nev", "id"
    Betolt Me.marka_lista, "markak", "marka", "marka"
    Betolt Me.focsop, "focsop", "nev", "id"
    Betolt Me.hajtoanyag, "uzemanyag", "nev", "id"
    Betolt Me.allam, "allamjelzes", "nev", "id"
    Betolt Me.afa, "afa", "afa", "id"
    
    allapot.List(0) = "ép"
    allapot.ItemData(0) = 1
    allapot.List(1) = "sérült"
    allapot.ItemData(1) = 2
    allapot.Text = allapot.List(0)
    
    partner.Partner_Listaba Me.elado_lista
    elado_lista.ListIndex = 0
    nyszam.Text = NySzamAjanlo("szam_a")
End Sub

Private Sub marka_lista_Change()
    Betolt Me.tipus_lista, "tipusok", "tipus", "tipus", "", "where  marka=" & marka_lista.ItemData(marka_lista.ListIndex)
End Sub
Public Sub uj(Optional Hova As Byte)
    Form_Initialize
    Kinek = Hova
    
    Me.Show
End Sub

Private Sub marka_lista_Click()
    marka_lista_Change
End Sub
Private Function felvesz() As Boolean
Dim p As String, i As Long
Dim Sor As New ADODB.Recordset
On Error GoTo Hiba:
     If all = 0 Then
        p = "INSERT INTO autok " & _
            "(uid, allapot, allapotlap, bontva, nyszam, datum, ido, elado, marka, tipus, kategoria, evjarat, motor, motorkod, allam, henger, hajtoanyag, ar, tomeg, szamla, megj) " & _
            "VALUES " & _
            "(" & Fid & ", 0, TRUE, TRUE, 'AL" & szamla.Text & "','" & Date & "', '" & Time & "', " & Elado & ", " & marka_lista.ItemData(marka_lista.ListIndex) & " , " & tipus_lista.ItemData(tipus_lista.ListIndex) & ", " & kategoria.ItemData(kategoria.ListIndex) & ", '" & evjarat.Text & "', '" & motor.Text & "', '" & motorkod.Text & "', '" & allam.Text & "', '" & henger.Text & "', '" & hajtoanyag.Text & "', " & Vesszotlenito(ar.Text) & ", " & Vesszotlenito(tomeg.Text) & ", '" & szamla.Text & "', '" & megj.Text & "');"
                    
        'SQL_p p, Sor
        FSQL p
        Debug.Print p
        'MsgBox p
        'Nyilvántartási szám növelése, ha az az ajánlott
        'If (nyszam.Text = NySzamAjanlo("szam_a")) Then Novel "bonto", "id", "1", "szam_a"
        
        'FRissen beszúrt kocsi idje
'ujra:
        'BezarR Sor
        'p = "SELECT id FROM autok WHERE allapot=0 and nyszam='AL" & szamla.Text & "' and marka=" & marka_lista.ItemData(marka_lista.ListIndex) & " and tipus=" & tipus_lista.ItemData(tipus_lista.ListIndex) & " and kategoria=" & kategoria.ItemData(kategoria.ListIndex) & " and evjarat='" & evjarat.Text & "' and motor='" & motor.Text & "' and henger='" & henger.Text & "' and allam='" & allam.Text & "' and ar=" & Vesszotlenito(ar.Text) & " and tomeg=" & Vesszotlenito(tomeg.Text) & " and szamla='" & szamla.Text & "' order by id desc"
        p = "SELECT id FROM autok WHERE uid=" & Fid & " and allapot=0 and nyszam='AL" & szamla.Text & "' and marka=" & marka_lista.ItemData(marka_lista.ListIndex) & " and tipus=" & tipus_lista.ItemData(tipus_lista.ListIndex) & " and kategoria=" & kategoria.ItemData(kategoria.ListIndex) & " and szamla='" & szamla.Text & "' order by id desc"
        Debug.Print p
        SQL_p p, Sor
        
        'If Not Sor.EOF Then Sor.MoveFirst Else GoTo ujra
        
        If Not Sor.EOF Then Sor.MoveFirst 'Else GoTo ujra
        i = CLng(Sor.Fields(0).Value)
        Sor.Close
        
        'Felvétel az alkatrészek közé
        'Régi adatstruktúra
        'SQL_p "INSERT INTO alkatreszek (alkatresz, auto, szine, allapot, suly, ar, afa) VALUES (" & alkatresz.ItemData(alkatresz.ListIndex) & ", " & i & ", " & szin.ListIndex & ", " & AllapotSzov & ", " & tomeg.Text & ", " & bear.Text & " , " & afa.Text & " )", Sor
        FSQL "INSERT INTO raktarkeszlet (tipus, alkatresz, auto, minosites, allapot, hianyos, suly, ar, afa, ewc, megj) VALUES (0," & alkatresz.ItemData(alkatresz.ListIndex) & ", " & i & ", " & szin.ListIndex & ", " & allapot.ItemData(allapot.ListIndex) & ", " & szin.ListIndex & ", " & Vesszotlenito(tomeg.Text) & ", " & Vesszotlenito(bear.Text) & " , " & afa.Text & ", 1, '" & megj.Text & "' )"
        'Felvétel a hulladékok közé
        'SQL_p "INSERT into hulladek (auto, ewc, suly) VALUES ( " & i & ", 1, " & Vesszotlenito(CStr(tomeg.Text)) & ")", sor
        If Kinek > 0 Then Visszajelez Kinek, i
    Else
        
        If Kinek > 0 Then Visszajelez Kinek, CInt(SID)
    End If
    'Unload Me
    felvesz = True
    
Exit Function
Hiba:
    MsgBox "Nem töltött ki minden adatot!", vbInformation, "Feltöltési hiba"
    felvesz = False
End Function

Private Sub uj_elado_Click()
    partner_lap.uj 60
End Sub
Public Sub BeszurPartner(Kit As Long)
    partner.Partner_Listaba Me.elado_lista
    Jelol Me.elado_lista, Kit
    Elado = Kit
End Sub
Private Function AllapotSzov() As Byte
    AllapotSzov = (allapot.ListIndex + 1) * 10 + hianyos.Value
End Function

Private Sub ujmtip_Click()
    markak.Show vbModal
    Betolt Me.marka_lista, "markak", "marka", "marka"
End Sub
