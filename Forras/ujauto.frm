VERSION 5.00
Begin VB.Form auto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gépjármû felvétele bontásra"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton megse 
      Caption         =   "Mégse"
      Height          =   375
      Left            =   5640
      TabIndex        =   30
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton felvesz 
      Caption         =   "Felvétel"
      Height          =   375
      Left            =   360
      TabIndex        =   29
      Top             =   8880
      Width           =   1335
   End
   Begin VB.Frame keret 
      Caption         =   "Okmányok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   2
      Left            =   120
      TabIndex        =   56
      Top             =   6360
      Width           =   6975
      Begin VB.TextBox bon_azon 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   69
         Top             =   840
         Width           =   2175
      End
      Begin VB.OptionButton x 
         Caption         =   "Meghatalmazott hozta"
         Height          =   195
         Index           =   2
         Left            =   4800
         TabIndex        =   21
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton x 
         Caption         =   "Adásvételi szerzõdéssel hozták"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   20
         Top             =   600
         Width           =   2775
      End
      Begin VB.OptionButton x 
         Caption         =   "Tulajdonos hozta"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox kivonas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox datum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   27
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox ido 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   28
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox nyszam 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   26
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox bon_szam 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   25
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox ar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   24
         Text            =   "0"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox kitiltas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   60
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox bon_forg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox forgalmi 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox torzskonyv 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Szig:"
         Height          =   195
         Index           =   20
         Left            =   4200
         TabIndex        =   70
         Top             =   840
         Width           =   345
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Idõ:"
         Height          =   195
         Index           =   15
         Left            =   5160
         TabIndex        =   66
         Top             =   2040
         Width           =   270
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Végleges Kivonás dátuma:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   65
         Top             =   1320
         Width           =   1890
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dátum:"
         Height          =   195
         Index           =   26
         Left            =   1560
         TabIndex        =   64
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nyilvántartási szám:"
         Height          =   195
         Index           =   23
         Left            =   4080
         TabIndex        =   63
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bontási átvételi igazolás:"
         Height          =   195
         Index           =   24
         Left            =   240
         TabIndex        =   62
         Top             =   1680
         Width           =   1740
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ár:"
         Height          =   195
         Index           =   17
         Left            =   5160
         TabIndex        =   61
         Top             =   1320
         Width           =   195
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forgalmi jogosultja:"
         Height          =   195
         Index           =   25
         Left            =   240
         TabIndex        =   59
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forgalmi engedély:"
         Height          =   195
         Index           =   14
         Left            =   3240
         TabIndex        =   58
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Törzskönyvszám:"
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   57
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame keret 
      Caption         =   "Jármû adatai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Index           =   1
      Left            =   120
      TabIndex        =   42
      Top             =   2520
      Width           =   6975
      Begin VB.ComboBox eredet 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2400
         Width           =   2895
      End
      Begin VB.ComboBox evjarat 
         Height          =   315
         Left            =   5640
         TabIndex        =   4
         Text            =   "evjarat"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox megj 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   2880
         Width           =   6615
      End
      Begin VB.ComboBox allam 
         Height          =   315
         ItemData        =   "ujauto.frx":0000
         Left            =   5640
         List            =   "ujauto.frx":0002
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton ujmtip 
         Caption         =   "Új márka vagy típus"
         Height          =   495
         Left            =   3720
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox tomeg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   13
         Text            =   "0"
         Top             =   2040
         Width           =   735
      End
      Begin VB.ComboBox hajtoanyag 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox henger 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         TabIndex        =   14
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox motorkod 
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
         Left            =   5640
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
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
         Left            =   960
         TabIndex        =   10
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox alvaz 
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
         Left            =   960
         TabIndex        =   8
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox rendszam 
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
         Left            =   5640
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox kategoria 
         Height          =   315
         ItemData        =   "ujauto.frx":0004
         Left            =   960
         List            =   "ujauto.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   3615
      End
      Begin VB.ComboBox szine 
         Height          =   315
         ItemData        =   "ujauto.frx":0008
         Left            =   960
         List            =   "ujauto.frx":0021
         TabIndex        =   12
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox tipus_lista 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2655
      End
      Begin VB.ComboBox marka_lista 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eredet:"
         Height          =   195
         Index           =   16
         Left            =   2760
         TabIndex        =   68
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Megjegyzés:"
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   67
         Top             =   2640
         Width           =   885
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saját tömege:"
         Height          =   195
         Index           =   18
         Left            =   2760
         TabIndex        =   55
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Évjárat:"
         Height          =   195
         Index           =   5
         Left            =   4920
         TabIndex        =   54
         Top             =   240
         Width           =   540
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hajtóanyag:"
         Height          =   195
         Index           =   28
         Left            =   4800
         TabIndex        =   53
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hengerûrtartalom:"
         Height          =   195
         Index           =   10
         Left            =   4800
         TabIndex        =   52
         Top             =   2040
         Width           =   1260
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motorkód:"
         Height          =   195
         Index           =   8
         Left            =   4800
         TabIndex        =   51
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motorszám:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   50
         Top             =   1680
         Width           =   810
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alvázszám:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   49
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Államjelzés:"
         Height          =   195
         Index           =   9
         Left            =   4800
         TabIndex        =   48
         Top             =   960
         Width           =   810
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rendszám:"
         Height          =   195
         Index           =   12
         Left            =   4800
         TabIndex        =   47
         Top             =   600
         Width           =   795
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kategória:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   720
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Színe:"
         Height          =   195
         Index           =   22
         Left            =   360
         TabIndex        =   45
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Típus:"
         Height          =   195
         Index           =   21
         Left            =   480
         TabIndex        =   44
         Top             =   600
         Width           =   450
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gyártmány:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame keret 
      Caption         =   "Eladó"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   720
      Width           =   6975
      Begin VB.OptionButton azonosito 
         Caption         =   "Adószám"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   31
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox ado 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton azonosito 
         Caption         =   "Személyi igazolvány:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox elado_lista 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   5175
      End
      Begin VB.CommandButton uj_elado 
         Caption         =   "Új"
         Height          =   255
         Left            =   5400
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox szemelyi 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox allampolg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox cim 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   600
         Width           =   6135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Módosít"
         Height          =   255
         Left            =   5880
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Állampolgárság:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   1320
         Width           =   1110
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cím:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   315
      End
   End
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7185
      TabIndex        =   71
      Top             =   0
      Width           =   7215
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gépjármû felvétele bontásra"
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
         TabIndex        =   72
         Top             =   240
         Width           =   3915
      End
   End
End
Attribute VB_Name = "auto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim all As Byte '0új, 1mód, 2besz, 3bonrontott
Dim SID As Long
Dim Elado As Long
Dim marka As Long
Dim tipus As Long
Dim Jarmu As Long
Dim Kinek As Byte

Public Sub Frissit()
    Elado_Frissit
    Marka_Frissit
End Sub
Public Sub Elado_Frissit()
    partner.Partner_Listaba Me.elado_lista
End Sub
Public Sub Marka_Frissit()
    Betolt Me.marka_lista, "markak", "marka", "marka", , , marka
End Sub

Public Sub BetoltElado(Id As Long)
    Dim Sor As New ADODB.Recordset
    SQL_p "SELECT * from partnerek where id=" & Id, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        Elado = Id
        cim.Text = Nstr(Sor!irszam & " " & Nstr(Sor!Varos) & " " & Sor!cim)
        szemelyi.Text = Nstr(Sor!szemelyi)
        ado.Text = Nstr(Sor!ado)
        allampolg.Text = Nstr(Sor!allampolg)
        bon_forg.Text = Nstr(Sor!vnev & " " & Sor!knev)
        bon_azon.Text = szemelyi.Text
    Else
        MsgBox "Nincs ilyen rekord!"
    End If
    Sor.Close
End Sub


Private Sub alvaz_LostFocus()
    alvaz.Text = UCase(alvaz.Text)
End Sub

Private Sub azonosito_Click(Index As Integer)
    If Index = 0 Then
        Lokkol Me.szemelyi, True
        Lokkol Me.ado, False
        Lokkol Me.allampolg, True
    Else
        Lokkol Me.szemelyi, False
        Lokkol Me.ado, True
        Lokkol Me.allampolg, False
    End If
End Sub

Private Sub Command1_Click()
    partner_lap.modosit elado_lista.ItemData(elado_lista.ListIndex), 20
End Sub

Private Sub elado_lista_Change()
On Error Resume Next
    BetoltElado elado_lista.ItemData(elado_lista.ListIndex)
End Sub

Private Sub elado_lista_Click()
    elado_lista_Change
End Sub

Private Sub elado_lista_Validate(Cancel As Boolean)
    elado_lista_Change
End Sub

Private Sub felvesz_Click()
 Dim p As String, i As Long
 
 'Ellenõrzések
    If Trim(bon_szam.Text) = "" Or Len(bon_szam.Text) < 3 Then GoTo hianyos
    If Trim(nyszam.Text) = "" Or Len(nyszam.Text) < 3 Then GoTo hianyos
 
 
 'Munka
 Dim Sor As New ADODB.Recordset
    'MsgBox all
    'Automatikus tanulás
        TanuldMeg "szinek", szine.Text
        TanuldMeg "allamjelzes", allam.Text
 
     If all = 0 Then
      If Not LetezikIlyen("autok", "nyszam", nyszam.Text) Then
        'Autó felvétele
        p = "INSERT INTO autok (allapot, eredet, uid, magan, nyszam, datum, ido, elado, marka, tipus, kategoria, evjarat, alvaz, motor, motorkod, szine, allam, henger, hajtoanyag, rendszam, torzskonyv, forgalmi, kivonas, kitiltas, ar, tomeg, bon_szam, bon_forg, bon_azon, x, hely, megj) VALUES " & "(1, '" & Chr(eredet.ItemData(eredet.ListIndex)) & "', " & Fid & "," & Alakit(azonosito(0).Value, "TRUE", "FALSE") & ", '" & nyszam.Text & "', '" & datum.Text & "', '" & ido.Text & "', " & CStr(Elado) & _
        ", " & marka_lista.ItemData(marka_lista.ListIndex) & ", " & tipus_lista.ItemData(tipus_lista.ListIndex) & ", " & kategoria.ItemData(kategoria.ListIndex) & ", '" & evjarat.Text & "', '" & alvaz.Text & "', '" & motor.Text & "', '" & motorkod.Text & "', '" & szine.Text & "', '" & allam.Text & "', '" & henger.Text & "', '" & hajtoanyag.Text & "', '" & rendszam.Text & "', '" & torzskonyv.Text & "', '" & forgalmi.Text & "', '" & kivonas.Text & "', '" & kitiltas.Text & "', " & ar.Text & ", " & tomeg.Text & ", '" & bon_szam.Text & "', '" & bon_forg.Text & "', '" & bon_azon.Text & "'," & Ix & ", '', '" & megj.Text & "')"
        SQL_p p, Sor
        
        'Nyilvántartási szám növelése, ha az az ajánlott
        If (nyszam.Text = NySzamAjanlo("szam_a")) Then Novel "bonto", "id", "1", "szam_a"
        
        'Frissen beszúrt kocsi idje
        p = "SELECT id FROM autok WHERE uid=" & Fid & " and nyszam='" & nyszam.Text & "' and marka=" & marka_lista.ItemData(marka_lista.ListIndex) & " and tipus=" & tipus_lista.ItemData(tipus_lista.ListIndex) & " and kategoria=" & kategoria.ItemData(kategoria.ListIndex) & " and evjarat='" & evjarat.Text & "' and alvaz='" & alvaz.Text & "' and motor='" & motor.Text & "' and rendszam='" & rendszam.Text & "' and bon_szam='" & bon_szam.Text & "' order by id desc"
        SQL_p p, Sor
        Sor.MoveFirst
        i = Sor!Id
        Sor.Close
        
        'Felvétel a hulladékok közé
        'SQL_p "INSERT INTO hulladek (auto, ewc, suly) VALUES ( " & i & ", 0 , " & tomeg.Text & ")", Sor
        SQL_p "INSERT INTO raktarkeszlet (tipus, auto, ewc, tszaz, suly) VALUES ( 2," & i & ", 0 , 100, " & tomeg.Text & ")", Sor
        'If Kinek > 0 Then Visszajelez Kinek, i
        auto_nyomtatas.mutat i, Kinek
        
        If (LCase(Trim(bon_szam)) <> "nincs" And Trim(bon_szam) <> "-") Then
            SQL_p "INSERT INTO kiadfelhnap (bon_szam, datum, allam, rendszam, alvaz, forgalmi, torzskonyv, auto) VALUES ( '" & bon_szam.Text & "', '" & datum.Text & "', '" & allam.Text & "', '" & rendszam.Text & "', '" & alvaz.Text & "', '" & forgalmi.Text & "', '" & torzskonyv.Text & "', " & i & ")", Sor
        End If
        Unload Me
     Else
        MsgBox "Már létezik ilyen nyilvántartási számú autó! Kérem adjon egy másikat!", vbCritical
     End If
    Else
         p = "UPDATE autok SET uid=" & Fid & ", allapot=1, eredet='" & Chr(eredet.ItemData(eredet.ListIndex)) & "', magan=" & Alakit(azonosito(0).Value, "TRUE", "FALSE") & ", nyszam='" & nyszam.Text & _
            "', datum='" & datum.Text & "', ido='" & ido.Text & "', elado=" & Elado & ", marka=" & marka_lista.ItemData(marka_lista.ListIndex) & ", tipus=" & tipus_lista.ItemData(tipus_lista.ListIndex) & ", kategoria=" & kategoria.ItemData(kategoria.ListIndex) & ", evjarat='" & evjarat.Text & "', alvaz='" & alvaz.Text & "', motor='" & motor.Text & "', motorkod='" & motorkod.Text & "', szine='" & szine.Text & "', allam='" & allam.Text & "', henger='" & henger.Text & "', hajtoanyag='" & hajtoanyag.Text & "', rendszam='" & rendszam.Text & "', torzskonyv='" & torzskonyv.Text & "', forgalmi='" & forgalmi.Text & "', kivonas='" & kivonas.Text & "', kitiltas='" & kitiltas.Text & "', ar=" & ar.Text & ", tomeg=" & tomeg.Text & ", bon_szam='" & bon_szam.Text & "', bon_forg='" & bon_forg.Text & "', bon_azon='" & bon_azon.Text & "', x=" & Ix & ", hely='', megj='" & megj.Text & "' where id=" & SID
            'MsgBox p
            SQL_p p, Sor
            
            'If Kinek > 0 Then Visszajelez Kinek, CInt(SID)
            auto_nyomtatas.mutat SID, Kinek
        If all = 3 Then
            SQL_p "UPDATE kiadfelhnap SET rontott=TRUE where auto=" & SID, Sor
            SQL_p "INSERT INTO kiadfelhnap (bon_szam, datum, allam, rendszam, alvaz, forgalmi, torzskonyv, auto) VALUES ( '" & bon_szam.Text & "', '" & datum.Text & "', '" & allam.Text & "', '" & rendszam.Text & "', '" & alvaz.Text & "', '" & forgalmi.Text & "', '" & torzskonyv.Text & "', " & i & ")", Sor
            Unload Me
        Else
            'SQL_p "INSERT INTO kiadfelhnap (bon_szam, datum, allam, rendszam, alvaz, forgalmi, torzskonyv) VALUES ( '" & bon_szam.Text & "', '" & datum.Text & "', '" & allam.Text & "', '" & rendszam.Text & "', '" & alvaz.Text & "', '" & forgalmi.Text & "', '" & torzskonyv.Text & "')", Sor
            SQL_p "UPDATE kiadfelhnap SET bon_szam='" & bon_szam.Text & "', datum='" & datum.Text & "', allam='" & allam.Text & "', rendszam='" & rendszam.Text & "', alvaz='" & alvaz.Text & "', forgalmi='" & forgalmi.Text & "', torzskonyv='" & torzskonyv.Text & "'where auto=" & SID, Sor
            Unload Me
        End If
    End If
    
Exit Sub
hianyos:
    MsgBox "Valamelyik azonosító mezõt elfelejtette kitölteni!", vbInformation, "Hiányos kitöltés!"
End Sub

Private Sub forgalmi_Click()
    forgalmi.Text = UCase(forgalmi.Text)
End Sub

Private Sub Form_Initialize()
On Error Resume Next
    felso = bontoware.zold
    
    Kinek = 0
    all = 0
    Frissit
    Elado = elado_lista.ItemData(0)
    JelolElado (Elado)
    EvjaratListaba Me.evjarat
    Betolt Me.szine, "szinek", "nev", "nev"
    Betolt Me.kategoria, "kategoria", "nev", "id"
    Betolt Me.hajtoanyag, "uzemanyag", "nev", "id"
    Betolt Me.allam, "allamjelzes", "nev", "id"
    
    eredet.List(0) = "L - lakosági"
    eredet.ItemData(0) = 76
    eredet.List(1) = "T - termelõi/intézményi"
    eredet.ItemData(1) = 84
    eredet.List(2) = "M - egyéb termelõi"
    eredet.ItemData(2) = 77
    eredet.List(3) = "I - import"
    eredet.ItemData(3) = 73
    eredet.ListIndex = 0
    
    allam.Text = "HU"
    datum.Text = Date
    ido.Text = Time
    
    x(0).Value = True
    azonosito(0).Value = True
    
    nyszam.Text = NySzamAjanlo("szam_a")
End Sub

Public Sub modosit(Id As Long, Optional Hova As Byte, Optional bonron As Byte)
    Dim kerdes As Byte
    
    
    kerdes = MsgBox("Biztos módosítani akarja az autó adatait?", vbYesNoCancel + vbCritical, "Utolsó megerõsítés")
    If kerdes <> vbYes Then
        Unload Me
        Exit Sub
    End If
    
    'Tényleges módosítás
    Dim Sor As New ADODB.Recordset
    Form_Initialize
    
    SQL_p "SELECT * from autok where id=" & Id, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        SID = Id
        If bonron = 1 Then
            all = 3
        Else
            all = 1
        End If
        Kinek = Hova
        felvesz.Caption = "Módosít"
        
        JelolElado (Nstr(Sor!Elado))
        JelolMarka (Nstr(Sor!marka))
        JelolTipus (Nstr(Sor!tipus))
        
        
        evjarat.Text = Nstr(Sor!evjarat)
        rendszam.Text = Nstr(Sor!rendszam)
        allam.Text = Nstr(Sor!allam)
        alvaz.Text = Nstr(Sor!alvaz)
        motor.Text = Nstr(Sor!motor)
        motorkod.Text = Nstr(Sor!motorkod)
        szine.Text = Nstr(Sor!szine)
        tomeg.Text = Nstr(Sor!tomeg)
        henger.Text = Nstr(Sor!henger)
        hajtoanyag.Text = Nstr(Sor!hajtoanyag)
        
        On Error Resume Next
        Jelol Me.eredet, CLng(Asc(Nstr(Sor!eredet)))
        'CLng(Asc(Nstr(Sor!eredet)))
        
        torzskonyv.Text = Nstr(Sor!torzskonyv)
        forgalmi.Text = Nstr(Sor!forgalmi)
        bon_forg.Text = Nstr(Sor!bon_forg)
        kitiltas.Text = Nstr(Sor!kitiltas)
        kivonas.Text = Nstr(Sor!kivonas)
        
        bon_szam.Text = Alakit(Nstr(Sor!bon_szam) = "", "nincs", Nstr(Sor!bon_szam))
        nyszam.Text = Nstr(Sor!nyszam)
        datum.Text = Nstr(Sor!datum)
        ido.Text = Nstr(Sor!ido)
        
        megj.Text = Nstr(Sor!megj)
        ar.Text = Nstr(Sor!ar)
        Sor.Close
        Me.Show
    Else
        MsgBox "Nincs ilyen rekord!"
        Unload Me
    End If
End Sub
Public Sub uj(Optional Hova As Byte)
    Form_Initialize
    'MsgBox all
    Kinek = Hova
    Me.Show
End Sub

Public Sub BeszurPartner(Id As Long)
    Elado_Frissit
    JelolElado Id
End Sub

Private Sub marka_lista_Change()
    marka = marka_lista.ItemData(marka_lista.ListIndex)
    Betolt Me.tipus_lista, "tipusok", "tipus", "tipus", , " where marka=" & marka
End Sub

Private Sub marka_lista_Click()
    marka_lista_Change
End Sub

Private Sub marka_lista_Validate(Cancel As Boolean)
    marka_lista_Change
End Sub

Private Sub megse_Click()
    Unload Me
End Sub


Private Sub motor_LostFocus()
    motor.Text = UCase(motor.Text)
End Sub

Private Sub motorkod_LostFocus()
   motorkod.Text = UCase(motorkod.Text)
End Sub

Private Sub rendszam_LostFocus()
    rendszam.Text = UCase(rendszam.Text)
End Sub

Private Sub szine_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Then szine.Text = ""
End Sub

Private Sub tomeg_GotFocus()
    tomeg.SelStart = 0
    tomeg.SelLength = Len(tomeg.Text)
End Sub

Private Sub torzskonyv_Click()
    torzskonyv.Text = UCase(torzskonyv.Text)
End Sub

Private Sub uj_elado_Click()
    partner_lap.uj 20
End Sub
Private Sub JelolElado(Id As Long)
    Jelol Me.elado_lista, Id
    elado_lista_Change
End Sub
Private Sub JelolMarka(Id As Long)
    Jelol Me.marka_lista, Id
    marka_lista_Change
End Sub
Private Sub JelolTipus(Id As Long)
    Jelol Me.tipus_lista, Id
    'tipus_lista_change
End Sub

Private Function Erteke(Minek As ComboBox) As Long
    Dim i As Long
    i = Minek.ListCount - 1
    Do While Minek.List(i) <> Minek.Text And i > 0
        i = i - 1
    Loop
    Erteke = Minek.ItemData(i)
End Function

Private Sub ujmtip_Click()
    markak.Show vbModal
    Marka_Frissit
End Sub
Private Function Ix() As Byte
    'Ix = 100 * x(0).Value + 10 * CInt(x(1).Value) + CInt(x(2).Value)
    Ix = 100 * Alakit(x(0).Value, 1, 0) + 10 * Alakit(x(1).Value, 1, 0) + Alakit(x(2).Value, 1, 0)
    'Ix = 100 * x(0).Value + 10 * CInt(x(1).Value) + CInt(x(2).Value)
End Function
Private Function FAzon() As String
    If azonosito(0).Value Then FAzon = "S" Else FAzon = "A"
End Function
Public Sub torol(Id As Long, Optional Hova As Byte)
    Dim kerdes As Byte
    
    kerdes = MsgBox("Biztos törölni akarja az autót?", vbYesNoCancel + vbCritical, "Utolsó megerõsítés")
    
    If kerdes = vbYes Then
        Dim Sor As New ADODB.Recordset
        SQL_p "SELECT * FROM raktarkeszlet WHERE (tipus=0 or tipus=1 or tipus=3) and auto=" & Id, Sor
        If Sor.RecordCount = 0 Then
            FSQL "DELETE * FROM autok WHERE id=" & Id
            FSQL "DELETE * FROM raktarkeszlet WHERE auto=" & Id
            
            Visszajelez Hova, 1
        Else
            MsgBox "Az autó a bontás ezen fázisában már nem törölhetõ!"
            'visszajele
        End If
    End If
    Unload Me
End Sub

Private Sub x_Click(Index As Integer)
    If Index = 0 Then
        Lokkol Me.bon_forg, False, True
        Lokkol Me.bon_azon, False, True
        
        elado_lista_Change
    Else
        Lokkol Me.bon_forg, True, True
        Lokkol Me.bon_azon, True, True
    End If
End Sub
