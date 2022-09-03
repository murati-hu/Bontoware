VERSION 5.00
Begin VB.Form felauto_lap 
   Caption         =   "Szárazrafektetett autó importálása"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7305
      TabIndex        =   50
      Top             =   0
      Width           =   7335
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Szárazrafektetett autó felvétele"
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
         Left            =   2880
         TabIndex        =   51
         Top             =   240
         Width           =   4260
      End
   End
   Begin VB.CommandButton felvesz_uj 
      Caption         =   "Felvesz és újat kezd"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   8040
      Width           =   1815
   End
   Begin VB.TextBox szamla 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox nyszam 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Frame keret 
      Caption         =   "Eladó adatai"
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
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   1200
      Width           =   6975
      Begin VB.CommandButton Command1 
         Caption         =   "Módosít"
         Height          =   375
         Left            =   6000
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox cim 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   720
         Width           =   5895
      End
      Begin VB.CommandButton uj_elado 
         Caption         =   "Új"
         Height          =   375
         Left            =   5400
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox elado_lista 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox ado 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1080
         Width           =   5895
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eladó:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   450
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adószám:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cím:"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   40
         Top             =   720
         Width           =   315
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
      Height          =   4935
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   2880
      Width           =   6975
      Begin VB.TextBox ar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Text            =   "0"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox ido 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         TabIndex        =   17
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox datum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   16
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox hajtoanyag 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox henger 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         TabIndex        =   14
         Top             =   2160
         Width           =   615
      End
      Begin VB.ComboBox marka_lista 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   3615
      End
      Begin VB.ComboBox tipus_lista 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   3615
      End
      Begin VB.ComboBox szine 
         Height          =   315
         ItemData        =   "felauto_lap.frx":0000
         Left            =   960
         List            =   "felauto_lap.frx":0019
         TabIndex        =   12
         Top             =   2160
         Width           =   1695
      End
      Begin VB.ComboBox kategoria 
         Height          =   315
         ItemData        =   "felauto_lap.frx":004E
         Left            =   960
         List            =   "felauto_lap.frx":0050
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
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
         Top             =   1440
         Width           =   3615
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
         TabIndex        =   9
         Top             =   1800
         Width           =   3615
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
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox tomeg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   13
         Text            =   "0"
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton ujmtip 
         Caption         =   "Új márka vagy típus"
         Height          =   255
         Left            =   4800
         TabIndex        =   24
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox allam 
         Height          =   315
         ItemData        =   "felauto_lap.frx":0052
         Left            =   5640
         List            =   "felauto_lap.frx":0054
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox megj 
         Appearance      =   0  'Flat
         Height          =   1485
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   3240
         Width           =   6495
      End
      Begin VB.ComboBox evjarat 
         Height          =   315
         Left            =   5640
         TabIndex        =   5
         Text            =   "evjarat"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ár:"
         Height          =   195
         Index           =   17
         Left            =   600
         TabIndex        =   47
         Top             =   2640
         Width           =   195
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dátum:"
         Height          =   195
         Index           =   26
         Left            =   2520
         TabIndex        =   46
         Top             =   2640
         Width           =   510
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Idõ:"
         Height          =   195
         Index           =   15
         Left            =   5040
         TabIndex        =   45
         Top             =   2640
         Width           =   270
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hajtóanyag:"
         Height          =   195
         Index           =   28
         Left            =   4680
         TabIndex        =   44
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hengerûrtartalom:"
         Height          =   195
         Index           =   10
         Left            =   4800
         TabIndex        =   43
         Top             =   2160
         Width           =   1260
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gyártmány:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   795
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Típus:"
         Height          =   195
         Index           =   21
         Left            =   480
         TabIndex        =   36
         Top             =   720
         Width           =   450
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Színe:"
         Height          =   195
         Index           =   22
         Left            =   360
         TabIndex        =   35
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kategória:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Államjelzés:"
         Height          =   195
         Index           =   9
         Left            =   4800
         TabIndex        =   33
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alvázszám:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motorszám:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motorkód:"
         Height          =   195
         Index           =   8
         Left            =   4800
         TabIndex        =   30
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Évjárat:"
         Height          =   195
         Index           =   5
         Left            =   4920
         TabIndex        =   29
         Top             =   720
         Width           =   540
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saját tömege:"
         Height          =   195
         Index           =   18
         Left            =   2760
         TabIndex        =   28
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Megjegyzés:"
         Height          =   195
         Index           =   19
         Left            =   240
         TabIndex        =   27
         Top             =   3000
         Width           =   885
      End
   End
   Begin VB.CommandButton felvesz_zar 
      Caption         =   "Felvétel és bezár"
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton megse 
      Caption         =   "Mégse"
      Height          =   375
      Left            =   5760
      TabIndex        =   21
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Számlaszám:"
      Height          =   195
      Index           =   2
      Left            =   3360
      TabIndex        =   49
      Top             =   840
      Width           =   915
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nyilvántartási szám:"
      Height          =   195
      Index           =   23
      Left            =   240
      TabIndex        =   48
      Top             =   840
      Width           =   1395
   End
End
Attribute VB_Name = "felauto_lap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim all As Byte '0új, 1mód, 2besz
Dim SID As Long
Dim Elado As Long
Dim marka As Long
Dim tipus As Long
Dim Jarmu As Long
Dim Kinek As Byte

Public Sub Frissit()
    Elado_Frissit
    Marka_Frissit
    'Tipus_Frissit
End Sub
Public Sub Elado_Frissit()
    partner.Partner_Listaba Me.elado_lista
End Sub
Public Sub Marka_Frissit()
'On Error GoTo hiba
Dim i As Long, p As String
    'Autómárkák listája
    marka_lista.Clear
    Rekord.CursorLocation = adUseClient
    sql_parancs ("SELECT * FROM markak order by marka")
    If Not Rekord.EOF Then Rekord.MoveFirst
    marka_lista.Visible = False
    i = 0
    Do While Not Rekord.EOF
        marka_lista.List(i) = Nstr(Rekord!marka)
        marka_lista.ItemData(i) = Rekord!Id
        Rekord.MoveNext
        i = i + 1
    Loop
    marka_lista.Visible = True
    Rekord.Close
    ElsotJelol Me.marka_lista
Exit Sub
Hiba:
    Hiba Err.Number, "Frissitési hiba"
End Sub
Public Sub Tipus_Frissit()
'On Error GoTo hiba
Dim i As Long, p As String
    'Autómárkák listája
    tipus_lista.Clear
    Rekord.CursorLocation = adUseClient
    sql_parancs ("SELECT * FROM tipusok where marka=" & Erteke(marka_lista) & " order by tipus")
    If Not Rekord.EOF Then Rekord.MoveFirst
    tipus_lista.Visible = False
    i = 0
    Do While Not Rekord.EOF
        tipus_lista.List(i) = Nstr(Rekord!tipus)
        tipus_lista.ItemData(i) = Rekord!Id
        Rekord.MoveNext
        i = i + 1
    Loop
    tipus_lista.Visible = True
    Rekord.Close
    ElsotJelol Me.tipus_lista
Exit Sub
Hiba:
    Hiba Err.Number, "Frissitési hiba"
End Sub
Public Sub BetoltElado(Id As Long)
    sql_parancs ("SELECT * from partnerek where id=" & Id)
    If Not Rekord.EOF Then
        Rekord.MoveFirst
        Elado = Id
        cim.Text = Nstr(Rekord!irszam & " " & Rekord!varos & " " & Rekord!cim)
        'szemelyi.Text = Nstr(Rekord!szemelyi)
        ado.Text = Nstr(Rekord!ado)
        'allampolg.Text = Nstr(Rekord!allampolg)
        Rekord.Close
    Else
        MsgBox "Nincs ilyen rekord!"
    End If
End Sub


Private Sub alvaz_LostFocus()
    alvaz.Text = UCase(alvaz.Text)
End Sub

Private Sub Command1_Click()
    partner_lap.modosit elado_lista.ItemData(elado_lista.ListIndex), 70
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

Private Sub felvesz()
 Dim p As String, i As Long
 Dim Sor As New ADODB.Recordset
     If all = 0 Then
        p = "INSERT INTO autok (allapot, bontva, magan, nyszam, datum, ido, elado, marka, tipus, kategoria, evjarat, alvaz, motor, motorkod, szine, allam, henger, hajtoanyag, ar, tomeg, hely, megj, szamla) VALUES " & "(2, TRUE, FALSE, '" & nyszam.Text & "', '" & datum.Text & "', '" & ido.Text & "', " & CStr(Elado) & _
        ", " & marka_lista.ItemData(marka_lista.ListIndex) & ", " & tipus_lista.ItemData(tipus_lista.ListIndex) & ", " & kategoria.ItemData(kategoria.ListIndex) & ", '" & evjarat.Text & "', '" & alvaz.Text & "', '" & motor.Text & "', '" & motorkod.Text & "', '" & szine.Text & "', '" & allam.Text & "', '" & henger.Text & "', '" & hajtoanyag.Text & "', " & ar.Text & ", " & tomeg.Text & ", '', '" & megj.Text & "', '" & szamla.Text & "')"
        SQL_p p, Sor
        
        'Nyilvántartási szám növelése, ha az az ajánlott
        If (nyszam.Text = NySzamAjanlo("szam_a")) Then Novel "bonto", "id", "1", "szam_a"
        
        'FRissen beszúrt kocsi idje
        p = "SELECT id FROM autok WHERE nyszam='" & nyszam.Text & "' and marka=" & marka_lista.ItemData(marka_lista.ListIndex) & " and tipus=" & tipus_lista.ItemData(tipus_lista.ListIndex) & " and kategoria=" & kategoria.ItemData(kategoria.ListIndex) & " and evjarat='" & evjarat.Text & "' and alvaz='" & alvaz.Text & "' and motor='" & motor.Text & "' and tomeg=" & tomeg.Text & " and szamla='" & szamla.Text & "' " ' order by id desc"
        SQL_p p, Sor
        Sor.MoveFirst
        i = Sor!Id
        Sor.Close
        
        'Felvétel a hulladékok közé
        'Régi
        'SQL_p "INSERT INTO hulladek (auto, ewc, suly) VALUES ( " & i & ", 1 , " & tomeg.Text & ")", Sor
        SQL_p "INSERT INTO raktarkeszlet (tipus, auto, ewc, tszaz, suly) VALUES ( 1," & i & ", 1 , 100, " & tomeg.Text & ")", Sor
        If Kinek > 0 Then Visszajelez Kinek, i
    Else
        'p = "UPDATE autok SET allapot=1, flag='" & FAzon & "', nyszam='" & nyszam.Text & "', datum='" & datum.Text & "', ido='" & ido.Text & "', elado=" & Elado & ", marka=" & marka_lista.ItemData(marka_lista.ListIndex) & ", tipus=" & tipus_lista.ItemData(tipus_lista.ListIndex) & ", kategoria=" & kategoria.ItemData(kategoria.ListIndex) & ", evjarat='" & evjarat.Text & "', alvaz='" & alvaz.Text & "', motor='" & motor.Text & "', motorkod='" & motorkod.Text & "', szine='" & szine.Text & "', allam='" & allam.Text & "', henger='" & henger.Text & "', hajtoanyag='" & hajtoanyag.Text & "', rendszam='" & rendszam.Text & "', torzskonyv='" & torzskonyv.Text & "', forgalmi='" & forgalmi.Text & "', kivonas='" & kivonas.Text & "', kitiltas='" & kitiltas.Text & "', ar=" & ar.Text & ", tomeg=" & tomeg.Text & ", bon_szam='" & bon_szam.Text & "', bon_forg='" & bon_forg.Text & "', x=" & Ix & ", hely='', megj='" & megj.Text & "' where id=" & SID
        'MsgBox p
        SQL_p p, Sor
        
        If Kinek > 0 Then Visszajelez Kinek, CInt(SID)
    End If
    Unload Me
End Sub

Private Sub felvesz_uj_Click()
    felvesz
    uj Kinek
End Sub

Private Sub felvesz_zar_Click()
    felvesz
    Unload Me
End Sub

Private Sub Form_Initialize()
On Error Resume Next
    felso = bontoware.zold
    Kinek = 0
    Frissit
    Elado = elado_lista.ItemData(0)
    JelolElado (Elado)
    EvjaratListaba Me.evjarat
    Betolt Me.szine, "szinek", "nev", "nev"
    Betolt Me.kategoria, "kategoria", "nev", "id"
    Betolt Me.hajtoanyag, "uzemanyag", "nev", "id"
    Betolt Me.allam, "allamjelzes", "nev", "id"
    allam.Text = "HU"
    datum.Text = Date
    ido.Text = Time
    
    nyszam.Text = NySzamAjanlo("szam_a")
    szamla.Text = ""
End Sub

Public Sub modosit(Id As Long, Optional Hova As Byte)
    Dim Sor As New ADODB.Recordset
    Form_Initialize
    
    SQL_p "SELECT * from autok where id=" & Id, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        SID = Id
        all = 1
        Kinek = Hova
        'felvesz.Caption = "Módosít"
        
        JelolElado (Nstr(Sor!Elado))
        JelolMarka (Nstr(Sor!marka))
        JelolTipus (Nstr(Sor!tipus))
        
        
        evjarat.Text = Nstr(Sor!evjarat)
        'rendszam.Text = Nstr(Sor!rendszam)
        allam.Text = Nstr(Sor!allam)
        alvaz.Text = Nstr(Sor!alvaz)
        motor.Text = Nstr(Sor!motor)
        motorkod.Text = Nstr(Sor!motorkod)
        szine.Text = Nstr(Sor!szine)
        tomeg.Text = Nstr(Sor!tomeg)
        henger.Text = Nstr(Sor!henger)
        hajtoanyag.Text = Nstr(Sor!hajtoanyag)
        
       
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
    Kinek = Hova
    Me.Show
End Sub

Public Sub BeszurPartner(Id As Long)
    Elado_Frissit
    JelolElado Id
End Sub

Private Sub marka_lista_Change()
    Tipus_Frissit
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

Private Sub szine_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Then szine.Text = ""
End Sub

Private Sub uj_elado_Click()
    partner_lap.uj 70
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
    markak.Show
    Marka_Frissit
    Tipus_Frissit
End Sub
