VERSION 5.00
Begin VB.Form bontoware 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Bonto-ware"
   ClientHeight    =   10170
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   12285
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox gomb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2685
      Index           =   1
      Left            =   2760
      Picture         =   "bontoware.frx":0000
      ScaleHeight     =   2685
      ScaleWidth      =   2055
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3600
      Width           =   2055
   End
   Begin VB.PictureBox alap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11520
      Left            =   0
      Picture         =   "bontoware.frx":24CE
      ScaleHeight     =   11520
      ScaleWidth      =   15360
      TabIndex        =   5
      Top             =   -720
      Width           =   15360
      Begin VB.PictureBox gomb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2685
         Index           =   8
         Left            =   9480
         Picture         =   "bontoware.frx":21CF2
         ScaleHeight     =   2685
         ScaleWidth      =   2055
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1500
         Width           =   2055
      End
      Begin VB.PictureBox gomb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2685
         Index           =   7
         Left            =   9480
         Picture         =   "bontoware.frx":240F7
         ScaleHeight     =   2685
         ScaleWidth      =   2055
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   7200
         Width           =   2055
      End
      Begin VB.PictureBox gomb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2685
         Index           =   6
         Left            =   600
         Picture         =   "bontoware.frx":261A6
         ScaleHeight     =   2685
         ScaleWidth      =   2055
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   7200
         Width           =   2055
      End
      Begin VB.PictureBox gomb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2700
         Index           =   5
         Left            =   600
         Picture         =   "bontoware.frx":27FCE
         ScaleHeight     =   2700
         ScaleWidth      =   2055
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4320
         Width           =   2055
      End
      Begin VB.PictureBox gomb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2685
         Index           =   4
         Left            =   2760
         Picture         =   "bontoware.frx":2A83D
         ScaleHeight     =   2685
         ScaleWidth      =   2070
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2070
      End
      Begin VB.PictureBox gomb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2685
         Index           =   3
         Left            =   7200
         Picture         =   "bontoware.frx":2D048
         ScaleHeight     =   2685
         ScaleWidth      =   2055
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   4320
         Width           =   2055
      End
      Begin VB.PictureBox gomb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2685
         Index           =   2
         Left            =   9480
         Picture         =   "bontoware.frx":2F497
         ScaleHeight     =   2685
         ScaleWidth      =   2055
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   4320
         Width           =   2055
      End
      Begin VB.PictureBox gomb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DataSource      =   "gfd"
         ForeColor       =   &H80000008&
         Height          =   2685
         Index           =   0
         Left            =   600
         Picture         =   "bontoware.frx":313B3
         ScaleHeight     =   2685
         ScaleWidth      =   2055
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2055
      End
   End
   Begin VB.PictureBox lila 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   -5000
      Picture         =   "bontoware.frx":3382F
      ScaleHeight     =   1050
      ScaleWidth      =   4200
      TabIndex        =   4
      Top             =   5640
      Visible         =   0   'False
      Width           =   4230
   End
   Begin VB.PictureBox narancs 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   -5000
      Picture         =   "bontoware.frx":351DF
      ScaleHeight     =   1050
      ScaleWidth      =   4200
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   4230
   End
   Begin VB.PictureBox zold 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   -5000
      Picture         =   "bontoware.frx":36993
      ScaleHeight     =   1050
      ScaleWidth      =   4200
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   4230
   End
   Begin VB.PictureBox piros 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   -5000
      Picture         =   "bontoware.frx":3820C
      ScaleHeight     =   1050
      ScaleWidth      =   4200
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   4230
   End
   Begin VB.PictureBox kek 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   -5000
      Picture         =   "bontoware.frx":3A0B9
      ScaleHeight     =   1050
      ScaleWidth      =   4200
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   4230
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   240
   End
   Begin VB.Menu bejelentkezes_mnu 
      Caption         =   "Felhasználóváltás"
   End
   Begin VB.Menu felvetel_nagy_mnu 
      Caption         =   "&Felvétel"
      Begin VB.Menu ujauto_mnu 
         Caption         =   "Gépjármû felvétele bontásra"
      End
      Begin VB.Menu szarazrafektetett 
         Caption         =   "Import szárazrafektett gépjármû felvétele"
      End
      Begin VB.Menu alkfelv_kintrol 
         Caption         =   "Alkatrész felvétel gépjármûhöz"
         Visible         =   0   'False
      End
      Begin VB.Menu alkatresz_felvetel_mnu 
         Caption         =   "Alkatrész felvétele"
      End
      Begin VB.Menu uj_partner_mnu 
         Caption         =   "&Partner felvétel"
      End
   End
   Begin VB.Menu autok_mnu 
      Caption         =   "&Listák"
      Begin VB.Menu autolista_mnu 
         Caption         =   "Nyilvántartott gépjármûvek"
      End
      Begin VB.Menu mnszfa_mnu 
         Caption         =   "Még nem szárazrafektetett gépjármûvek"
      End
      Begin VB.Menu szamlalista_mnu 
         Caption         =   "Számlák"
      End
      Begin VB.Menu partner_listaja 
         Caption         =   "&Nyilvántartott partnerek"
      End
   End
   Begin VB.Menu leltar_fomnu 
      Caption         =   "Leltárak"
      Begin VB.Menu alvaz_letar 
         Caption         =   "Alváz leltár"
      End
      Begin VB.Menu motor_leltar_mnu 
         Caption         =   "Motor leltár"
      End
      Begin VB.Menu valto_leltar_mnu 
         Caption         =   "Sebességváltó leltár"
      End
   End
   Begin VB.Menu raktarkeszlet_mnu 
      Caption         =   "&Raktárkészlet és telep"
      Begin VB.Menu szamlazo_mnu 
         Caption         =   "&Alkatrész eladás"
      End
      Begin VB.Menu raktar_mnu 
         Caption         =   "&Jelenlegi raktárkészlet"
      End
      Begin VB.Menu v 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu hulladek_mnu 
         Caption         =   "Hulladék eladás"
      End
      Begin VB.Menu auto_selejt_mnu 
         Caption         =   "Több autó selejtezése"
      End
      Begin VB.Menu v0 
         Caption         =   "-"
      End
      Begin VB.Menu teleprend_mnu 
         Caption         =   "&Teleprendezés"
      End
   End
   Begin VB.Menu hull_lista_mnu 
      Caption         =   "Aktuális hulladéklista"
   End
   Begin VB.Menu statisztika_mnu 
      Caption         =   "&Sztatisztikák"
      Begin VB.Menu kiadas_mnu 
         Caption         =   "Kiadások"
      End
      Begin VB.Menu v1 
         Caption         =   "-"
      End
      Begin VB.Menu felh_stat_mnu 
         Caption         =   "Felhasználók"
      End
   End
   Begin VB.Menu uzemkonyv_mnu 
      Caption         =   "Naplók és Üzemkönyvek"
      Begin VB.Menu kiad_mnu 
         Caption         =   "Kiadási és felhasználási napló nyomtatása"
      End
      Begin VB.Menu atv_mnu 
         Caption         =   "KFF átadott gépjármûvek"
      End
      Begin VB.Menu v6 
         Caption         =   "-"
      End
      Begin VB.Menu bemeneti_uzk 
         Caption         =   "Bemeneti üzemkönyv"
      End
      Begin VB.Menu v5 
         Caption         =   "-"
      End
      Begin VB.Menu alt_kim_uz_mnu 
         Caption         =   "Általános Kimeneti üzemkönyv"
      End
      Begin VB.Menu v4 
         Caption         =   "-"
      End
      Begin VB.Menu napi_kim_mnu 
         Caption         =   "Napi kimeneti üzemkönyv"
      End
      Begin VB.Menu haki_kim_mnu 
         Caption         =   "Havi kimeneti üzemkönyv"
      End
   End
   Begin VB.Menu beall_mnu 
      Caption         =   "&Beállítások"
      Begin VB.Menu allapotlap_bov_mnu 
         Caption         =   "Állapotlap bõvítése"
      End
      Begin VB.Menu felh_kezel_mnu 
         Caption         =   "Felhasználók kezelése"
      End
      Begin VB.Menu markak_mnu 
         Caption         =   "&Márkák és típusok"
      End
      Begin VB.Menu ewc_lista_mnu 
         Caption         =   "&Bontási napló tételei"
      End
      Begin VB.Menu v3 
         Caption         =   "-"
      End
      Begin VB.Menu bonto_mnu 
         Caption         =   "Bontó adatai"
      End
   End
   Begin VB.Menu fejl_mnu 
      Caption         =   "..:::Fejlesztõi menü:::.."
      Begin VB.Menu ab_karbantarto 
         Caption         =   "Adatbázis karbantartás"
      End
      Begin VB.Menu szam_mod_mnu 
         Caption         =   "Alkatrész számla betöltése"
      End
      Begin VB.Menu modhull_mnu 
         Caption         =   "Hulladék számla betöltése"
      End
   End
End
Attribute VB_Name = "bontoware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aktualis As Integer
'Public Fid As Integer

Private Sub ab_karbantarto_Click()
    abk.Show
End Sub

Private Sub alap_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    GombElrendezesek True
End Sub

Private Sub alkatresz_felvetel_mnu_Click()
    alkatresz_lap.uj
End Sub

Private Sub alkfelv_kintrol_Click()
       ' kintrol.Show
End Sub

Private Sub allapotlap_bov_mnu_Click()
    bovites.Show
End Sub

Private Sub alt_kim_uz_mnu_Click()
    kimeneti_alt.meghiv
End Sub

Private Sub alvaz_letar_Click()
    auto_lista.Mutasd Alvaz_leltar
End Sub

Private Sub atv_mnu_Click()
    atvetel.Show
End Sub

Private Sub auto_selejt_mnu_Click()
    auto_selejt.Show
End Sub

Private Sub autolista_mnu_Click()
    auto_lista.Mutasd Teljes_lista
End Sub

Private Sub bejelentkezes_mnu_Click()
    Beleptet
End Sub

Private Sub bemeneti_uzk_Click()
    bemeneti.menubol_hiv
    'bemeneti.meghiv CDate(tol), CDate(ig)
End Sub

Private Sub bonto_mnu_Click()
    bonto.Show
End Sub


Private Sub ewc_lista_mnu_Click()
    ewc_lista.Show
End Sub

Private Sub felh_kezel_mnu_Click()
    felhasznalok.Show
End Sub

Private Sub felh_stat_mnu_Click()
    felhasznalo_stat.Mutasd Fid
End Sub

Private Sub Form_Initialize()
    adatmotor.Megnyitas
    Fid = 0
    Beleptet
    Me.Caption = ProgramNeve & " - " & App.Major & "." & NKieg(App.Minor) & "." & NKieg(App.Revision)
    
    GombElrendezesek
    'aktualis = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Biztos abban, hogy ki akar lépni a teljes programból?", vbQuestion + vbYesNo, "Kilépés megerõsítése") = vbYes Then
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub gomb_Click(Index As Integer)
    Select Case Index
        Case 0
            auto.uj
        Case 1
            szamlazo.uj
        Case 2
            hulladek_lista.Show
        Case 3
            hulladek_szamla.uj
        Case 4
            felauto_lap.uj
        Case 5
            alkatresz_lap.uj
        Case 6
            alt_kim_uz_mnu_Click
        Case 7
            kiad_felh_nap.Show
        Case 8
            auto_lista.Mutasd Teljes_lista
    End Select
    GombElrendezesek True
End Sub

Private Sub gomb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    If aktualis = Index Then Exit Sub
    GombElrendezesek True
    With gomb(Index)
        aktualis = Index
        .Move .DataField + 100, .Tag + 100
    End With
End Sub

Private Sub haki_kim_mnu_Click()
    kimeneti_havi.meghiv Date
End Sub

Private Sub hull_lista_mnu_Click()
    hulladek_lista.Show
End Sub

Private Sub hulladek_mnu_Click()
    hulladek_szamla.uj
End Sub

Private Sub kiad_mnu_Click()
    kiad_felh_nap.Show
End Sub

Private Sub GombElrendezesek(Optional Vissza As Boolean)
    Dim i As Byte
    For i = 0 To gomb.Count - 1
        If Vissza Then
            gomb(i).Move gomb(i).DataField, gomb(i).Tag
        Else
            gomb(i).DataField = gomb(i).Left
            gomb(i).Tag = gomb(i).Top
        End If
    Next i
    aktualis = -1
End Sub
Private Sub kiadas_mnu_Click()
    'kiadasok.meghiv
    Dim tol As String, ig As String
    If Month(Date) <= 3 Then
        tol = Year(Date) & ".01.01."
        ig = Year(Date) & ".03.31."
    Else
        If Month(Date) <= 6 Then
            tol = Year(Date) & ".04.01."
            ig = Year(Date) & ".06.30."
        Else
            If Month(Date) <= 9 Then
                tol = Year(Date) & ".07.01."
                ig = Year(Date) & ".09.30."
            Else
                tol = Year(Date) & ".10.01."
                ig = Year(Date) & ".12.31."
            End If
        End If
    End If
    
    kiadasok.meghiv CDate(tol), CDate(ig)
End Sub

Private Sub markak_mnu_Click()
    markak.Show
End Sub

Private Sub mnszfa_mnu_Click()
    auto_lista.Mutasd NemSzarazrafektetett_lista
End Sub

Private Sub modhull_mnu_Click()
    hulladek_szamla.modosit InputBox("p")
End Sub

Private Sub motor_leltar_mnu_Click()
    auto_lista.Mutasd Motor_leltar
End Sub

Private Sub napi_kim_mnu_Click()
    kimeneti.meghiv Date
End Sub


Private Sub partner_listaja_Click()
    partner_lista.Show
End Sub

Private Sub raktar_mnu_Click()
    alkatresz_lista.meghiv
End Sub

Private Sub szam_mod_mnu_Click()
    szamlazo.modosit InputBox("nyszam")
End Sub

Private Sub szamlalista_mnu_Click()
    szamla_lista.meghiv
End Sub

Private Sub szamlazo_mnu_Click()
    szamlazo.uj
End Sub

Private Sub szarazrafektetett_Click()
    felauto_lap.uj
End Sub

Private Sub teleprend_mnu_Click()
    raktar.Mutasd 1
End Sub

Private Sub Timer1_Timer()
    bontoware.Cls
    bontoware.Print Fid
End Sub

Private Sub uj_partner_mnu_Click()
    partner_lap.uj
End Sub

Private Sub ujauto_mnu_Click()
    auto.uj
End Sub

Private Sub valto_leltar_mnu_Click()
    auto_lista.Mutasd Sebvalto_leltar
End Sub
