VERSION 5.00
Begin VB.Form auto_nyomtatas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nyomtatás"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Tag             =   "5"
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   4785
      TabIndex        =   5
      Top             =   0
      Width           =   4815
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adatok nyomatása"
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
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   2580
      End
   End
   Begin VB.CommandButton rontva 
      Caption         =   "Bontási átvételi igazolást elrontottam"
      Height          =   735
      Left            =   2520
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton allapotlap_nyom 
      Caption         =   "Üres állapotlap és szárazrafektetési napló nyomtatása"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton modosit 
      Caption         =   "Adatok módosítása"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton bontasi 
      Caption         =   "Bontási átvételi igazolás nyomtatása"
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton adasveteli 
      Caption         =   "Adásvételi szerzõdés nyomtatása"
      Default         =   -1  'True
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "auto_nyomtatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Kinek As Byte
Dim SID As Long


Public Sub mutat(Kit As Long, Optional Hova As Byte)
    Dim Sor As New ADODB.Recordset
    Form_Initialize
    SID = Kit
    Kinek = Hova
    
    
    SQL_p "SELECT NYSZAM FROM autok WHERE ID=" & SID, Sor
    ElnevezAblak Me, "(" & Sor!nyszam & ")"
    Sor.Close
    
    Me.Show 'vbModal
End Sub

Private Sub adasveteli_Click()
    Dim p As String
    Dim Sor As New ADODB.Recordset
     
    p = "SELECT autok.nyszam AS NYSZAM, autok.magan as MAGAN, autok.bon_szam AS BONSZAM, partnerek!vnev+' '+partnerek!knev AS ENEV, partnerek.irszam+' '+partnerek.varos+', '+partnerek.cim AS ECIM, partnerek.szemelyi AS ESZEM, partnerek.ado AS EADO, bonto.nev AS BNEV, bonto.varos AS BVAROS, bonto.irszam AS BIRSZAM, [bonto].[utca]+' '+[bonto].[hazszam]+'.' AS BCIM, bonto.cg AS BCG, bonto.ado AS BADO, bonto.felelos AS BKEPV, bonto.knysz AS KFFNYSZ, bonto.telepeng AS TELEPENG, bonto.vhull AS VESZHULL, markak.marka AS GYARTMANY, tipusok.tipus AS TIPUS, autok.evjarat AS GYEV, autok.torzskonyv AS TORZSKONYV, autok.alvaz AS ALVAZ, autok.motor AS MOTOR, autok.forgalmi AS FORGALMI, autok.rendszam AS RENDSZAM, autok.allam AS ALLAM, autok.kivonas AS KIVONAS, kategoria.kod AS KATEGORIA, autok.tomeg AS TOMEG, autok.henger AS HENGER, autok.ar AS AR " & _
        "FROM bonto, (markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN (partnerek INNER JOIN (kategoria INNER JOIN autok ON kategoria.id = autok.kategoria) ON partnerek.id = autok.elado) ON tipusok.id = autok.tipus " & _
        "WHERE autok.id=" & SID
    
    nyomtatasikep.Kitolt2 "adasveteli.htm", p
End Sub

Private Sub allapotlap_nyom_Click()
    Fejl
End Sub

Private Sub bezar_Click()
    Unload Me
End Sub

Private Sub bontasi_Click()
    Dim p As String
    Dim Sor As New ADODB.Recordset
     
    p = "SELECT autok.id, bonto.vhull AS KVSZAM, bonto.vhullerv AS KVERV, autok.rendszam AS RENDSZAM, autok.allam AS ALLAM, autok.alvaz AS ALVAZ, kategoria.kod AS KATEGORIA, markak.marka AS GYARTMANY, tipusok.tipus AS TIPUS, autok.evjarat AS EVJARAT, autok.hajtoanyag AS UZEMANYAG, autok.torzskonyv AS TORZSKONYV, autok.forgalmi AS FORGALMI, partnerek.vnev AS VNEV, partnerek.knev AS KNEV, partnerek.irszam+' '+partnerek.varos+', '+partnerek.cim AS CIM, autok.bon_forg AS FORGJOG, autok.x AS IX, autok.bon_azon AS SZEMIG, partnerek.allampolg AS ALLAMPOLG, bonto.varos AS VAROS, autok.datum AS DATUM, autok.ido AS IDO " & _
        "FROM bonto, partnerek INNER JOIN (kategoria INNER JOIN ((markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN autok ON tipusok.id = autok.tipus) ON kategoria.id = autok.kategoria) ON partnerek.id = autok.elado " & _
        "WHERE (((autok.id)=" & SID & "));"
        
    SQL_p p, Sor
    'MsgBox Sor.RecordCount
    nyomtatasikep.Kitolt "bontasi.htm", Sor
    Sor.Close
End Sub

Private Sub Form_Initialize()
    felso = bontoware.narancs
    Kinek = 0
    SID = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload nyomtatasikep
    Visszajelez Kinek, SID
End Sub

Private Sub megbiz_Click()

End Sub

Private Sub meghat_Click()

End Sub

Private Sub modosit_Click()
    auto.modosit SID, Kinek
End Sub

Private Sub rontva_Click()
    auto.modosit SID, Kinek, 1
End Sub
