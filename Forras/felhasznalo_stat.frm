VERSION 5.00
Begin VB.Form felhasznalo_stat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Felhasználó Statisztika"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   4425
      TabIndex        =   7
      Top             =   0
      Width           =   4455
      Begin VB.Label felhasznalo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "felhasznalo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Label sztor_szamlak 
      AutoSize        =   -1  'True
      Caption         =   "0/0"
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Sztornózott számlák:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label autok 
      AutoSize        =   -1  'True
      Caption         =   "0/0"
      Height          =   195
      Left            =   3360
      TabIndex        =   4
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Felvett gépjármûvek:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label szamlak 
      AutoSize        =   -1  'True
      Caption         =   "0/0"
      Height          =   195
      Left            =   3360
      TabIndex        =   2
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Kiállított számlák:"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label jogok 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "felhasznalo_stat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SID As Long


Private Sub Frissit()
    Dim Sor As New ADODB.Recordset, p As String
    
    felhasznalo.Caption = Ertek("felhasznalok", "id", CStr(SID), "nev")
    
    'Számlák számolása
    SQL_p "SELECT Count(szamla.uid) AS CountOfuid From szamla WHERE (((szamla.uid)=" & SID & "))", Sor
    szamlak.Caption = Sor.Fields(0)
    Sor.Close
    
    SQL_p "SELECT Count(szamla.uid) AS CountOfuid From szamla", Sor
    szamlak.Caption = szamlak.Caption & "/" & Sor.Fields(0)
    Sor.Close
    
    'Sztornózott számlák összeszámolása
    SQL_p "SELECT Count(szamla.uid) AS CountOfuid From szamla WHERE (((szamla.uid)=" & SID & ") AND( ((szamla.tipus)=1) OR ((szamla.tipus)=3)))", Sor
    sztor_szamlak.Caption = Sor.Fields(0)
    Sor.Close
    
    SQL_p "SELECT Count(szamla.uid) AS CountOfuid From szamla WHERE (((szamla.tipus)=1) OR ((szamla.tipus)=3))", Sor
    sztor_szamlak.Caption = sztor_szamlak.Caption & "/" & Sor.Fields(0)
    Sor.Close
    
    'Felvett autók
    SQL_p "SELECT Count(autok.uid) AS CountOfuid From autok WHERE (((autok.uid)=" & SID & "))", Sor
    autok.Caption = Sor.Fields(0)
    Sor.Close
    
    SQL_p "SELECT Count(autok.uid) AS CountOfuid From autok", Sor
    autok.Caption = autok.Caption & "/" & Sor.Fields(0)
    Sor.Close
End Sub

Public Sub Mutasd(Optional Kit As Integer)
    SID = Kit
    Frissit
    Me.Show vbModal
End Sub

Private Sub Form_Load()
    felso = bontoware.lila
End Sub
