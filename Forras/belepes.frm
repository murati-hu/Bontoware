VERSION 5.00
Begin VB.Form belepes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bejelentkezés"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   3825
      TabIndex        =   6
      Top             =   0
      Width           =   3855
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bejelentkezés"
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
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bejelentkezés"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton bezar 
      Caption         =   "Mégse"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox jelszo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "@"
      TabIndex        =   3
      Text            =   "proba"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ComboBox felhasznalok 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jelszó:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Felhasználó:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   885
   End
End
Attribute VB_Name = "belepes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bezar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    If jelszo.Text = Ertek("felhasznalok", "id", felhasznalok.ItemData(felhasznalok.ListIndex), "jelszo") Then
        Fid = felhasznalok.ItemData(felhasznalok.ListIndex)
        Unload Me
    Else
        MsgBox "Hibás jelszó", vbCritical
    End If
End Sub

Private Sub Form_Initialize()
    Load bontoware
    felso = bontoware.piros
    Betolt Me.felhasznalok, "felhasznalok", "nev", "nev"
    If Fid <> 0 Then Jelol Me.felhasznalok, CLng(Fid)
End Sub

Private Sub Form_Load()
    Form_Initialize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Fid = 0 Then
        End
    Else
        Unload Me
    End If
End Sub
