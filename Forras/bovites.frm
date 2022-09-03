VERSION 5.00
Begin VB.Form bovites 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nyilvántartott Autómárkák és Típusok"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5055
      Begin VB.ComboBox alkatresz 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Bezárás"
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox alcsop 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   3855
      End
      Begin VB.ComboBox focsop 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox uj_alk 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   1080
         Width           =   3855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Hozzáadás"
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fõcsoport:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Felvétele:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alkatrészek:"
         Height          =   195
         Left            =   720
         TabIndex        =   6
         Top             =   1680
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label tip_fel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alcsoportok:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   885
      End
   End
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   5265
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Label focim 
         Alignment       =   1  'Right Justify
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
         Height          =   690
         Left            =   1680
         TabIndex        =   1
         Top             =   120
         Width           =   3495
      End
   End
End
Attribute VB_Name = "bovites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alcsop_Change()
    alcsop_Click
End Sub

Private Sub alcsop_Click()
'On Error Resume Next
    Betolt Me.alkatresz, "alkatresznevek", "nev", "nev", , " WHERE alcsop=" & alcsop.ItemData(alcsop.ListIndex)
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If alcsop.ListIndex >= 0 Then
        Dim p As String
        
        p = "INSERT INTO alkatresznevek (nev, alcsop, cikkszam, vtsz, ewc, tobb) VALUES ('" & uj_alk.Text & "', " & alcsop.ItemData(alcsop.ListIndex) & ", " & alkatresz.ListCount + 1 & ", 1, 1, FALSE)"
        sql_parancs (p)
        uj_alk.Text = ""
    End If
    
    Frissit
End Sub

Private Sub focsop_Change()
    focsop_Click
End Sub

Private Sub focsop_Click()
On Error Resume Next
    Betolt Me.alcsop, "alcsop", "nev", "cikkszam", , " WHERE focsop=" & focsop.ItemData(focsop.ListIndex)
End Sub

Private Sub Form_Initialize()
    Frissit
    felso.Picture = bontoware.zold
End Sub

Private Sub Form_Load()
    Form_Initialize
End Sub
Public Sub Frissit()
    Betolt Me.focsop, "focsop", "nev", "cikkszam"
End Sub

