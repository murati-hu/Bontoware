VERSION 5.00
Begin VB.Form markak 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nyilvántartott Autómárkák és Típusok"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   8265
      TabIndex        =   12
      Top             =   0
      Width           =   8295
      Begin VB.Label focim 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gyártmányok és típusok karbantartartása"
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
         Left            =   4680
         TabIndex        =   13
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.TextBox nev_tip 
      Height          =   315
      Left            =   4200
      TabIndex        =   9
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton felv_tip 
      Caption         =   "&Hozzáadás"
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton torol_tip 
      Caption         =   "&Eltávolítás"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ListBox tipusok 
      Height          =   2790
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton torol 
      Caption         =   "&Eltávolítás"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton felvesz 
      Caption         =   "&Hozzáadás"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox nev 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.ListBox markak 
      Height          =   2790
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label tip_fel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Felvett:"
      Height          =   195
      Left            =   4200
      TabIndex        =   11
      Top             =   1680
      Width           =   525
   End
   Begin VB.Label tip_uj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Új típus:"
      Height          =   195
      Left            =   4200
      TabIndex        =   10
      Top             =   960
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Már felvett autómárkák:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Új márka:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   675
   End
End
Attribute VB_Name = "markak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Kinek As Byte
Private Sub felv_tip_Click()
    Dim p As String
    p = "INSERT INTO tipusok (marka, tipus) VALUES (" & markak.ItemData(Valasztott) & ",'" & nev_tip.Text & "')"
    sql_parancs (p)
    nev_tip.Text = ""
    Frissites_tip
End Sub

Private Sub felvesz_Click()
    Dim p As String
    p = "INSERT INTO markak (marka) VALUES ('" & nev.Text & "')"
    sql_parancs (p)
    nev.Text = ""
    frissites
End Sub

Private Sub Form_Load()
    felso = bontoware.narancs
    frissites
    markak.Selected(0) = True
End Sub
Public Sub Frissites_tip()
'On Error GoTo hiba
Dim Id As Long
    tipusok.Clear
    Rekord.CursorLocation = adUseClient
    Call sql_parancs("SELECT * FROM tipusok where marka=" & markak.ItemData(Valasztott()))
    If Not Rekord.EOF Then Rekord.MoveFirst
    Id = 0
    Do While Not Rekord.EOF
        tipusok.List(Id) = Rekord!tipus
        tipusok.ItemData(Id) = Rekord!Id
        Rekord.MoveNext
        Id = Id + 1
    Loop
    Rekord.Close
Exit Sub
Hiba:
    Hiba Err.Number, "Frissitési hiba"
End Sub
Public Sub frissites()
On Error GoTo Hiba
Dim Id As Long
    markak.Clear
    Rekord.CursorLocation = adUseClient
    Call sql_parancs("SELECT * from markak order by marka")
    If Not Rekord.EOF Then Rekord.MoveFirst
    Id = 0
    Do While Not Rekord.EOF
        markak.List(Id) = Rekord!marka
        markak.ItemData(Id) = Rekord!Id
        Rekord.MoveNext
        Id = Id + 1
    Loop
    Rekord.Close
Exit Sub
Hiba:
    Hiba Err.Number, "Frissitési hiba"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visszajelez Kinek, 0
End Sub

Private Sub markak_Click()
    tip_uj.Caption = "Új " & markak.List(Valasztott) & " típus:"
    tip_fel.Caption = "Már felvett " & markak.List(Valasztott) & " típusok:"
    Frissites_tip
End Sub

Private Sub ok_Click()
    Unload Me
End Sub

Private Sub torol_Click()
If markak.ListIndex < 0 Then Exit Sub
    Dim p As String
    p = "DELETE FROM markak WHERE id=" & markak.ItemData(Valasztott())
    'MsgBox p
    sql_parancs (p)
    frissites
    nev.Text = ""
End Sub
Private Function Valasztott() As Long
    Dim i As Long
    i = 0
    Do While Not markak.Selected(i)
        i = i + 1
    Loop
    Valasztott = i
End Function
Private Function Valasztott_tip() As Long
    Dim i As Long
    i = 0
    Do While Not tipusok.Selected(i)
        i = i + 1
    Loop
    Valasztott_tip = i
End Function
Private Sub torol_tip_Click()
If tipusok.ListIndex < 0 Then Exit Sub
    Dim p As String
    p = "DELETE FROM tipusok WHERE id=" & tipusok.ItemData(Valasztott_tip())
    'MsgBox p
    sql_parancs (p)
    Frissites_tip
    nev.Text = ""
End Sub
Public Sub uj(Optional Hova As Byte)
    Me.Show
    Kinek = Hova
End Sub
