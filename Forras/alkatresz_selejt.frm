VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form alkatresz_selejt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alkatrész selejtezés"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   10305
      TabIndex        =   5
      Top             =   0
      Width           =   10335
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Jelölje ki azokat az alkatrészeket amiket meg szeretne tartani, majd kattintson a mentés gombra."
         Height          =   435
         Left            =   6000
         TabIndex        =   7
         Top             =   480
         Width           =   4170
      End
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alkatrészek selejtezése autóból"
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
         Left            =   4680
         TabIndex        =   6
         Top             =   0
         Width           =   4365
      End
   End
   Begin VB.CommandButton mentes 
      Caption         =   "Mentés"
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin MSComctlLib.TreeView fa 
      Height          =   5295
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9340
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.CommandButton frissites 
      Caption         =   "Frissít"
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label auto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NYszam:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Label nyszam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NYszam:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1245
   End
End
Attribute VB_Name = "alkatresz_selejt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim elem As ListItem
Dim SID As Long

Public Sub Mutasd(Kit As Long)
    Dim Sor As New ADODB.Recordset
    'Form_Initialize
    SID = Kit
    
    SQL_p "SELECT autok.nyszam, markak.marka, tipusok.tipus, autok.evjarat " & _
        "FROM (markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN autok ON tipusok.id = autok.tipus " & _
        "WHERE autok.id=" & SID, Sor
    
    nyszam.Caption = Nstr(Sor.Fields(0).Value)
    auto.Caption = Nstr(Sor.Fields(1).Value & " " & Sor.Fields(2).Value & " [" & Sor.Fields(3).Value & "]")
    Sor.Close
    Frissit
    Me.Show
End Sub
Private Sub Frissit()
    Dim Sor As New ADODB.Recordset
    Dim Akt As Node
    fa.Nodes.Clear
    
    SQL_p "SELECT * FROM focsop", Sor
    Sor.MoveFirst
    Do While Not Sor.EOF
        fa.Nodes.Add , , Nstr("f" & Sor!Id), Nstr(Sor!nev)
        Sor.MoveNext
    Loop
    Sor.Close
    
    SQL_p "SELECT * FROM alcsop", Sor
    Sor.MoveFirst
    Do While Not Sor.EOF
        fa.Nodes.Add "f" & Sor!focsop, tvwChild, Nstr("a" & Sor!Id), Nstr(Sor!nev)
        Sor.MoveNext
    Loop
    Sor.Close
    
    '                   0                   1                   2                       3                       4               5                   6                       7                   8                   9
    SQL_p "SELECT focsop.cikkszam, alcsop.cikkszam, alkatresznevek.cikkszam, alkatresznevek.alcsop, alkatresznevek.nev, raktarkeszlet.tipus, raktarkeszlet.elkelt, raktarkeszlet.sztorno, raktarkeszlet.selejt, raktarkeszlet.auto, raktarkeszlet.id " & _
        "FROM focsop INNER JOIN (alcsop INNER JOIN (alkatresznevek INNER JOIN raktarkeszlet ON alkatresznevek.id = raktarkeszlet.alkatresz) ON alcsop.id = alkatresznevek.alcsop) ON focsop.id = alcsop.focsop " & _
        "WHERE (((raktarkeszlet.tipus)=0) AND ((raktarkeszlet.elkelt)=False) AND ((raktarkeszlet.sztorno)=False) AND ((raktarkeszlet.auto)=" & SID & "))", Sor
    
    If Not Sor.EOF Then Sor.MoveFirst
    Do While Not Sor.EOF
        Set Akt = fa.Nodes.Add("a" & Sor.Fields(3).Value, tvwChild, Nstr("r" & Nstr(Sor.Fields(10).Value)), Nstr(Sor.Fields(4).Value))
        Akt.Checked = Sor.Fields(8).Value
        Sor.MoveNext
    Loop
    Sor.Close
End Sub

Private Sub fa_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    For i = 1 To fa.Nodes.Count
        If Nstr(fa.Nodes(i).Parent) = Node.Text Then
            fa.Nodes(i).Checked = Node.Checked
            If fa.Nodes(i).Children > 0 Then
                fa_NodeCheck fa.Nodes(i)
            End If
        End If
    Next i
End Sub

Private Sub Form_Initialize()
    felso = bontoware.piros
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fa.Width = Me.ScaleWidth - 2 * fa.Left
    fa.Height = Me.ScaleHeight - fa.Top - fa.Left
End Sub

Private Sub frissites_Click()
    Frissit
End Sub

Private Sub mentes_Click()
Dim i As Long
    For i = 1 To fa.Nodes.Count
        If Mid(fa.Nodes(i).Key, 1, 1) = "r" Then
            FSQL "UPDATE raktarkeszlet SET selejt=" & Alakit(fa.Nodes(i).Checked, "FALSE", "TRUE") & " WHERE id=" & Mid(fa.Nodes(i).Key, 2)
        End If
    Next i
    If vbYes = MsgBox("Az adatok elmentve! Be kívánja zárni ezt az abakot?", vbYesNo + vbQuestion, "Mentés befejezõdöt") Then Unload Me
End Sub
