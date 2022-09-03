VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form atvetel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Közlekedésfelügyeleti átvétel"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   8265
      TabIndex        =   5
      Top             =   0
      Width           =   8295
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Jelölje ki azon gépjármûveket, melyeket átad a közlekedésfelügyelet képviselõjének:"
         Height          =   375
         Left            =   4680
         TabIndex        =   7
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KFF átvétel"
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
         Left            =   5400
         TabIndex        =   6
         Top             =   0
         Width           =   1560
      End
   End
   Begin VB.CommandButton frissites 
      Caption         =   "Frissít"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   1575
   End
   Begin MSComctlLib.ListView autok 
      Height          =   4575
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Azonosító"
         Object.Width           =   573
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nyilvántartási szám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Gyártmány, típus"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton rontott 
      Caption         =   "Ooops, rontottam"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton bezar 
      Caption         =   "Bezárás"
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton ment 
      Caption         =   "Mentés"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   5640
      Width           =   1575
   End
End
Attribute VB_Name = "atvetel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim elem As ListItem

Private Sub bezar_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
    felso = bontoware.piros
    Frissit
End Sub

Public Sub Frissit()
    Dim Sor As New ADODB.Recordset, p As String
    
    autok.Visible = False
    autok.ListItems.Clear
    
    p = "SELECT kiadfelhnap.id, autok.nyszam, markak.marka, tipusok.tipus "
    p = p & "FROM (markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN (kiadfelhnap INNER JOIN autok ON kiadfelhnap.auto = autok.id) ON tipusok.id = autok.tipus "
    p = p & "WHERE (((kiadfelhnap.atveve)=False) AND ((kiadfelhnap.rontott)=False)) "
    p = p & "ORDER by autok.nyszam"
    
    SQL_p p, Sor
    
    If Not Sor.EOF Then Sor.MoveFirst
    Do While Not Sor.EOF
        Set elem = autok.ListItems.Add(, , Sor.Fields(0))
            elem.ListSubItems.Add , , Sor.Fields(1)
            elem.ListSubItems.Add , , Sor.Fields(2) & " " & Sor.Fields(3)
        Sor.MoveNext
    Loop
    Sor.Close
    
    autok.Visible = True
End Sub

Private Sub Form_Load()
    Frissit
End Sub

Private Sub frissites_Click()
    Frissit
End Sub

Private Sub ment_Click()
    Dim i As Integer, Sor As New ADODB.Recordset
    For i = 1 To autok.ListItems.Count
        If autok.ListItems(i).Checked = True Then
            'MsgBox autok.ListItems(i).Text
            SQL_p "UPDATE kiadfelhnap SET atveve=True, atv_datum='" & Date & "' WHERE id=" & autok.ListItems(i).Text, Sor
        End If
        'if autok.ListItems
    Next i
    Frissit
End Sub

Private Sub rontott_Click()
    atv_ron.Show
End Sub
