VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form atv_ron 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Közlekedésfelügyeleti átvétel módosítása"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   10305
      TabIndex        =   4
      Top             =   0
      Width           =   10335
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KFF átvétel módosítása"
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
         Left            =   4560
         TabIndex        =   6
         Top             =   0
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Jelölje ki azon gépjármûveket, melyeket mégsem ad át a közlekedésfelügyelet képviselõjének:"
         Height          =   495
         Left            =   4560
         TabIndex        =   5
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.CommandButton frissites 
      Caption         =   "Frissít"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton ment 
      Caption         =   "Mentés"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton bezar 
      Caption         =   "Bezárás"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   5640
      Width           =   2175
   End
   Begin MSComctlLib.ListView autok 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   8295
      _ExtentX        =   14631
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
End
Attribute VB_Name = "atv_ron"
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

Private Sub Frissit()
    Dim Sor As New ADODB.Recordset, p As String
    
    autok.Visible = False
    autok.ListItems.Clear
    
    p = "SELECT kiadfelhnap.id, autok.nyszam, markak.marka, tipusok.tipus "
    p = p & "FROM (markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN (kiadfelhnap INNER JOIN autok ON kiadfelhnap.auto = autok.id) ON tipusok.id = autok.tipus "
    p = p & "WHERE (((kiadfelhnap.atveve)=True) AND ((kiadfelhnap.rontott)=False)) "
    p = p & "ORDER by autok.nyszam"
    
    SQL_p p, Sor
    
    If Not Sor.EOF Then Sor.MoveFirst
    Do While Not Sor.EOF
        Set elem = autok.ListItems.Add(, , Sor.Fields(0))
            elem.ListSubItems.Add , , Sor.Fields(1)
            elem.ListSubItems.Add , , Sor.Fields(2) & " " & Sor.Fields(3)
            elem.Checked = True
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
        If autok.ListItems(i).Checked = False Then
            'MsgBox autok.ListItems(i).Text
            SQL_p "UPDATE kiadfelhnap SET atveve=False, atv_datum='" & Date & "' WHERE id=" & autok.ListItems(i).Text, Sor
        End If
    Next i
    Frissit
    atvetel.Frissit
End Sub

