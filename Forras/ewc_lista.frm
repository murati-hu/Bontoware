VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ewc_lista 
   Caption         =   "Bontási napló tételei"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   7065
      TabIndex        =   1
      Top             =   0
      Width           =   7095
      Begin VB.Label cimke 
         BackStyle       =   0  'Transparent
         Caption         =   "Jelöljeki annak a hulladéknak a négyzetét, amelyet kibont a szárazrafektetésnél"
         Height          =   735
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Szárazrafektetési napló tételei"
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
         Left            =   2760
         TabIndex        =   2
         Top             =   0
         Width           =   4155
      End
   End
   Begin MSComctlLib.ListView ewc_lista 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7011
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "EWC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Megnevezés"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Mértékegység"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Szorzó"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "ewc_lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim elem As ListItem

Private Sub ewc_lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
     sql_parancs "UPDATE ewc SET bontas=" & Alakit(Item.Checked, "1", "0") & " where ewc='" & Item.Text & "'"
End Sub

Public Sub Frissit()
'On Error GoTo hiba
Dim Id As Long, p As String
    ewc_lista.ListItems.Clear
    Rekord.CursorLocation = adUseClient
    sql_parancs ("SELECT * FROM ewc ORDER BY ewc")
    If Not Rekord.EOF Then Rekord.MoveFirst
    ewc_lista.Visible = False
    Do While Not Rekord.EOF
        Set elem = ewc_lista.ListItems.Add(, , Rekord!ewc & Alakit(Rekord!veszelyes, "*", ""))
            elem.ListSubItems.Add , , Nstr(Rekord!nev)
            elem.ListSubItems.Add , , Nstr(Rekord!Me)
            elem.ListSubItems.Add , , Nstr(Rekord!szorzo)
            elem.Checked = Nstr(Rekord!bontas)
            Rekord.MoveNext
    Loop
    ewc_lista.Visible = True
    Rekord.Close
Exit Sub
Hiba:
    Hiba Err.Number, "Frissitési hiba"
End Sub

Private Sub Form_Load()
    felso = bontoware.narancs
    Frissit
End Sub
