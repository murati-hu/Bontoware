VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form partner_lista 
   Caption         =   "Nyilvántartott partnerek"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Új partner"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Módosít"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox keres 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   5775
   End
   Begin MSComctlLib.ListView partnerek 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11456
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Azonosító"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Név"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cím"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Személyi"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Adószám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tel:"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "E-mail"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "KÜJ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "KTJ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Állampolgárság"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Megj"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Töröl"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gyorskeresés:"
      Height          =   195
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1005
   End
End
Attribute VB_Name = "partner_lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim elem As ListItem
Public Szures As String

Public Sub Frissit(Optional Hogy As String)
'On Error GoTo hiba
Dim Id As Long, p As String
    partnerek.ListItems.Clear
    Rekord.CursorLocation = adUseClient
    sql_parancs ("SELECT * FROM partnerek " & Hogy)
    Szures = Hogy
    If Not Rekord.EOF Then Rekord.MoveFirst
    partnerek.Visible = False
    Do While Not Rekord.EOF
        Set elem = partnerek.ListItems.Add(, , Rekord!Id)
            elem.ListSubItems.Add , , Nstr(Rekord!vnev & " " & Rekord!knev)
            elem.ListSubItems.Add , , Nstr(Rekord!irszam & " " & Rekord!varos & " " & Rekord!cim)
            elem.ListSubItems.Add , , Nstr(Rekord!szemelyi)
            elem.ListSubItems.Add , , Nstr(Rekord!ado)
            elem.ListSubItems.Add , , Nstr(Rekord!tel)
            elem.ListSubItems.Add , , Nstr(Rekord!fax)
            elem.ListSubItems.Add , , Nstr(Rekord!email)
            elem.ListSubItems.Add , , Nstr(Rekord!kuj)
            elem.ListSubItems.Add , , Nstr(Rekord!ktj)
            elem.ListSubItems.Add , , Nstr(Rekord!allampolg)
            elem.ListSubItems.Add , , Nstr(Rekord!megj)
            
            Rekord.MoveNext
    Loop
    partnerek.Visible = True
    Rekord.Close
Exit Sub
Hiba:
    Hiba Err.Number, "Frissitési hiba"
End Sub

Private Sub Command1_Click()
    partner_lap.modosit partnerek.SelectedItem.Text, 10
End Sub

Private Sub Command2_Click()
    Partner_Torol partnerek.SelectedItem.Text, 10
    Frissit
End Sub

Private Sub Command3_Click()
    partner_lap.uj 10
End Sub

Private Sub Form_Load()
    Frissit
End Sub

Private Sub Form_Resize()
On Error Resume Next
    partnerek.Width = Me.ScaleWidth - partnerek.Left * 2
    partnerek.Height = Me.ScaleHeight - partnerek.Top - partnerek.Left
    
    keres.Width = Me.ScaleWidth - keres.Left - partnerek.Left
End Sub

Private Sub keres_Change()
        If keres.Text <> "" Then
            Frissit ("where vnev LIKE '%" & keres.Text & "%' or  knev LIKE '%" & keres.Text & "%'  or  cim LIKE '%" & keres.Text & "%'  or  irszam LIKE '%" & keres.Text & "%'  or  varos LIKE '%" & keres.Text & "%' or  ado LIKE '%" & keres.Text & "%' or  tel LIKE '%" & keres.Text & "%' or  fax LIKE '%" & keres.Text & "%' or  email LIKE '%" & keres.Text & "%' or  szemelyi LIKE '%" & keres.Text & "%' or  kuj LIKE '%" & keres.Text & "%' or  ktj LIKE '%" & keres.Text & "%' or  megj LIKE '%" & keres.Text & "%'")
        Else
            Frissit
        End If
End Sub

Private Sub partnerek_DblClick()
    Command1_Click
End Sub
