VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form auto_selejt 
   Caption         =   "Autók selejtezése"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bezar 
      Caption         =   "Bezár"
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton ment 
      Caption         =   "Mentés"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   6120
      Width           =   2055
   End
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   9105
      TabIndex        =   1
      Top             =   0
      Width           =   9135
      Begin VB.Label cimke 
         BackStyle       =   0  'Transparent
         Caption         =   "Kérem pipálja ki azoknak az autóknak a jelölõnégyzetét, amelyeket le kívánja selejtezni."
         Height          =   375
         Left            =   5760
         TabIndex        =   3
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Autók selejtezése"
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
         Left            =   6120
         TabIndex        =   2
         Top             =   0
         Width           =   2430
      End
   End
   Begin MSComctlLib.ListView auto_selejt 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8916
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Azonosító"
         Object.Width           =   618
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nyilvántartási szám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Állapot"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Átvétel napja"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Gyártmány"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Típus"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Alvázszám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Motorszám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Rendszám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Eladó"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Önsúly"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "auto_selejt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim elem As ListItem
Public Szures As String
Private Seged As New ADODB.Recordset
Private betoltve As Boolean



Private Sub auto_selejt_DblClick()
    'adatlap.Megnyit auto_selejt.SelectedItem.Text, 30
    alkatresz_selejt.Mutasd auto_selejt.SelectedItem.Text
End Sub
'Private Sub auto_selejt_ItemCheck2(ByVal Item As MSComctlLib.ListItem)
'    If MsgBox("Biztos meg akarja változtatni az autó állapotát?", vbQuestion + vbYesNo, "Selejtezés") = vbYes Then
'        SelejtezAuto Item.Text, Item.Checked
'   Else
'        Item.Checked = Not Item.Checked
'    End If
'End Sub


Public Sub Frissit(Optional Hogy As String)
'On Error GoTo hiba

If Not betoltve Then Exit Sub
If Not Me.Visible Then Exit Sub

Dim Id As Long, p As String
    auto_selejt.ListItems.Clear
    'Rekord.CursorLocation = adUseClient
    'sql_parancs ("SELECT * FROM autok " & Hogy)
    
    Dim sor As New ADODB.Recordset
    p = "SELECT autok.id, autok.nyszam, autok.magan, autok.datum, markak.marka, tipusok.tipus, autok.alvaz, autok.motor, autok.rendszam, partnerek.vnev, partnerek.knev, autok.tomeg, autok.selejt " & _
        "FROM partnerek INNER JOIN ((markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN autok ON tipusok.id = autok.tipus) ON partnerek.id = autok.elado; "

    SQL_p p, sor
    Szures = Hogy
    auto_selejt.Visible = False
    If Not sor.EOF Then sor.MoveFirst
    
    Do While Not sor.EOF
        'Seged = Rekord
        Set elem = auto_selejt.ListItems.Add(, , sor!Id)
            elem.ListSubItems.Add , , Nstr(sor!nyszam)
            elem.ListSubItems.Add , , Nstr(sor!magan)
            elem.ListSubItems.Add , , Nstr(sor!datum)
            elem.ListSubItems.Add , , Nstr(sor!marka)
            elem.ListSubItems.Add , , Nstr(sor!tipus)
            elem.ListSubItems.Add , , Nstr(sor!alvaz)
            elem.ListSubItems.Add , , Nstr(sor!motor)
            elem.ListSubItems.Add , , Nstr(sor!rendszam)
            elem.ListSubItems.Add , , Nstr(sor!vnev) & " " & Nstr(sor!knev)
            elem.ListSubItems.Add , , Nstr(sor!tomeg)
            elem.Checked = sor!selejt
            
            sor.MoveNext
    Loop
    auto_selejt.Visible = True
    sor.Close
Exit Sub
Hiba:
    Hiba Err.Number, "Frissitési hiba"
End Sub

Private Sub bezar_Click()
    Unload Me
End Sub
Private Sub Form_Paint()
    If Not betoltve And Visible Then
        betoltve = True
        Frissit
    End If
End Sub
Private Sub Form_Load()
    betoltve = False
    felso = bontoware.piros
    Frissit
End Sub

Private Sub ment_Click()
    Dim i As Integer, sor As New ADODB.Recordset
    For i = 1 To auto_selejt.ListItems.Count
            SQL_p "UPDATE autok SET selejt=" & Alakit(auto_selejt.ListItems(i).Checked, "TRUE", "FALSE") & " WHERE id=" & auto_selejt.ListItems(i).Text, sor
        'if autok.ListItems
    Next i
    Frissit
End Sub
