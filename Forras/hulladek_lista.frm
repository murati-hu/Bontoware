VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form hulladek_lista 
   Caption         =   "Hulladék lista"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Frissit"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.ListView hulladekok 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   12515
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Azonosító"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "EWC Kód"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "EWC Megnevezés"
         Object.Width           =   4763
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Suly"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label ossztomeg 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   480
   End
   Begin VB.Label ossz 
      Caption         =   "Összes tömeg"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "hulladek_lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim elem As ListItem
Public Szures As String

Public Sub Frissit()
    'On Error GoTo hiba
    Dim Id As Long, p As String, osszsuly As Double
    Dim Sor As New ADODB.Recordset
    
    osszsuly = 0
    ossztomeg.Visible = False
    hulladekok.ListItems.Clear
    
    'sor.CursorLocation = adUseClient
    
     '              0           1           2       3           4               5
     'p = "SELECT ewc.id, ewc.veszelyes, ewc.ewc, ewc.nev, autok.nyszam, Sum(([raktarkeszlet]![irany]*[raktarkeszlet]![suly])) AS SumofSuly " & _
     '       "FROM autok INNER JOIN (ewc INNER JOIN raktarkeszlet ON ewc.id = raktarkeszlet.ewc) ON autok.id = raktarkeszlet.auto " & _
     '       "GROUP BY ewc.id, ewc.veszelyes, ewc.ewc, ewc.nev, autok.nyszam "
            
    'p = "SELECT ewc.id, ewc.veszelyes, ewc.ewc, ewc.nev, Sum(([raktarkeszlet]![irany]*[raktarkeszlet]![suly])) AS SumofSuly, raktarkeszlet.auto, raktarkeszlet.sztorno " & _
    '    "FROM ewc INNER JOIN raktarkeszlet ON ewc.id = raktarkeszlet.ewc " & _
    '    "GROUP BY ewc.id, ewc.veszelyes, ewc.ewc, ewc.nev, raktarkeszlet.auto, raktarkeszlet.elkelt, raktarkeszlet.sztorno " & _
    '    "HAVING (((raktarkeszlet.elkelt)=False) AND ((raktarkeszlet.sztorno)=False)) "
    
    p = "SELECT ewc.id, ewc.veszelyes, ewc.ewc, ewc.nev, Sum(([raktarkeszlet]![irany]*[raktarkeszlet]![suly])) AS SumofSuly "
    p = p & "FROM ewc INNER JOIN raktarkeszlet ON ewc.id = raktarkeszlet.ewc "
    p = p & "GROUP BY ewc.id, ewc.veszelyes, ewc.ewc, ewc.nev, raktarkeszlet.sztorno, raktarkeszlet.elkelt "
    'p = p & "HAVING (((raktarkeszlet.sztorno)=False) AND ((raktarkeszlet.elkelt)=False) AND ((ewc.id)<>7))"
    p = p & "HAVING (((raktarkeszlet.sztorno)=False) AND ((raktarkeszlet.elkelt)=False))"

    SQL_p p, Sor
    
    If Not Sor.EOF Then Sor.MoveFirst
    hulladekok.Visible = False
    Do While Not Sor.EOF
        If Sor.Fields(0).Value <> 7 Then
        Set elem = hulladekok.ListItems.Add(, , Sor.Fields(0).Value)
            elem.ListSubItems.Add , , Nstr(Sor.Fields(2).Value) & Alakit(Sor.Fields(1).Value, "*", "")
            elem.ListSubItems.Add , , Nstr(Sor.Fields(3).Value)
            elem.ListSubItems.Add , , Nstr(Sor.Fields(4).Value)
            
            If Sor.Fields(1).Value Then
                RowColor vbRed, elem
            Else
                'RowColor vbGreen, elem
            End If
            
            osszsuly = osszsuly + (Sor.Fields(4).Value)
        End If
            Sor.MoveNext
    Loop
    hulladekok.Visible = True
    Sor.Close
    ossztomeg.Caption = osszsuly & " kg"
    ossztomeg.Visible = True
Exit Sub
Hiba:
    Hiba Err.Number, "Frissitési hiba"
End Sub

Private Sub Command1_Click()
    Frissit
End Sub

Private Sub Form_Initialize()
    'Dim sor As New ADODB.Recordset, i As Long
    'adatmotor.Megnyitas
    'Sor.CursorLocation = adUseClient
    'SQL_p "SELECT * FROM ewc order by ewc", sor
    'If Not sor.EOF Then sor.MoveFirst
    'valasztas.List(0) = "Mind"
    'valasztas.ItemData(0) = 0
    'i = 1
    'Do While Not sor.EOF
    '    valasztas.List(i) = sor!ewc & Alakit(sor!veszelyes, "*", "") & " - " & sor!nev
    '    valasztas.ItemData(i) = sor!Id
    '    i = i + 1
    '    sor.MoveNext
    'Loop
    'valasztas.ListIndex = 0
    'sor.Close
    Frissit
End Sub

Private Sub Form_Load()
    Form_Initialize
End Sub


Private Sub hulladekok_DblClick()
    If hulladekok.SelectedItem.Text <> -1 Then
        hulladek_lista_reszl.meghiv hulladekok.SelectedItem.Text
    End If
End Sub
