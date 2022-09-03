VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form szamla_lista 
   Caption         =   "Számlák listája"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame szurok 
      Caption         =   "Szûrés"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   9975
      Begin VB.CheckBox isdatum 
         Caption         =   "Megadott dátumtól dátumig"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox fizmod 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox partner 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker ig 
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   45219841
         CurrentDate     =   38561
      End
      Begin MSComCtl2.DTPicker tol 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   45219841
         CurrentDate     =   38561
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fizetési mód:"
         Height          =   195
         Index           =   1
         Left            =   6120
         TabIndex        =   7
         Top             =   240
         Width           =   915
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adott partner számlái:"
         Height          =   195
         Index           =   0
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.ComboBox szszam 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin MSComctlLib.ListView szamlak 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   13361
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Azonosító"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nyilvántartási szám"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Típusa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Vevõ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Szállító"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fizetési mód"
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Számla kelte"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Teljesítés"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Határidõ"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Nyomtatott példányok száma"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Súly"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Érték"
         Object.Width           =   2187
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Számlaszám"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "szamla_lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jelez As Boolean

Private Sub Form_Initialize()
    Me.meghiv
End Sub

Private Sub Frissit()
    Dim elem As ListItem
    Dim p As String, p1 As String, p2 As String, p3 As String, p4 As String, p5 As String, Sor As New ADODB.Recordset
    Dim szin As ColorConstants
    
    p1 = "1=1"
    p2 = "1=1"
    p3 = "1=1"
    p4 = "1=1"
    p5 = "1=1"
    
    szamlak.ListItems.Clear
    szamlak.Visible = False
    
    If szszam.ListIndex <> 0 Then p1 = "(szamla.id)=" & szszam.ItemData(szszam.ListIndex)
    If partner.ListIndex <> 0 Then p2 = "(partnerek.id)=" & partner.ItemData(partner.ListIndex)
    If fizmod.ListIndex <> 0 Then p3 = "(szamla.fizmod)=" & fizmod.ItemData(fizmod.ListIndex)
    If isdatum.Value = 1 Then p4 = "(szamla.kelt)>=" & DatumAtir(tol.Value)
    If isdatum.Value = 1 Then p5 = "(szamla.kelt)<=" & DatumAtir(ig.Value)
    
    p = "SELECT szamla.id, szamla.szam, partnerek.vnev, partnerek.knev, szamla.fizmod, szamla.kelt, szamla.ido, szamla.teljesites, szamla.hatarido, szamla.peldany, Sum(raktarkeszlet.suly) AS SumOfsuly, Sum(Abs([raktarkeszlet]![ar]*(([raktarkeszlet]![afa]/100)+1))) AS Kif1, szamla.tipus "
    p = p & "FROM partnerek INNER JOIN ((elkelt INNER JOIN raktarkeszlet ON elkelt.tetel = raktarkeszlet.id) INNER JOIN szamla ON elkelt.szamla = szamla.id) ON partnerek.id = szamla.vevo "
    p = p & "WHERE (" & p1 & ") AND (" & p2 & ") AND (" & p3 & ") AND (" & p4 & ") AND (" & p5 & ") "
    p = p & "GROUP BY szamla.id, szamla.szam, partnerek.vnev, partnerek.knev, szamla.fizmod, szamla.kelt, szamla.ido, szamla.teljesites, szamla.hatarido, szamla.peldany, szamla.tipus "
    p = p & "ORDER BY szamla.szam"
    
    SQL_p p, Sor
    If Not Sor.EOF Then Sor.MoveFirst
    Do While Not Sor.EOF
        Set elem = szamlak.ListItems.Add(, , Sor.Fields(0).Value)                       'Azonosito
            
            Select Case Sor.Fields(12).Value
                Case 0
                    p = "Alkatrész számla"
                    p1 = 0
                    szin = vbBlack
                Case 1
                    p = "Szornózott alkatrész számla"
                    p1 = 0
                    szin = vbRed
                Case 2
                    p = "Hulladék számla"
                    p1 = 1
                    szin = vbBlack
                Case 3
                    p = "Szornózott hulladék számla"
                    p1 = 1
                    szin = vbRed
            End Select
            
            elem.ListSubItems.Add , , p1                                             'Tipus. 0=alkatreszes, 1= hulladekos
            elem.ListSubItems.Add , , Sor.Fields(1).Value & "/" & Year(Sor.Fields(5).Value)                              'Szamlaszam
            elem.ListSubItems.Add , , p  'Típusa
            elem.ListSubItems.Add , , Sor.Fields(2).Value & " " & Sor.Fields(3).Value   'Vevo neve
            elem.ListSubItems.Add , , "-"                                            'Szállító neve - itt nincs
            
            
            If Sor.Fields(4) = 0 Then                                                   'Fizetés módja
                elem.ListSubItems.Add , , "Kézpénz"
            Else
                elem.ListSubItems.Add , , "Átutalás"
            End If
            
            elem.ListSubItems.Add , , Sor.Fields(5).Value & " " & Sor.Fields(6).Value   'Számla kelte
            elem.ListSubItems.Add , , Sor.Fields(7).Value                               'Teljesites
            elem.ListSubItems.Add , , Sor.Fields(8).Value                               'Hatarido
            elem.ListSubItems.Add , , Sor.Fields(9).Value                               'Nyomtatott peldanyok szama
            elem.ListSubItems.Add , , Sor.Fields(10).Value                              'Osszsuly
            elem.ListSubItems.Add , , Sor.Fields(11).Value                              'Ar
        
            RowColor szin, elem
            Sor.MoveNext
    Loop
    Sor.Close
    
    szamlak.Visible = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
    szamlak.Move szamlak.Left, szamlak.Top, Me.ScaleWidth - 2 * szamlak.Left, Me.ScaleHeight - szamlak.Left - szamlak.Top
    szurok.Width = Me.ScaleWidth - 2 * szurok.Left
End Sub

Private Sub isdatum_Click()
    Frissit
    tol.Enabled = isdatum.Value
    ig.Enabled = tol.Enabled
End Sub

Private Sub Label3_Click()

End Sub

Private Sub szamlak_DblClick()
    If szamlak.SelectedItem.ListSubItems(1) = 0 Then
        szamlazo.modosit szamlak.SelectedItem.Text, 90
    Else
        hulladek_szamla.modosit szamlak.SelectedItem.Text, 90
    End If
End Sub

Private Sub szszam_Change()
    If jelez Then Frissit
End Sub

Private Sub szszam_Click()
    szszam_Change
End Sub

Private Sub partner_Change()
    If jelez Then Frissit
End Sub

Private Sub partner_Click()
    partner_Change
End Sub

Private Sub fizmod_Change()
    If jelez Then Frissit
End Sub

Private Sub fizmod_Click()
    fizmod_Change
End Sub

Private Sub tol_Change()
    If (jelez) And (isdatum.Value = 1) Then Frissit
End Sub

Private Sub ig_Change()
    If (jelez) And (isdatum.Value = 1) Then Frissit
End Sub

Private Sub partner_betolt()
    Partner_Listaba Me.partner, "Minden partnertõl"
End Sub

Private Sub szszam_betolt()
    Dim Sor As New ADODB.Recordset, i As Integer
    szszam.Clear
    szszam.List(0) = "Számlaszám"
    szszam.ItemData(0) = 0
    SQL_p "SELECT szamla.id, szamla.szam FROM szamla", Sor
    If Not Sor.EOF Then Sor.MoveFirst
    i = 1
    Do While Not Sor.EOF
       szszam.List(i) = Sor!Szam
       szszam.ItemData(i) = Sor!Id
       i = i + 1
       Sor.MoveNext
    Loop
    ElsotJelol Me.szszam
    Sor.Close
End Sub

Private Sub fizmod_betolt()
    fizmod.Clear
    fizmod.List(0) = "Mindegy"
    fizmod.ItemData(0) = 3
    fizmod.List(1) = "Kézpénz"
    fizmod.ItemData(1) = 0
    fizmod.List(2) = "Átutalás"
    fizmod.ItemData(2) = 1
    ElsotJelol Me.fizmod
End Sub

Public Sub meghiv()
    jelez = False
    szszam_betolt
    partner_betolt
    fizmod_betolt
    tol.Value = Date
    ig.Value = Date
    jelez = True
    Frissit
    Me.Show
End Sub
