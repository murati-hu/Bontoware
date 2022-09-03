VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form kiad_felh_nap 
   Caption         =   "Kiad�si �s felhaszn�l�si napl�"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton nyomtatas 
      Caption         =   "Nyomtat�s"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.ComboBox mennyi 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton akarok 
      Caption         =   "Nyomtatni akarok"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Azonos�t�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bont�si �tv�teli igazol�s sorsz�ma"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bont�si �tv�teli igazol�s �tv�tel�nek d�tuma"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�llamjelz�s"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Rendsz�m"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Azonos�t�si (alv�z) sz�ma"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Forgalmi enged�ly sz�ma"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "T�rzsk�nyv sz�ma"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "�tv�ve"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "�tv�tel d�tuma"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "�tadni k�v�nt g�pj�rm�vek sz�ma:"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "kiad_felh_nap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim elem As ListItem
Dim obj As PageSet.PrinterControl
'Nyomtat�si marg�k
Dim mb As String, mt As String, ml As String, mr As String
Dim fej As String, lab As String


Private Sub akarok_Click()          '########## Nyomtat�si sz�nd�kot jelzi, innen sz�molja,hogy h�ny kocsit ###########
    Dim Sor As New ADODB.Recordset
    Dim i As Integer
    On Error GoTo errorhandler
        SQL_p "SELECT * FROM kiadfelhnap WHERE atveve=False order by bon_szam", Sor
        'MsgBox Sor.RecordCount \ 10
        
        If Sor.RecordCount < 10 Then
            mennyi.List(0) = 0
            mennyi.ItemData(0) = 0
        Else
            For i = 0 To Sor.RecordCount \ 10
                mennyi.List(i) = CStr(10 * (i + 1))
                mennyi.ItemData(i) = 10 * (i + 1)
            Next i
        End If
        mennyi.ListIndex = 0
        mennyi.Visible = True
        nyomtatas.Visible = True
        Label1.Visible = True
        'Set DataReport1.DataSource = Sor
        'DataReport1.Show
    Exit Sub

errorhandler:
       MsgBox Err.Description
       obj.ReSetOrientation
End Sub

Private Sub nyomtatas_Click()       '####### Maga a nyomtat�s, megh�vja a datareportot
    Dim Sor As New ADODB.Recordset
    If mennyi.ItemData(mennyi.ListIndex) <> 0 Then
        SQL_p "SELECT TOP " & mennyi.ItemData(mennyi.ListIndex) & " * FROM kiadfelhnap WHERE atveve=False order by bon_szam", Sor
        Set kiadfelhnap.DataSource = Sor
        kiadfelhnap.Show
    Else
        MsgBox "Legal�bb 10 aut� kell ahhoz hogy �t tudjuk adni"
    End If
End Sub


Private Sub Form_Load()
    'On Error Resume Next
    Dim Id As Long, p As String
    Dim Sor As New ADODB.Recordset
    
    Set obj = New PrinterControl
    mennyi.Visible = False
    nyomtatas.Visible = False
    Label1.Visible = False
    
    
    
    'Marg�be�ll�t�sok beolvas�sa a k�s�bbi visszat�lt�shez
    fej = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "header")
    lab = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "footer")
    mb = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_bottom")
    mt = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_top")
    ml = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_left")
    mr = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_right")
    
    'Marg�k t�rl�se �s fejl�cek t�rl�se
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_bottom", "0.28504"
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_left", "0.39370"
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_right", "0.47244"
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_top", "0.16500"
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "footer", ""
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "header", ""
    
    obj.ChngOrientationLandscape
    'kiadfelhnap.
    lista.Visible = False
    
    p = "SELECT * FROM kiadfelhnap WHERE atveve=False order by bon_szam"
    
    SQL_p p, Sor
    If Not Sor.EOF Then Sor.MoveFirst
    Do While Not Sor.EOF
        Set elem = lista.ListItems.Add(, , Sor.Fields(0))
            If Sor.Fields(10) = False Then
                elem.ListSubItems.Add , , Sor.Fields(1)
                elem.ListSubItems.Add , , Sor.Fields(2)
                elem.ListSubItems.Add , , Sor.Fields(3)
                elem.ListSubItems.Add , , Sor.Fields(4)
                elem.ListSubItems.Add , , Sor.Fields(5)
                elem.ListSubItems.Add , , Sor.Fields(6)
                elem.ListSubItems.Add , , Nstr(Sor.Fields(7))
                'elem.ListSubItems.Add , , Nstr(Sor.Fields(8))
                elem.ListSubItems.Add , , Alakit(Nstr(Sor.Fields(8).Value), "�tv�ve", "Nincs �tv�ve")
                elem.ListSubItems.Add , , Alakit(Nstr(Sor.Fields(8).Value), Nstr(Sor.Fields(9)), "")
            Else            'Ha rontva van, akkor nem t�lt�m ki, csak ki�rom, hogy rontott
                elem.ListSubItems.Add , , Sor.Fields(1)
                elem.ListSubItems.Add , , "RONTOTT"
            End If
        Sor.MoveNext
    Loop
    Sor.Close
    lista.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
        On Error Resume Next
        'obj.ReSetOrientation 'This resets the printer to portrait.
        obj.ChngOrientationPortrait
      
        'Marg��rt�kek vissza�r�sa
        SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_bottom", mb
        SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_left", ml
        SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_right", mr
        SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_top", mt
        SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "footer", fej
        SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "header", lab

   End Sub
