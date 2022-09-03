VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form nyomtatasikep 
   Caption         =   "Nyomtatási kép"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer ido 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   14520
      Top             =   240
   End
   Begin VB.PictureBox gombsor 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   0
      ScaleHeight     =   7935
      ScaleWidth      =   9495
      TabIndex        =   1
      Top             =   0
      Width           =   9495
      Begin VB.CommandButton Command3 
         Caption         =   "Nyomtat"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   495
         Left            =   6120
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nyomtatási kép"
         Height          =   435
         Left            =   7440
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label var 
         Alignment       =   2  'Center
         Caption         =   "Nyomtatási kép készítése..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   6
         Top             =   4320
         Width           =   7335
      End
   End
   Begin SHDocVwCtl.WebBrowser bongeszo 
      Height          =   6975
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9255
      ExtentX         =   16325
      ExtentY         =   12303
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComDlg.CommonDialog parbeszed 
      Left            =   14520
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Nyomtatás elõkészítése..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   4080
      Width           =   3855
   End
End
Attribute VB_Name = "nyomtatasikep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim all As Byte
Dim Fajl As String '0 mutat, 1 kitölt
Dim Sablon As String

'Nyomtatási margók
Dim mb As String, mt As String, ml As String, mr As String
Dim fej As String, lab As String

Private Sub Command1_Click()
    '***Prevusialisation***
    'GetSettingByte HKEY_CURRENT_USER, "CompanyMyApp", "MyBinaryData"
    
    bongeszo.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER

End Sub

Private Sub Command2_Click()
On Error GoTo Hiba
    parbeszed.ShowPrinter
    'MsgBox parbeszed
Hiba:
End Sub

Private Sub Command3_Click()
    'bongeszo.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
    bongeszo.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub Form_Initialize()
    Fajl = ""
    all = 0
    Sablon = ""
    'Margóbeállítások beolvasása
    fej = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "header")
    lab = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "footer")
    mb = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_bottom")
    mt = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_top")
    ml = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_left")
    mr = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_right")
    
    'Margók törlése és fejlécek törlése
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_bottom", "0.28504"
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_left", "0.39370"
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_right", "0.47244"
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_top", "0.16500"
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "footer", ""
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "header", ""
End Sub

'Nyomtatási kép megjelenítése
Public Sub Mutasd(Mit As String)
    Form_Initialize
    Form_Resize
    Fajl = Mit
    bongeszo.Navigate2 Mit
    Me.Show vbModal
    'Command1_Click
End Sub
Private Sub Form_Resize()
On Error Resume Next

    Me.Width = 11325
    
    gombsor.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    var.Move 0, (Me.ScaleHeight - var.Height) / 2, Me.ScaleWidth
    bongeszo.Move 100, 500, Me.ScaleWidth - 2 * (bongeszo.Left), Me.ScaleHeight - 2 * (bongeszo.Left) - bongeszo.Top
    'kep.Move bongeszo.Left, bongeszo.Top, bongeszo.Width, bongeszo.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Tempfájl törlése
    If all = 1 Then Kill Fajl
    
    'Margóértékek visszaírása
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_bottom", mb
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_left", ml
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_right", mr
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_top", mt
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "footer", fej
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "header", lab
    
End Sub
'Szerzõdéstípusú lapok Behelyettesítés és kitöltés
Public Sub Kitolt(Sbl As String, Adatok As Recordset)
    Form_Initialize
    all = 1
    Sablon = Konyvtar & "Sablonok\" & Sbl
    Fajl = "c:\Windows\Temp\" & TmpGeneral(Sablon)
    
    MsgBox Sablon & " - " & Fajl
    
    Behelyettesit Sablon, Fajl, Adatok
    MsgBox "behelyettesitve"
    bongeszo.Navigate2 Fajl
    Me.Show 'vbModal
    ido.Enabled = True
End Sub

'kitolt2
Public Sub Kitolt2(Sbl As String, p As String)
    Form_Initialize
    Dim Adatok As New ADODB.Recordset
    
    SQL_p p, Adatok
    'MsgBox Sor.RecordCount
    'nyomtatasikep.Kitolt "adasveteli.htm", Sor
    'Sor.Close
    'Me.Hide
    
    all = 1
    Sablon = Konyvtar & "Sablonok\" & Sbl
    Fajl = "c:\Windows\Temp\" & TmpGeneral(Sablon)
    
    MsgBox Sablon & " - " & Fajl
    
    Behelyettesit Sablon, Fajl, Adatok
    MsgBox "behelyettesitve"
    bongeszo.Navigate2 Fajl
    Me.Show 'vbModal
    ido.Enabled = True
    Adatok.Close
End Sub

Public Sub szamla(szamla As String)
    Form_Initialize
    all = 1
    
    bongeszo.Navigate2 szamla
    
    
    'Egyedi margók és fejlécek
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_bottom", "0.28504"
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_left", "0.15370"
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_right", "0.17244"
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "margin_top", "0.16500"
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "footer", ""
    SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", "header", ""
    Me.Show 'vbModal
    ido.Enabled = True
End Sub
'Behelyettesítés a sablonba
Private Sub Behelyettesit(Sablon As String, Fajl As String, Adatok As Recordset)
'On error goto hiba
    Dim Sor As String
        Open Sablon For Input As 1
        Open Fajl For Output As 2
            Do While Not EOF(1)
                Line Input #1, Sor
                    Print #2, Helyettesit(Sor, Adatok)
            Loop
        Close 2
        Close 1
Exit Sub
    Close 1
    Close 2
End Sub

'SZövegek behelyettesítése
Public Function Helyettesit(Sor As String, Adatok As Recordset) As String
On Error GoTo Hiba:
    Dim Kezd As Long, Veg As Long
    Dim kerdes As String, Koze As String
    
    Kezd = InStr(1, Sor, "[#!")
    Veg = InStr(1, Sor, "!#]")
    
    If Kezd * Veg > 0 Then
        kerdes = Mid(Sor, Kezd + 3, Veg - Kezd - 3)
        Koze = Mid(kerdes, 6)
        Select Case Mid(kerdes, 1, 5)
            Case "SZSZ_" 'SZámot szöveggé
                Dim Szam As Object
                Set Szam = CreateObject("Szamok.Irasa")
                Koze = Szam.szamot_szovegge(Oszlop(Adatok, Koze))
            Case "AZONO" 'AZonosító választó
               If Oszlop(Adatok, "MAGAN") = True Then
                    Koze = "szem.ig. sz: " & Nstr(Adatok!ESZEM)
               Else
                    Koze = "adószám: " & Nstr(Adatok!EADO)
               End If
            Case "DATE_"
                Koze = Format(Date, Koze)
            Case "TIME_"
                Koze = Format(Time, Koze)
            Case "XJEL_"
                If IxElem(Adatok!Ix, CByte(Koze)) Then Koze = "X" Else Koze = "0"
            Case Else
                Koze = Oszlop(Adatok, kerdes)
        End Select
            
        Helyettesit = Helyettesit(Mid(Sor, 1, Kezd - 1) & Koze & Mid(Sor, Veg + 3, Len(Sor)), Adatok)
    Else
            Helyettesit = Sor
    End If
Exit Function
Hiba:
    MsgBox "Szólj ÁKOSNAK:" & Err.Number & " " & Err.Description
End Function
Private Sub ido_Timer()
    Command1_Click
    ido.Enabled = False
    Me.Hide
End Sub
