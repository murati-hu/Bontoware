VERSION 5.00
Begin VB.Form partner_lap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Partner adatai"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6945
      TabIndex        =   36
      Top             =   0
      Width           =   6975
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partner adatlapja"
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
         Left            =   4320
         TabIndex        =   37
         Top             =   360
         Width           =   2370
      End
   End
   Begin VB.TextBox telazon 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox orszag 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   1800
      Width           =   5415
   End
   Begin VB.ComboBox varos 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Text            =   "varos"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox vhk 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   15
      Top             =   6000
      Width           =   5415
   End
   Begin VB.TextBox megj 
      Appearance      =   0  'Flat
      Height          =   765
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   6480
      Width           =   5415
   End
   Begin VB.CommandButton bezar 
      Caption         =   "Mégse"
      Height          =   375
      Left            =   3960
      TabIndex        =   18
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton felvesz 
      Caption         =   "Felvétel"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox fax 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4080
      TabIndex        =   11
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox kuj 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   5640
      Width           =   5415
   End
   Begin VB.TextBox szemelyi 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox ado 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   3480
      Width           =   5415
   End
   Begin VB.TextBox ktj 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Top             =   5280
      Width           =   5415
   End
   Begin VB.TextBox allampolg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox tel 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox email 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   4560
      Width           =   5415
   End
   Begin VB.TextBox cim 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   2640
      Width           =   4455
   End
   Begin VB.TextBox irszam 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox knev 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   5415
   End
   Begin VB.TextBox vnev 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   5415
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Település azonosító:"
      Height          =   195
      Index           =   16
      Left            =   3600
      TabIndex        =   35
      Top             =   2160
      Width           =   1470
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Orszag:"
      Height          =   195
      Index           =   15
      Left            =   480
      TabIndex        =   34
      Top             =   1800
      Width           =   540
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VHK:"
      Height          =   195
      Index           =   14
      Left            =   480
      TabIndex        =   33
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Megjegyzés:"
      Height          =   195
      Index           =   13
      Left            =   120
      TabIndex        =   32
      Top             =   6480
      Width           =   885
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
      Height          =   195
      Index           =   12
      Left            =   3720
      TabIndex        =   31
      Top             =   4200
      Width           =   300
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KTJ:"
      Height          =   195
      Index           =   11
      Left            =   600
      TabIndex        =   30
      Top             =   5280
      Width           =   330
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KÜJ:"
      Height          =   195
      Index           =   10
      Left            =   600
      TabIndex        =   29
      Top             =   5640
      Width           =   345
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Állampolgárság:"
      Height          =   195
      Index           =   9
      Left            =   3480
      TabIndex        =   28
      Top             =   3120
      Width           =   1110
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Személyi:"
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   27
      Top             =   3120
      Width           =   660
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefon:"
      Height          =   195
      Index           =   7
      Left            =   360
      TabIndex        =   26
      Top             =   4200
      Width           =   585
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      Height          =   195
      Index           =   6
      Left            =   480
      TabIndex        =   25
      Top             =   4560
      Width           =   465
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adó-szám:"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   24
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cím:"
      Height          =   195
      Index           =   4
      Left            =   1800
      TabIndex        =   23
      Top             =   2640
      Width           =   315
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Irányítószám:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   930
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Település:"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keresztnév:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   1320
      Width           =   840
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vezetéknév:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "partner_lap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SID As Long
Dim all As Byte '0új, 1mód, 2beszur
Dim Kinek As Byte

Private Sub bezar_Click()
    Unload Me
End Sub

Private Sub felvesz_Click()
    Dim p As String, i As Long
    If all = 0 Then
        p = "INSERT INTO partnerek (vnev, knev, orszag, varos, telazon, irszam, cim, ado, tel, fax, email, szemelyi, kuj, ktj, allampolg, megj, vhk) VALUES ('" & vnev.Text & "', '" & knev.Text & "', '" & orszag.Text & "', '" & Varos.Text & "', '" & Telazon.Text & "', '" & irszam.Text & "', '" & cim.Text & "', '" & ado.Text & "', '" & tel.Text & "', '" & fax.Text & "', '" & email.Text & "', '" & szemelyi.Text & "', '" & kuj.Text & "', '" & ktj.Text & "', '" & allampolg.Text & "', '" & megj.Text & "', '" & vhk.Text & "')"
        sql_parancs (p)
        
        VarosTanul Varos.Text, Telazon.Text
        
        If Kinek > 0 Then
            p = "SELECT * FROM partnerek WHERE vnev='" & vnev.Text & "' and  knev='" & knev.Text & "' and orszag='" & orszag.Text & "' and varos='" & Varos.Text & "' and telazon='" & Telazon.Text & "' and irszam='" & irszam.Text & "' and cim='" & cim.Text & "' and ado='" & ado.Text & "' and tel='" & tel.Text & "' and fax='" & fax.Text & "' and email='" & email.Text & "' and szemelyi='" & szemelyi.Text & "' and kuj='" & kuj.Text & "' and ktj='" & ktj.Text & "' and allampolg='" & allampolg.Text & "' and megj='" & megj.Text & "' and vhk='" & vhk.Text & "' order by id desc"
            'MsgBox p
            sql_parancs (p)
            If Not Rekord.EOF Then Rekord.MoveFirst
            i = Rekord!Id
            Rekord.Close
            Visszajelez Kinek, i
        End If
    Else
        p = "UPDATE partnerek SET vnev='" & vnev.Text & "',  knev='" & knev.Text & "', orszag='" & orszag.Text & "', varos='" & Varos.Text & "', telazon='" & Telazon.Text & "' ,irszam='" & irszam.Text & "', cim='" & cim.Text & "', ado='" & ado.Text & "', tel='" & tel.Text & "', fax='" & fax.Text & "', email='" & email.Text & "', szemelyi='" & szemelyi.Text & "', kuj='" & kuj.Text & "', ktj='" & ktj.Text & "', allampolg='" & allampolg.Text & "', megj='" & megj.Text & "', vhk='" & vhk.Text & "' where id=" & SID
        sql_parancs (p)
        
        If Kinek > 0 Then Visszajelez Kinek, SID
    End If
    
    Unload Me
End Sub

Public Sub modosit(Id As Long, Optional Hova As Byte)
    Form_Initialize
    If Not Partner_Load(Id, Me) Then MsgBox "Nincs ilyen rekord!" Else Me.Show
    all = 1
    Kinek = Hova
    SID = Id
End Sub

Private Sub Tisztit()
        vnev.Text = ""
        knev.Text = ""
        orszag.Text = "Magyarország"
        'varos.Text = "Szeged"
        'telazon.Text = "33367"
        'varos.Text = "Szeged"
        varos_Click
        irszam.Text = "0"
        cim.Text = ""
        ado.Text = ""
        tel.Text = ""
        fax.Text = ""
        email.Text = ""
        szemelyi.Text = ""
        kuj.Text = ""
        ktj.Text = ""
        allampolg.Text = "magyar"
        megj.Text = ""
        vhk.Text = ""
        Jelol Me.Varos, 33367
        Telazon.Text = 33367
End Sub

Public Sub uj(Optional Hova As Byte)
    Form_Initialize
    Kinek = Hova
    Me.Show
End Sub

Private Sub Form_Initialize()
    felso = bontoware.zold
    all = 0
    SID = 0
    Kinek = 0
    Betolt Me.Varos, "telepulesek", "telepules", "telepules"
    felvesz.Caption = "Felvétel"
    'allampolg.Text = "magyar"
    Tisztit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            felvesz_Click
        Case 27
            bezar_Click
    End Select
    
End Sub

Private Sub varos_Change()
On Error Resume Next
    Telazon.Text = Varos.ItemData(Varos.ListIndex)
End Sub
Private Sub varos_Click()
    varos_Change
End Sub
