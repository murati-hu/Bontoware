VERSION 5.00
Begin VB.Form alkatresz_eladasra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alkatrész adatai"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton muv_mod 
      Caption         =   "Hely módosítása"
      Height          =   375
      Left            =   4800
      TabIndex        =   23
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton muv_hely 
      Caption         =   "Mutasd hol van!"
      Height          =   375
      Left            =   4800
      TabIndex        =   22
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6585
      TabIndex        =   20
      Top             =   0
      Width           =   6615
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alkatrész adatai"
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
         Left            =   4200
         TabIndex        =   21
         Top             =   240
         Width           =   2205
      End
   End
   Begin VB.TextBox megj 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   6375
   End
   Begin VB.ComboBox allapot 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox hianyos 
      Caption         =   "H"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox tomeg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox ar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton mentes 
      Caption         =   "Adatok mentése"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton bezar 
      Caption         =   "Mégse"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton mentes_hozzaad 
      Caption         =   "Felvétel a számlára"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   2055
   End
   Begin VB.ComboBox afa 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2280
      Width           =   735
   End
   Begin VB.ComboBox szin 
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label hely 
      AutoSize        =   -1  'True
      Caption         =   "Megtalálható:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3600
      TabIndex        =   24
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Megjegyzés:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   885
   End
   Begin VB.Label gyszam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gyértésiszám"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1800
      TabIndex        =   18
      Top             =   840
      Width           =   1125
   End
   Begin VB.Label cikkszam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cikkszam:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   870
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Állapota:"
      Height          =   195
      Index           =   15
      Left            =   5400
      TabIndex        =   16
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label tipus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gyartmany"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label alkatresz 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alkatrész"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Áfakulcs:"
      Height          =   195
      Index           =   2
      Left            =   2280
      TabIndex        =   12
      Top             =   2040
      Width           =   660
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minõsítés:"
      Height          =   195
      Index           =   1
      Left            =   3360
      TabIndex        =   11
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tömeg:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Netto ár:"
      Height          =   195
      Index           =   12
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   615
   End
End
Attribute VB_Name = "alkatresz_eladasra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim all As Byte
Dim SID As Long
Dim Kinek As Byte



Private Sub bezar_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
    felso = bontoware.narancs
    all = 0
    SID = 0
    Betolt Me.afa, "afa", "afa", "afa"
    Szin_Betolt Me.szin
    
    allapot.List(0) = "nincs"
    allapot.ItemData(0) = 0
    allapot.List(1) = "ép"
    allapot.ItemData(1) = 1
    allapot.List(2) = "sérült"
    allapot.ItemData(1) = 2
    allapot.ListIndex = 1
    
End Sub

Public Sub Bearaz(Melyiket As Long, Optional Hova As Byte, Optional Hogy As Byte)
    Dim p As String
    Dim Sor As New ADODB.Recordset
    Form_Initialize
    all = Hogy
    
    If all = 2 Then '2-es számlán van már
        mentes_hozzaad.Visible = False
        mentes.Default = True
    End If
    
    SID = Melyiket
    Kinek = Hova
    
   
    p = "SELECT raktarkeszlet.ar, raktarkeszlet.afa, raktarkeszlet.suly,  raktarkeszlet.minosites, autok.nyszam, autok.evjarat, markak.marka, tipusok.tipus, focsop.nev, alcsop.nev, alkatresznevek.nev, raktarkeszlet.allapot, raktarkeszlet.hianyos, raktarkeszlet.cikkszam, raktarkeszlet.gyszam, raktarkeszlet.megj, focsop.cikkszam, alcsop.cikkszam, alkatresznevek.cikkszam, raktarkeszlet.hely "
    p = p & " FROM (markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN (focsop INNER JOIN (autok INNER JOIN ((alcsop INNER JOIN alkatresznevek ON alcsop.id = alkatresznevek.alcsop) INNER JOIN raktarkeszlet ON alkatresznevek.id = raktarkeszlet.alkatresz) ON autok.id = raktarkeszlet.auto) ON focsop.id = alcsop.focsop) ON tipusok.id = autok.tipus "
    p = p & " WHERE (((raktarkeszlet.id)=" & SID & "));"
    
    SQL_p p, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        ar.Text = Nstr(Sor.Fields(0).Value)
        afa.Text = Nstr(Sor.Fields(1).Value)
        tomeg.Text = Nstr(Sor.Fields(2).Value)
        szin.ListIndex = Sor.Fields(3).Value
        
        cikkszam.Caption = NKieg(Nstr(Sor.Fields(16).Value)) & NKieg(Nstr(Sor.Fields(17).Value)) & NKieg(Nstr(Sor.Fields(18).Value))
        gyszam.Caption = Nstr(Sor.Fields(14).Value)
        megj.Text = Nstr(Sor.Fields(15).Value)
        
        allapot.ListIndex = Sor.Fields(11).Value
        hianyos.Value = Alakit(Sor.Fields(12).Value, "1", "0")
        
        tipus.Caption = Nstr(Sor.Fields(6).Value & " - " & Sor.Fields(7).Value & "(" & Sor.Fields(5).Value & ") - " & Sor.Fields(4).Value)
        alkatresz.Caption = Nstr(Sor.Fields(8).Value & "/" & Sor.Fields(9).Value & "/" & Sor.Fields(10).Value)
        
        muv_hely.Caption = Nstr(Sor.Fields(19).Value)
    End If
    Sor.Close
    
    Me.Show vbModal
End Sub
Private Sub ment()
    Dim Sor As New ADODB.Recordset
    
    SQL_p "UPDATE raktarkeszlet SET minosites=" & szin.ListIndex & ", ar=" & Vesszotlenito(ar.Text) & ", afa=" & afa.List(afa.ListIndex) & ", suly=" & Vesszotlenito(tomeg.Text) & ", megj='" & megj.Text & "' WHERE id=" & SID, Sor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visszajelez Kinek, SID
End Sub

Private Sub mentes_Click()
    ment
    Unload Me
    If Kinek = 255 Then alkatresz_lista.Frissit
End Sub

Private Sub mentes_hozzaad_Click()
    Dim AID As Long
    AID = SID
    ment
    
    If all = 0 Then
        Unload Me
        szamlazo.UjGyorsTetel AID
    End If
    
    szamlazo.Beszur AID
    Unload Me
    
    If Kinek = 255 Then alkatresz_lista.Frissit
End Sub

Private Sub muv_hely_Click()
    raktar.Mutasd muv_hely.Caption
End Sub

Private Sub muv_mod_Click()
    raktar.HelyValaszt Me.muv_hely
    FSQL "UPDATE raktarkeszlet SET hely='" & muv_hely.Caption & "' where id=" & SID
End Sub
