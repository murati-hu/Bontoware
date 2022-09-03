VERSION 5.00
Begin VB.Form bonto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bontó Adatai"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame eng_frm 
      Enabled         =   0   'False
      Height          =   7335
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   6615
      Begin VB.TextBox nev 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   28
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox rnev 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox ksh 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   26
         Top             =   1680
         Width           =   4935
      End
      Begin VB.TextBox kuj 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   25
         Top             =   2040
         Width           =   4935
      End
      Begin VB.TextBox ktj 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   2400
         Width           =   4935
      End
      Begin VB.TextBox orszag 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox irszam 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5880
         TabIndex        =   22
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox utca 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Top             =   3480
         Width           =   3135
      End
      Begin VB.TextBox hazszam 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   20
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox helyrajz 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox felelos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   18
         Top             =   4200
         Width           =   4935
      End
      Begin VB.TextBox beosztas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   4560
         Width           =   4935
      End
      Begin VB.TextBox tel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   5040
         Width           =   2175
      End
      Begin VB.TextBox fax 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   15
         Top             =   5040
         Width           =   2175
      End
      Begin VB.TextBox email 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   5400
         Width           =   4935
      End
      Begin VB.TextBox weblap 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   5760
         Width           =   4935
      End
      Begin VB.TextBox telepeng 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Top             =   6120
         Width           =   3855
      End
      Begin VB.TextBox vhull 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Top             =   6480
         Width           =   3855
      End
      Begin VB.TextBox knysz 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Top             =   6840
         Width           =   3855
      End
      Begin VB.ComboBox emberek 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5160
         TabIndex        =   9
         Text            =   "10"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox cg 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   1320
         Width           =   4935
      End
      Begin VB.TextBox ado 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox telazon 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   6
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ComboBox varos 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   3120
         Width           =   3135
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bontó neve:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   52
         Top             =   240
         Width           =   870
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rövid neve:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   51
         Top             =   600
         Width           =   870
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KSH:"
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   50
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vezetõ:"
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   49
         Top             =   4200
         Width           =   540
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KÜJ:"
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   48
         Top             =   2040
         Width           =   345
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ország:"
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   47
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Város"
         Height          =   195
         Index           =   7
         Left            =   840
         TabIndex        =   46
         Top             =   3120
         Width           =   405
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KTJ:"
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   45
         Top             =   2400
         Width           =   330
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Helyrajzi szám:"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   44
         Top             =   3840
         Width           =   1035
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Utca:"
         Height          =   195
         Index           =   10
         Left            =   960
         TabIndex        =   43
         Top             =   3480
         Width           =   390
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Irányítószám"
         Height          =   195
         Index           =   11
         Left            =   4920
         TabIndex        =   42
         Top             =   3120
         Width           =   885
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beosztása:"
         Height          =   195
         Index           =   12
         Left            =   600
         TabIndex        =   41
         Top             =   4560
         Width           =   780
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Házszám:"
         Height          =   195
         Index           =   13
         Left            =   4680
         TabIndex        =   40
         Top             =   3480
         Width           =   690
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefon:"
         Height          =   195
         Index           =   14
         Left            =   840
         TabIndex        =   39
         Top             =   5040
         Width           =   585
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         Height          =   195
         Index           =   15
         Left            =   3840
         TabIndex        =   38
         Top             =   5040
         Width           =   300
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   195
         Index           =   16
         Left            =   960
         TabIndex        =   37
         Top             =   5400
         Width           =   420
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weblap:"
         Height          =   195
         Index           =   17
         Left            =   840
         TabIndex        =   36
         Top             =   5760
         Width           =   600
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alkalmazottak:"
         Height          =   195
         Index           =   18
         Left            =   4080
         TabIndex        =   35
         Top             =   3840
         Width           =   1035
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telepengedély:"
         Height          =   195
         Index           =   19
         Left            =   1200
         TabIndex        =   34
         Top             =   6120
         Width           =   1095
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Veszélyeshulladék azonosító:"
         Height          =   195
         Index           =   20
         Left            =   240
         TabIndex        =   33
         Top             =   6600
         Width           =   2085
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KFF regszám:"
         Height          =   195
         Index           =   21
         Left            =   1320
         TabIndex        =   32
         Top             =   6960
         Width           =   960
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cégjegyzékszám:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   1230
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adószám:"
         Height          =   195
         Index           =   22
         Left            =   600
         TabIndex        =   30
         Top             =   960
         Width           =   690
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Település azon:"
         Height          =   195
         Index           =   23
         Left            =   3960
         TabIndex        =   29
         Top             =   2760
         Width           =   1125
      End
   End
   Begin VB.PictureBox felso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6945
      TabIndex        =   2
      Top             =   0
      Width           =   6975
      Begin VB.Label focim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bontó adatai"
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
         TabIndex        =   3
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.CommandButton bezar 
      Caption         =   "Mégse"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton ment 
      Caption         =   "Mentés"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   8280
      Width           =   1695
   End
End
Attribute VB_Name = "bonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bezar_Click()
    Unload Me
End Sub

Private Sub focim_DblClick()
    eng_frm.Enabled = True
End Sub

Private Sub Form_Load()
    felso = bontoware.lila
    Dim Sor As New ADODB.Recordset
    SQL_p "SELECT * FROM bonto where id=1", Sor
    Sor.MoveFirst
    nev.Text = Nstr(Sor!nev)
    rnev.Text = Nstr(Sor!rnev)
    ksh.Text = Nstr(Sor!ksh)
    kuj.Text = Nstr(Sor!kuj)
    ktj.Text = Nstr(Sor!ktj)
    orszag.Text = Nstr(Sor!orszag)
    Betolt Me.varos, "telepulesek", "telepules", "telepules"
    varos.Text = Nstr(Sor!varos)
    'Jelol Me.varos, Nstr(sor!varos)
    irszam.Text = Nstr(Sor!irszam)
    telazon.Text = Nstr(Sor!telazon)
    utca.Text = Nstr(Sor!utca)
    hazszam.Text = Nstr(Sor!hazszam)
    helyrajz.Text = Nstr(Sor!helyrajz)
    felelos.Text = Nstr(Sor!felelos)
    beosztas.Text = Nstr(Sor!beosztas)
    tel.Text = Nstr(Sor!tel)
    fax.Text = Nstr(Sor!fax)
    email.Text = Nstr(Sor!email)
    weblap.Text = Nstr(Sor!weblap)
    emberek.Text = Nstr(Sor!emberek)
    telepeng.Text = Nstr(Sor!telepeng)
    vhull.Text = Nstr(Sor!vhull)
    knysz.Text = Nstr(Sor!knysz)
    ado.Text = Nstr(Sor!ado)
    cg.Text = Nstr(Sor!cg)
    
End Sub

Private Sub ment_Click()
    Dim Sor As New ADODB.Recordset
    SQL_p "UPDATE bonto SET nev='" & nev.Text & "', rnev='" & rnev.Text & "', ksh='" & ksh.Text & "', kuj='" & kuj.Text & "', ktj='" & ktj.Text & "', orszag='" & orszag.Text & "', varos='" & varos.Text & "', telazon='" & telazon.Text & "', irszam='" & irszam.Text & "', utca='" & utca.Text & "', hazszam='" & hazszam.Text & "', helyrajz='" & helyrajz.Text & "', felelos='" & felelos.Text & "', beosztas='" & beosztas.Text & "', tel='" & tel.Text & "', fax='" & fax.Text & "', email='" & email.Text & "', weblap='" & weblap.Text & "', emberek='" & emberek.Text & "', telepeng='" & telepeng.Text & "', vhull='" & vhull.Text & "', knysz='" & knysz.Text & "', ado='" & ado.Text & "', cg='" & cg.Text & "' WHERE id=1", Sor
    MsgBox "Adatok elmentve"
End Sub

Private Sub varos_Change()
On Error Resume Next
    telazon.Text = varos.ItemData(varos.ListIndex)
End Sub

Private Sub varos_Click()
    varos_Change
End Sub
