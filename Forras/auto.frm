VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form adatlap 
   Caption         =   "Autó adatlapja"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame lap 
      Caption         =   "Frame1"
      Height          =   7575
      Index           =   7
      Left            =   10320
      TabIndex        =   91
      Top             =   4080
      Width           =   7695
      Begin VB.CommandButton muv_mod 
         Caption         =   "Hely módosítása"
         Height          =   495
         Left            =   3720
         TabIndex        =   96
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CommandButton muv_modosit 
         Caption         =   "Az autó nyilvántartási adatainak módosítása"
         Height          =   495
         Left            =   720
         TabIndex        =   95
         Top             =   2880
         Width           =   5655
      End
      Begin VB.CommandButton muv_torol 
         Caption         =   "Az autó minden adatának törlése a nyilvántartásból"
         Height          =   495
         Left            =   720
         TabIndex        =   94
         Top             =   3600
         Width           =   5655
      End
      Begin VB.CommandButton muv_hely 
         Caption         =   "Mutasd hol van!"
         Height          =   495
         Left            =   720
         TabIndex        =   93
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CommandButton muv_selejt 
         Caption         =   "Teljes autó selejtezése"
         Height          =   495
         Left            =   720
         TabIndex        =   92
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label ohaszon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6000
         TabIndex        =   103
         Top             =   4320
         Width           =   720
      End
      Begin VB.Label obevetel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4320
         TabIndex        =   102
         Top             =   4320
         Width           =   720
      End
      Begin VB.Label okiad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1680
         TabIndex        =   101
         Top             =   4320
         Width           =   720
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Haszon:"
         Height          =   195
         Index           =   44
         Left            =   5280
         TabIndex        =   100
         Top             =   4320
         Width           =   585
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bevétel:"
         Height          =   195
         Index           =   43
         Left            =   3000
         TabIndex        =   99
         Top             =   4320
         Width           =   585
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Összes kiadás:"
         Height          =   195
         Index           =   42
         Left            =   360
         TabIndex        =   98
         Top             =   4320
         Width           =   1065
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Statisztikák:"
         Height          =   195
         Index           =   41
         Left            =   360
         TabIndex        =   97
         Top             =   4800
         Width           =   855
      End
   End
   Begin VB.Frame lap 
      Caption         =   "Képek"
      Height          =   1455
      Index           =   6
      Left            =   6840
      TabIndex        =   86
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Frame lap 
      Caption         =   "Frame1"
      Height          =   4215
      Index           =   3
      Left            =   2280
      TabIndex        =   11
      Top             =   1200
      Width           =   9375
      Begin VB.TextBox valto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   121
         Top             =   600
         Width           =   1095
      End
      Begin MSComctlLib.ImageList allapotok 
         Left            =   15240
         Top             =   11520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "auto.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "auto.frx":0352
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "auto.frx":06A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "auto.frx":09F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "auto.frx":0D48
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox km 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   118
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox kerekek 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   89
         Text            =   "4"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton allapotlap_ment 
         Caption         =   "Mentés"
         Height          =   375
         Left            =   6120
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
      Begin MSComctlLib.TreeView fa 
         Height          =   2895
         Left            =   120
         TabIndex        =   119
         Top             =   1080
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5106
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "allapotok"
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Váltó kód:"
         Height          =   255
         Left            =   2280
         TabIndex        =   122
         Top             =   600
         Width           =   735
      End
      Begin VB.Label sugo 
         Caption         =   "Álljon a kiválasztott tételre és üsse le  az alábbi billentyûket F1 - nincs F2-ép F3-sérült"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   120
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Futott KM:"
         Height          =   255
         Left            =   240
         TabIndex        =   117
         Top             =   600
         Width           =   735
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gumi (db)"
         Height          =   195
         Index           =   40
         Left            =   240
         TabIndex        =   90
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame lap 
      Caption         =   "Árazás"
      Height          =   10095
      Index           =   5
      Left            =   6360
      TabIndex        =   10
      Top             =   3360
      Width           =   10575
      Begin VB.Frame ar_felirat 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6360
         TabIndex        =   126
         Top             =   360
         Width           =   3735
         Begin VB.Label cimke 
            AutoSize        =   -1  'True
            Caption         =   "Ára"
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
            Index           =   39
            Left            =   0
            TabIndex        =   130
            Top             =   0
            Width           =   300
         End
         Begin VB.Label cimke 
            AutoSize        =   -1  'True
            Caption         =   "Gyári"
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
            Index           =   46
            Left            =   600
            TabIndex        =   129
            Top             =   0
            Width           =   450
         End
         Begin VB.Label cimke 
            AutoSize        =   -1  'True
            Caption         =   "Utángy"
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
            Index           =   47
            Left            =   1320
            TabIndex        =   128
            Top             =   0
            Width           =   615
         End
         Begin VB.Label cimke 
            AutoSize        =   -1  'True
            Caption         =   "Gyári szám"
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
            Index           =   48
            Left            =   2520
            TabIndex        =   127
            Top             =   0
            Width           =   930
         End
      End
      Begin VB.CommandButton ment_arak 
         BackColor       =   &H00004000&
         Caption         =   "Mentés"
         Height          =   375
         Left            =   2880
         TabIndex        =   116
         Top             =   240
         Width           =   1815
      End
      Begin VB.PictureBox arazo_godor 
         BorderStyle     =   0  'None
         Height          =   7455
         Left            =   120
         ScaleHeight     =   7455
         ScaleWidth      =   10095
         TabIndex        =   109
         Top             =   720
         Width           =   10095
         Begin VB.PictureBox arazo_frm 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   6255
            Left            =   0
            ScaleHeight     =   6255
            ScaleWidth      =   10095
            TabIndex        =   110
            Top             =   0
            Width           =   10095
            Begin VB.TextBox gysz 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   8400
               TabIndex        =   123
               Text            =   "0"
               Top             =   -240
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.TextBox ara 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   6240
               TabIndex        =   113
               Text            =   "0"
               Top             =   -240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox gyari 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   6960
               TabIndex        =   112
               Text            =   "0"
               Top             =   -240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox utangyartott 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   7680
               TabIndex        =   111
               Text            =   "0"
               Top             =   -240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label cikkszam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "cikkszam"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   0
               TabIndex        =   115
               Top             =   -120
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label megnevezes 
               Caption         =   "Label1"
               Height          =   195
               Index           =   0
               Left            =   720
               TabIndex        =   114
               Top             =   -120
               Visible         =   0   'False
               Width           =   5520
            End
         End
      End
      Begin VB.VScrollBar ar_csuszka 
         Height          =   7455
         Left            =   10200
         TabIndex        =   108
         Top             =   720
         Width           =   255
      End
      Begin VB.ComboBox arkategoria 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   240
         Width           =   735
      End
      Begin VB.Label cimke 
         Alignment       =   2  'Center
         Caption         =   "Csak akkor árazhatja be az autót, ha már kitöltötte az állapotlapot!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1335
         Index           =   49
         Left            =   360
         TabIndex        =   125
         Top             =   2040
         Width           =   8535
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Árkategória:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   34
         Left            =   240
         TabIndex        =   106
         Top             =   360
         Width           =   1290
      End
   End
   Begin VB.Frame lap 
      Caption         =   "Eladó"
      Height          =   8415
      Index           =   1
      Left            =   9480
      TabIndex        =   9
      Top             =   1200
      Width           =   9375
      Begin VB.TextBox marka 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   85
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox megj_a 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   240
         TabIndex        =   84
         Top             =   3720
         Width           =   5295
      End
      Begin VB.TextBox torzskonyv 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   74
         Top             =   4560
         Width           =   3975
      End
      Begin VB.TextBox forgalmi 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   73
         Top             =   4920
         Width           =   3975
      End
      Begin VB.TextBox bon_forg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   72
         Top             =   5280
         Width           =   3255
      End
      Begin VB.TextBox kivonas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   71
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox ar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   70
         Text            =   "0"
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox bon_szam 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   69
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox nyszam 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   68
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox datum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   67
         Top             =   6480
         Width           =   1215
      End
      Begin VB.TextBox ido 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   66
         Top             =   6480
         Width           =   1215
      End
      Begin VB.TextBox kategoria 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   65
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox szine 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   64
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox hajtoanyag 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   63
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox allam 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   62
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox tipus 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   61
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox rendszam 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   47
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox alvaz 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   46
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox motor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   45
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox motorkod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   44
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox henger 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   43
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox evjarat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   42
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox tomeg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         TabIndex        =   41
         Text            =   "0"
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label uzenet 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Uzenet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   120
         TabIndex        =   88
         Top             =   7080
         Width           =   6975
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Megjegyzés:"
         Height          =   195
         Index           =   38
         Left            =   240
         TabIndex        =   87
         Top             =   3480
         Width           =   885
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Törzskönyvszám:"
         Height          =   195
         Index           =   37
         Left            =   240
         TabIndex        =   83
         Top             =   4560
         Width           =   1230
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forgalmi engedély:"
         Height          =   195
         Index           =   36
         Left            =   240
         TabIndex        =   82
         Top             =   4920
         Width           =   1320
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forgalmi engedély jogosultja:"
         Height          =   195
         Index           =   35
         Left            =   240
         TabIndex        =   81
         Top             =   5280
         Width           =   2025
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ár:"
         Height          =   195
         Index           =   33
         Left            =   3960
         TabIndex        =   80
         Top             =   5760
         Width           =   195
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bontási átvételi igazolás:"
         Height          =   195
         Index           =   32
         Left            =   240
         TabIndex        =   79
         Top             =   6120
         Width           =   1740
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nyszám:"
         Height          =   195
         Index           =   31
         Left            =   3600
         TabIndex        =   78
         Top             =   6120
         Width           =   600
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dátum:"
         Height          =   195
         Index           =   30
         Left            =   1560
         TabIndex        =   77
         Top             =   6480
         Width           =   510
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Végleges Kivonás dátuma:"
         Height          =   195
         Index           =   29
         Left            =   240
         TabIndex        =   76
         Top             =   5760
         Width           =   1890
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Idõ:"
         Height          =   195
         Index           =   27
         Left            =   3960
         TabIndex        =   75
         Top             =   6480
         Width           =   270
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gyártmány:"
         Height          =   195
         Index           =   26
         Left            =   240
         TabIndex        =   60
         Top             =   240
         Width           =   795
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Típus:"
         Height          =   195
         Index           =   25
         Left            =   600
         TabIndex        =   59
         Top             =   600
         Width           =   450
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Színe:"
         Height          =   195
         Index           =   24
         Left            =   480
         TabIndex        =   58
         Top             =   2760
         Width           =   450
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kategória:"
         Height          =   195
         Index           =   23
         Left            =   240
         TabIndex        =   57
         Top             =   960
         Width           =   720
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rendszám:"
         Height          =   195
         Index           =   22
         Left            =   240
         TabIndex        =   56
         Top             =   3120
         Width           =   795
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Államjelzés:"
         Height          =   195
         Index           =   21
         Left            =   2760
         TabIndex        =   55
         Top             =   3120
         Width           =   810
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alvázszám:"
         Height          =   195
         Index           =   20
         Left            =   240
         TabIndex        =   54
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motorszám:"
         Height          =   195
         Index           =   19
         Left            =   240
         TabIndex        =   53
         Top             =   1680
         Width           =   810
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motorkód:"
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   52
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hengerûrtartalom:"
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   51
         Top             =   2400
         Width           =   1260
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hajtóanyag:"
         Height          =   195
         Index           =   28
         Left            =   2640
         TabIndex        =   50
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Évjárat:"
         Height          =   195
         Index           =   15
         Left            =   2880
         TabIndex        =   49
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saját tömege:"
         Height          =   195
         Index           =   18
         Left            =   2760
         TabIndex        =   48
         Top             =   2400
         Width           =   975
      End
   End
   Begin VB.Frame lap 
      Caption         =   "Eladó"
      Height          =   7215
      Index           =   2
      Left            =   5640
      TabIndex        =   8
      Top             =   2880
      Width           =   7935
      Begin VB.TextBox telazon 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   124
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox orszag 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   104
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox vnev 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   26
         Top             =   240
         Width           =   5415
      End
      Begin VB.TextBox knev 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Top             =   600
         Width           =   5415
      End
      Begin VB.TextBox varos 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   24
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox irszam 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   23
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox cim 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   1680
         Width           =   5415
      End
      Begin VB.TextBox email 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Top             =   3360
         Width           =   5415
      End
      Begin VB.TextBox tel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox allampolg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   19
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox ktj 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   3840
         Width           =   5415
      End
      Begin VB.TextBox ado 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   2520
         Width           =   5415
      End
      Begin VB.TextBox szemelyi 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox kuj 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   4200
         Width           =   5415
      End
      Begin VB.TextBox fax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         TabIndex        =   14
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox megj 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   4680
         Width           =   5415
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ország:"
         Height          =   195
         Index           =   45
         Left            =   360
         TabIndex        =   105
         Top             =   960
         Width           =   540
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vezetéknév:"
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   40
         Top             =   240
         Width           =   900
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keresztnév:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   39
         Top             =   600
         Width           =   840
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Város:"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   38
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Irányítószám:"
         Height          =   195
         Index           =   3
         Left            =   4440
         TabIndex        =   37
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cím:"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   36
         Top             =   1680
         Width           =   315
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adó-szám:"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   35
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   34
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefon:"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   33
         Top             =   3000
         Width           =   585
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Személyi:"
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   32
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Állampolgárság:"
         Height          =   195
         Index           =   9
         Left            =   3600
         TabIndex        =   31
         Top             =   2160
         Width           =   1110
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KÜJ:"
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   30
         Top             =   4200
         Width           =   345
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KTJ:"
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   29
         Top             =   3840
         Width           =   330
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         Height          =   195
         Index           =   12
         Left            =   3840
         TabIndex        =   28
         Top             =   3000
         Width           =   300
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Megjegyzés:"
         Height          =   195
         Index           =   13
         Left            =   360
         TabIndex        =   27
         Top             =   4680
         Width           =   885
      End
   End
   Begin VB.CommandButton bezar 
      Caption         =   "Bezár"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Frame lap 
      Caption         =   "Bontási napló"
      Height          =   7215
      Index           =   4
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7575
      Begin VB.CommandButton bontasi_ment 
         Caption         =   "Bontási napló mentése"
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   6600
         Width           =   1815
      End
      Begin VB.TextBox mennyiseg 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3600
         TabIndex        =   3
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Szárazrafektetési napló"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label suly 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "suly"
         Height          =   195
         Index           =   0
         Left            =   6720
         TabIndex        =   4
         Top             =   2640
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label hullfel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "hullfel"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   2640
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin MSComctlLib.TabStrip fulek 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   14208
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gépjármû adatlapja"
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Eladó adatai"
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Állapotfelmérõ lap"
            Object.Tag             =   "3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Szárazrafektetési napló"
            Object.Tag             =   "4"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Alkatrész árazás"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Képek"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Egyéb"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "adatlap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Tetejem = 10000
Const Allapotbehuz = 2000
Const Szel = 0 '7215
Const Magassag = 0   '7575

Dim SID As Long
Dim all As Byte
Dim Kinek As Byte
Dim EladoID As Long

'Speciális fül betöltése
Public Sub MegnyitFul(Melyiket As Long, Ful As Byte, Optional Hova As Byte)
    Megnyit Melyiket, Hova
    
    fulek.Tabs(Ful).Selected = True
    fulek_Click
End Sub

'Autó adatlapjának megnyitása
Public Sub Megnyit(Melyiket As Long, Optional Hova As Byte)
    Dim Sor As New ADODB.Recordset
    
    'Form_initialize
    all = 0
    Pozicional
    uzenet = ""
    
    'arkategoria.Clear
    arkategoria.List(0) = "A"
    arkategoria.ItemData(0) = Asc(arkategoria.List(0))
    arkategoria.List(1) = "B"
    arkategoria.ItemData(1) = Asc(arkategoria.List(1))
    arkategoria.List(2) = "C"
    arkategoria.ItemData(2) = Asc(arkategoria.List(2))
    arkategoria.ListIndex = 0
    
    'ff
    SID = Melyiket
    Kinek = Hova
    
    SQL_p "SELECT * FROM autok WHERE id=" & SID, Sor
        EladoID = Sor!Elado
        Me.Caption = Nstr(Sor!nyszam) & " Adatlapja"
        Partner_Load EladoID, Me
        'bontasi
        Allapotlap_load
        
        marka.Text = Ertek("markak", "id", Nstr(Sor!marka), "marka")
        tipus.Text = Ertek("tipusok", "id", Nstr(Sor!tipus), "tipus")
        
        kategoria.Text = Ertek("kategoria", Nstr(Sor!kategoria), "id", "nev")
        evjarat.Text = Nstr(Sor!evjarat)
        rendszam.Text = Nstr(Sor!rendszam)
        allam.Text = Nstr(Sor!allam)
        alvaz.Text = Nstr(Sor!alvaz)
        motor.Text = Nstr(Sor!motor)
        motorkod.Text = Nstr(Sor!motorkod)
        szine.Text = Nstr(Sor!szine)
        tomeg.Text = Nstr(Sor!tomeg)
        henger.Text = Nstr(Sor!henger)
        hajtoanyag.Text = Nstr(Sor!hajtoanyag)
        
        torzskonyv.Text = Nstr(Sor!torzskonyv)
        forgalmi.Text = Nstr(Sor!forgalmi)
        bon_forg.Text = Nstr(Sor!bon_forg)
        kivonas.Text = Nstr(Sor!kivonas)
        megj_a.Text = Nstr(Sor!megj)
        
        bon_szam.Text = Nstr(Sor!bon_szam)
        nyszam.Text = Nstr(Sor!nyszam)
        datum.Text = Nstr(Sor!datum)
        ido.Text = Nstr(Sor!ido)
        
        valto.Text = Nstr(Sor!valto)
        km.Text = Nstr(Sor!km)
        ar.Text = Nstr(Sor!ar)
        
        ''Üzenetek beállítása
        If Not Sor!bontva Then
            LUzenet Me.uzenet, "Az autó még nincs szárazrafektetve!"
        Else
            
        End If
        
        If Not Sor!allapotlap Then
            'Árazás
            'ment_arak.Visible = False
            arazo_godor.Visible = False
            ar_csuszka.Visible = False
            ar_felirat.Visible = False
            
            
            LUzenet Me.uzenet, "Amíg az állapotlap nincs kitöltve, addig az alkatrészei nem kerülnek a nyilvántartásba"
        Else
            allapotlap_ment.Visible = False
            
            'Árazás
            arazo_godor.Visible = True
            ar_csuszka.Visible = True
            ar_felirat.Visible = True
            'ment_arak.Visible = True
            Jelol Me.arkategoria, Sor!arkategoria
            
            
            'Állapotlap
            Allapotlap_Betolt
        End If
    Sor.Close
    
    Me.Width = 11220
    Me.Height = 10220


    Me.Show
End Sub

Private Sub allapotlap_ment_Click()
    Dim p As String
    Dim Sor As New ADODB.Recordset
    Dim Kitoltve As Boolean
    Dim i As Integer
    
    Kitoltve = Ertek("autok", "id", CStr(SID), "allapotlap")
    FSQL "UPDATE autok SET km='" & km.Text & "', valto='" & valto.Text & "' WHERE id=" & SID
    
    If Not Kitoltve Then
        'Kerekek raktárkészletbe pakolása
        For i = 0 To CInt(kerekek.Text) - 1
            'FSQL "INSERT into raktarkeszlet (tipus, alkatresz, auto, allapot, hianyos, suly, elkelt, ewc) VALUES (0, " & GumiID & ", " & SID & " , 1, FALSE, " & Ertek("ewc", "id", "7", "szorzo") & " , FALSE, 7 )"
            FSQL "INSERT into raktarkeszlet (tipus, alkatresz, auto, allapot, hianyos, suly, elkelt, ewc) VALUES (0, " & GumiID & ", " & SID & " , 1, FALSE, 0, FALSE, 7 )"
        Next i
        
        For i = 1 To fa.Nodes.Count
            If Mid(fa.Nodes(i).Key, 1, 1) = "r" And fa.Nodes(i).Image > 1 Then
                p = "INSERT into raktarkeszlet (tipus, alkatresz, auto, allapot, ewc) VALUES (0, " & Mid(fa.Nodes(i).Key, 2) & ", " & SID & ", " & fa.Nodes(i).Image - 1 & ", " & fa.Nodes(i).Tag & " )"
                Debug.Print fa.Nodes(i).Text
                'Debug.Print p
                
                FSQL p
            End If
        Next i
        
    End If
    
    FSQL "UPDATE autok SET allapotlap=1 where id=" & SID
    
    MentKaszniTomege SID
    MsgBox "Az adatok elmentve"
    
    'Újratöltés
    Megnyit SID, Kinek
End Sub

Private Sub ar_csuszka_Change()
On Error Resume Next
    arazo_frm.Top = 1 * ar_csuszka.Value * (cikkszam(0).Top - cikkszam(1).Top)
End Sub

Private Sub bezar_Click()
    Unload Me
End Sub

Private Sub bontasi_ment_Click()
    Dim i As Long
    Dim p As String
    'Dim Teljes_tomeg As Double
    Dim Sor As New ADODB.Recordset
    Dim bontva As Boolean
    
    bontva = Ertek("autok", "id", CStr(SID), "bontva")
    'MInden hulladék és autóadat törlése a kocsitól
    SQL_p "DELETE * From raktarkeszlet where (tipus=1 or tipus=2) and auto=" & SID, Sor
    SQL_p "UPDATE autok SET bontva=1 where id=" & SID, Sor

    
    Rekord.CursorLocation = adUseClient
    sql_parancs ("SELECT * FROM ewc where bontas=TRUE")
    If Not Rekord.EOF Then
        Rekord.MoveFirst
        Do While Not Rekord.EOF
            i = Rekord!Id
            If hullfel(i).Tag <> 0 Then
                If mennyiseg(i).Text <> "" And IsNumeric(mennyiseg(i).Text) Then
                    p = "UPDATE raktarkeszlet SET suly=" & Vesszotlenito(CDbl(mennyiseg(i).Text * mennyiseg(i).Tag)) & " WHERE id=" & hullfel(i).Tag
                Else
                    'p = "DELETE * FROM raktarkeszlet WHERE id=" & hullfel(i).Tag
                    p = "UPDATE raktarkeszlet SET suly=0 WHERE id=" & hullfel(i).Tag
                End If
                FSQL p
            Else
                If False Then
                    If mennyiseg(i).Text <> "" And IsNumeric(mennyiseg(i).Text) Then
                        p = "UPDATE raktarkeszlet SET suly=" & Vesszotlenito(CDbl(mennyiseg(i).Text * mennyiseg(i).Tag)) & " WHERE ewc=" & i & " and auto=" & SID & " and tipus=0"
                    Else
                        'p = "DELETE * FROM raktarkeszlet WHERE id=" & hullfel(i).Tag
                        MsgBox "ez nem jó"
                    End If
                    FSQL p
                Else
                    If mennyiseg(i).Text <> "" And IsNumeric(mennyiseg(i).Text) Then
                        p = "INSERT INTO raktarkeszlet (tipus, auto, ewc, suly) VALUES (" & Alakit(Rekord!termek, "0", "1") & ", " & SID & ", " & i & ", " & Vesszotlenito(CDbl(mennyiseg(i).Text * mennyiseg(i).Tag)) & ")"
                        FSQL p
                    End If
                End If
            End If
            Rekord.MoveNext
        Loop
    End If
    Rekord.Close
    
    'Kaszni adatainak kiírása 2-es típus
    MentKaszniTomege SID
    MsgBox "Az adatok elmentve"
End Sub

Private Sub Pozicional()
    Dim i As Byte
    For i = 1 To lap.Count
        lap(i).Caption = fulek.SelectedItem.Caption
        If i = fulek.SelectedItem.Index Then lap(i).Visible = True Else lap(i).Visible = False
    Next i
End Sub

Private Sub AllapotValtas(Melyik As Node, Mire As Byte)
    Melyik.Image = Mire
    Melyik.SelectedImage = Mire
    
    Dim i As Integer
    For i = 1 To fa.Nodes.Count
        If Nstr(fa.Nodes(i).Parent) = Melyik.Text Then
            fa.Nodes(i).Image = Mire
            fa.Nodes(i).SelectedImage = Mire
            If fa.Nodes(i).Children > 0 Then
                AllapotValtas fa.Nodes(i), Mire
            End If
        End If
    Next i
    
End Sub

Private Sub fa_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Byte
    Select Case KeyCode
        Case 112
            i = 1
        Case 113
            i = 2
        Case 114
            i = 3
        Case 27
            i = fa.SelectedItem.Image
            i = i + 1
            If i = 4 Then i = 1
        Case Else
            Exit Sub
    End Select
    AllapotValtas fa.SelectedItem, i
End Sub

Private Sub bontasi()
 'On Error GoTo hiba
    Dim Id As Long, p As String, Mag As Long
    Dim Sor As New ADODB.Recordset
    Dim s2 As New ADODB.Recordset
    
    Mag = 2500
    
    
    Rekord.CursorLocation = adUseClient
    sql_parancs ("SELECT * FROM ewc where bontas=TRUE order by id")
    If Not Rekord.EOF Then Rekord.MoveFirst
    Do While Not Rekord.EOF
            Id = Rekord!Id
            On Error GoTo letezik
            
            Load hullfel(Id)
            hullfel(Id).Caption = Nstr(Rekord!ewc) & Alakit(Nstr(Rekord!veszelyes), "*", "") & " - " & Nstr(Rekord!nev) & " (" & Nstr(Rekord!me) & ")"
            hullfel(Id).Top = Mag
            hullfel(Id).Tag = 0
            hullfel(Id).Visible = True
            
            
            Load mennyiseg(Id)
            mennyiseg(Id).Top = hullfel(Id).Top
            mennyiseg(Id).Text = 0
            mennyiseg(Id).Tag = Nstr(Rekord!szorzo)
            mennyiseg(Id).ToolTipText = Nstr(Rekord!ewc)
            mennyiseg_Change (Id)
            mennyiseg(Id).Visible = True
            
            Load suly(Id)
            suly(Id).Left = mennyiseg(Id).Left + mennyiseg(Id).Width + 300
            suly(Id).Top = hullfel(Id).Top
            suly(Id).Caption = 0
            suly(Id).Visible = True
            
            'Ha alkatrészként is el lenne tárolva
            'SQL_p "SELECT * FROM raktarkeszlet where tipus=0 and ewc=" & Id & " and auto=" & SID, s2
            'SQL_p "SELECT Sum(raktarkeszlet.suly) AS SumOfsuly " & _
            '    "From raktarkeszlet " & _
            '    "GROUP BY raktarkeszlet.tipus, raktarkeszlet.ewc, raktarkeszlet.auto " & _
            '    "HAVING (((raktarkeszlet.tipus)=0) AND ((raktarkeszlet.ewc)=10) AND ((raktarkeszlet.auto)=30)); ", s2

            'If Not s2.RecordCount > 0 Then
            '    s2.MoveFirst
            '    mennyiseg(Id).Text = s2.Fields(0).Value / Rekord!szorzo
            '    hullfel(Id).Tag = s2!Id
            'End If
            's2.Close
            
            Mag = Mag + 250
            Rekord.MoveNext
    Loop
    Rekord.Close
    
    'Sor.CursorLocation = adUseClient
    BetoltBontasi
Exit Sub
letezik:
     Unload hullfel(Id)
     Unload mennyiseg(Id)
     Unload suly(Id)
     Resume
End Sub

Private Sub Form_Resize()
On Error Resume Next
 Dim i As Byte
    fulek.Width = Me.ScaleWidth - (2 * fulek.Left)
    fulek.Height = Me.ScaleHeight - bezar.Height - 400
    
    Kozepre cimke(49), lap(0)
    bezar.Move (Me.ScaleWidth - bezar.Width) / 2, fulek.Top + fulek.Height + 100
    For i = 1 To lap.Count
        lap(i).Move 240, 480
        lap(i).Width = fulek.Width - (2 * fulek.Left)
        lap(i).Height = fulek.Height - lap(i).Top - fulek.Left
    Next i
    Pozicional
    
    fa.Height = lap(1).Height - fa.Top - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visszajelez Kinek, SID
End Sub

Private Sub fulek_Click()
    If fulek.SelectedItem.Tag = "o" Then Exit Sub
    Pozicional
    Select Case fulek.SelectedItem.Index
        Case 4
            bontasi
            'BetoltBontasi
        Case 5
            Arazo_Betolt
        Case 7
            KoltsegSzamitas
    End Select
    
End Sub

Private Sub ment_arak_Click()
    Dim i As Integer
    
    FSQL "UPDATE autok SET arkategoria=" & arkategoria.ItemData(arkategoria.ListIndex) & " WHERE id=" & SID
    
    For i = 1 To cikkszam.Count - 1
        FSQL ("UPDATE raktarkeszlet SET ar=" & Vesszotlenito(ara(i).Text) & ", gyari=" & Vesszotlenito(gyari(i).Text) & ", utan=" & Vesszotlenito(utangyartott(i).Text) & ", gyszam='" & gysz(i).Text & "' WHERE id=" & cikkszam(i).Tag)
    Next i
    
    MsgBox "Adatok elmentve!"
End Sub

Private Sub mennyiseg_Change(Index As Integer)
On Error Resume Next
    suly(Index).Caption = "~" & mennyiseg(Index).Tag * mennyiseg(Index).Text & " kg"
End Sub

Private Sub BetoltBontasi()
    Dim Id As Long
    Dim Sor As New ADODB.Recordset
    Dim s2 As New ADODB.Recordset

    'Új betöltõ
    SQL_p "SELECT * FROM ewc WHERE bontas=TRUE", Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        Do While Not Sor.EOF
            Id = Sor!Id
            If Sor!termek Then
                'tip 1 nem fog kelleni
                SQL_p "SELECT * FROM raktarkeszlet where (tipus=0 or tipus=1 or tipus=3) and ewc=" & Id & " and auto=" & SID, s2
            Else
                SQL_p "SELECT * FROM raktarkeszlet where tipus=1 and ewc=" & Id & " and auto=" & SID, s2
           End If
           
           If Not s2.EOF Then
                s2.MoveFirst
                mennyiseg(Id).Text = (s2!suly / Sor!szorzo)
                hullfel(Id).Tag = s2!Id
            Else
                mennyiseg(Id).Text = ""
                hullfel(Id).Tag = 0
            End If
            s2.Close
           
           Sor.MoveNext
        Loop
    End If
End Sub

Private Sub Allapotlap_Betolt()
    Dim Id As Long
    Dim Sor As New ADODB.Recordset
    
    For Id = 1 To fa.Nodes.Count
            fa.Nodes(Id).Image = 1
            fa.Nodes(Id).SelectedImage = 1
    Next Id
    
    SQL_p "SELECT * from raktarkeszlet where tipus=0 and auto=" & SID, Sor
    If Not Sor.EOF Then
            Sor.MoveFirst
            Do While Not Sor.EOF
                    If Sor!elkelt = True Then
                        ValtoztatKep Sor!alkatresz, 5
                    Else
                        ValtoztatKep Sor!alkatresz, Sor!allapot + 1
                    End If
                Sor.MoveNext
            Loop
            Sor.Close
    End If
    
Exit Sub
Hiba:
    Hiba Err.Number, "Szin Frissitési hiba"
End Sub
Private Sub ValtoztatKep(Alk As Long, kep As Byte)
    Dim i As Long
    For i = 1 To fa.Nodes.Count
        If fa.Nodes(i).Key = CStr("r" & Alk) Then
            fa.Nodes(i).Image = kep
            fa.Nodes(i).SelectedImage = kep
            
            JelolFelettes fa.Nodes(i).Root, fa.Nodes(i).Parent
            Exit Sub
        End If
    Next i
End Sub
Private Sub JelolFelettes(Gyoker As String, Kit As String)
On Error GoTo Hiba
    Dim i As Long
    For i = 1 To fa.Nodes.Count
        If fa.Nodes(i).Text = Kit And fa.Nodes(i).Root = Gyoker Then
            fa.Nodes(i).Image = 4
            fa.Nodes(i).SelectedImage = 4
            
            JelolFelettes fa.Nodes(i).Root, fa.Nodes(i).Parent
            Exit Sub
        End If
    Next i
Hiba:
End Sub
Private Sub Allapotlap_load()
    Dim Sor As New ADODB.Recordset
    Dim Akt As Node
    
    fa.Nodes.Clear
    
    SQL_p "SELECT * FROM focsop", Sor
    Sor.MoveFirst
    Do While Not Sor.EOF
        fa.Nodes.Add , , Nstr("f" & Sor!Id), Nstr(NKieg(Sor!cikkszam) & " - " & Sor!nev), 2, 2
        Sor.MoveNext
    Loop
    Sor.Close
    
    SQL_p "SELECT * FROM alcsop", Sor
    Sor.MoveFirst
    Do While Not Sor.EOF
        fa.Nodes.Add "f" & Sor!focsop, tvwChild, Nstr("a" & Sor!Id), Nstr(NKieg(Sor!cikkszam) & " - " & Sor!nev), 2, 2
        Sor.MoveNext
    Loop
    Sor.Close
    
    SQL_p "SELECT * FROM alkatresznevek", Sor
    If Not Sor.EOF Then Sor.MoveFirst
    Do While Not Sor.EOF
        Set Akt = fa.Nodes.Add("a" & Sor!alcsop, tvwChild, "r" & Nstr(Sor!Id), Nstr(NKieg(Sor!cikkszam) & " - " & Sor!nev), 2, 2)
        'Akt.Checked = sor.Fields(8).Value
        Akt.Tag = Sor!ewc
        Sor.MoveNext
    Loop
    Sor.Close
End Sub

Private Sub mennyiseg_Click(Index As Integer)
    'MsgBox hullfel(Index).Tag
End Sub

'Költségsszámítás
Private Sub KoltsegSzamitas()
    Dim ossz As Double
    Dim bevetel As Double
    Dim Sor As New ADODB.Recordset
    
    SQL_p "SELECT Sum([ar]*(1+([afa]/100))) AS Kif1, raktarkeszlet.auto, raktarkeszlet.elkelt, raktarkeszlet.selejt, raktarkeszlet.sztorno, raktarkeszlet.tipus " & _
            "From raktarkeszlet " & _
            "GROUP BY raktarkeszlet.auto, raktarkeszlet.elkelt, raktarkeszlet.selejt, raktarkeszlet.sztorno, raktarkeszlet.tipus " & _
            "HAVING (((raktarkeszlet.auto)=" & SID & ") AND ((raktarkeszlet.elkelt)=True) AND ((raktarkeszlet.selejt)=False) AND ((raktarkeszlet.sztorno)=False) AND ((raktarkeszlet.tipus)=0)); ", Sor
    
    okiad.Caption = Ertek("autok", "id", CStr(SID), "ar")
    
    
    If Sor.RecordCount > 0 Then
        Sor.MoveFirst
        obevetel.Caption = Sor.Fields(0).Value
    Else
        obevetel.Caption = 0
    End If
    ohaszon.Caption = obevetel.Caption - okiad.Caption
    Sor.Close
End Sub

Private Sub muv_modosit_Click()
    Dim u As String
    
    u = Ujsor(u, "FIGYELEM! KÉREM OLVASSA EL FIGYELMESEN!")
    u = Ujsor(u, vbCrLf)
    u = Ujsor(u, "ÖN MOST AZ AUTÓ ADATAINAK MÓDOSÍTÁSÁT VÁLASZTOTTA!")
    u = Ujsor(u, "AMENNYIBEN ÖN VÉLETLENÜL KATTINTOTT ERRE A GOMBRA, AKKOR ITT MOST KILÉPHET MÉG.")
    u = Ujsor(u, "CSAK AKKOR FOLYTASSA EZT A MÛVELETET, HA TELJESEN BIZTOS ABBA, HOGY MIT CSINÁL!")
    u = Ujsor(u, "KATTINTSON AZ OK-RA HA FOLYTATNI KÍVÁNJA A MÛVELETET!")
    
    If MsgBox(u, vbOKCancel + vbExclamation, "Autó adatainak módosítása") = vbOK Then
        auto.modosit SID
        Unload Me
    End If
End Sub

Private Sub muv_selejt_Click()
    If MsgBox("Biztos le akarja selejtezni az autót?", vbYesNoCancel, "Autó selejtezése") = vbYes Then
        SelejtezAuto SID, True
        'Unload Me
    End If
    
End Sub

Private Sub muv_torol_Click()
    Dim u As String
    
    u = Ujsor(u, "FIGYELEM! KÉREM OLVASSA EL FIGYELMESEN!")
    u = Ujsor(u, vbCrLf)
    u = Ujsor(u, "ÖN MOST AZ AUTÓ TÖRLÉSÉT VÁLASZTOTTA!")
    u = Ujsor(u, "AMENNYIBEN ÖN VÉLETLENÜL KATTINTOTT ERRE A GOMBRA, AKKOR ITT MOST KILÉPHET MÉG.")
    u = Ujsor(u, "CSAK AKKOR FOLYTASSA EZT A MÛVELETET, HA TELJESEN BIZTOS ABBA, HOGY MIT CSINÁL!")
    u = Ujsor(u, "KATTINTSON AZ OK-RA HA FOLYTATNI KÍVÁNJA A MÛVELETET!")
    
    If MsgBox(u, vbOKCancel + vbExclamation, "Autó adatainak törlés") = vbOK Then
        auto.torol SID
        Unload Me
    End If
End Sub

Private Sub Arazo_Betolt()
    Dim Sor As New ADODB.Recordset
    Dim i As Integer
    
    arazo_frm.Visible = False
    
    For i = 1 To cikkszam.Count - 1
            Unload cikkszam(i)
            Unload megnevezes(i)
            Unload ara(i)
            Unload gyari(i)
            Unload utangyartott(i)
            Unload gysz(i)
    Next i
    
    i = 1
    
    
    '                       0                   1               2                   3                   4                      5            6               7               8               9                   10                  11                      12                      13                  14                      15                      16
    SQL_p "SELECT raktarkeszlet.auto, raktarkeszlet.ar, raktarkeszlet.gyari, raktarkeszlet.utan, raktarkeszlet.cikkszam, focsop.nev, alcsop.nev, alkatresznevek.nev, raktarkeszlet.id, focsop.cikkszam, alcsop.cikkszam, alkatresznevek.cikkszam, raktarkeszlet.gyszam, raktarkeszlet.allapot, raktarkeszlet.elkelt, raktarkeszlet.selejt, raktarkeszlet.sztorno " & _
            "FROM (focsop INNER JOIN (alcsop INNER JOIN alkatresznevek ON alcsop.id = alkatresznevek.alcsop) ON focsop.id = alcsop.focsop) INNER JOIN raktarkeszlet ON alkatresznevek.id = raktarkeszlet.alkatresz " & _
            "WHERE (((raktarkeszlet.auto)=" & SID & ") and raktarkeszlet.tipus=0); ", Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        Do While Not Sor.EOF
        
            Load cikkszam(i)
            Load megnevezes(i)
            Load ara(i)
            Load gyari(i)
            Load utangyartott(i)
            Load gysz(i)
            
            cikkszam(i).Caption = Nstr(NKieg(Sor.Fields(9).Value) & NKieg(Sor.Fields(10).Value) & NKieg(Sor.Fields(11).Value))
            cikkszam(i).Move cikkszam(0).Left, cikkszam(i - 1).Top + cikkszam(i - 1).Height + 200
            cikkszam(i).Tag = Nstr(Sor.Fields(8).Value)
            cikkszam(i).ForeColor = SzinAllapot(Sor.Fields(13).Value)
            cikkszam(i).Visible = True
            
            megnevezes(i).Caption = Nstr(Sor.Fields(5).Value & "/ " & Sor.Fields(6).Value & "/ " & Sor.Fields(7).Value)
            megnevezes(i).Move megnevezes(0).Left, cikkszam(i).Top
            megnevezes(i).ForeColor = SzinAllapot(Sor.Fields(13).Value)
            megnevezes(i).Visible = True
            
            ara(i).Text = Nstr(Sor.Fields(1).Value)
            ara(i).Move ara(0).Left, cikkszam(i).Top
            ara(i).Enabled = Not (Sor.Fields(14).Value Or Sor.Fields(15).Value Or Sor.Fields(16).Value)
            ara(i).Visible = True
            
            gyari(i).Text = Nstr(Sor.Fields(2).Value)
            gyari(i).Move gyari(0).Left, cikkszam(i).Top
            gyari(i).Enabled = Not (Sor.Fields(14).Value Or Sor.Fields(15).Value Or Sor.Fields(16).Value)
            gyari(i).Visible = True
            
            utangyartott(i).Text = Nstr(Sor.Fields(3).Value)
            utangyartott(i).Move utangyartott(0).Left, cikkszam(i).Top
            utangyartott(i).Enabled = Not (Sor.Fields(14).Value Or Sor.Fields(15).Value Or Sor.Fields(16).Value)
            utangyartott(i).Visible = True
            
            gysz(i).Text = Nstr(Sor.Fields(12).Value)
            gysz(i).Move gysz(0).Left, cikkszam(i).Top
            gysz(i).Enabled = Not (Sor.Fields(14).Value Or Sor.Fields(15).Value Or Sor.Fields(16).Value)
            gysz(i).Visible = True
            
            i = i + 1
            Sor.MoveNext
        Loop
    End If
    
    arazo_frm.Height = utangyartott(i - 1).Top + utangyartott(i - 1).Height
    arazo_frm.Visible = True
    ar_csuszka.Min = 0
    'ar_csuszka.Max = (arazo_frm.Height - arazo_godor.Height) / 500
    ar_csuszka.Max = i
    ar_csuszka.SmallChange = Abs(ar_csuszka.Max / i)
    ar_csuszka.LargeChange = Abs(ar_csuszka.SmallChange * 3)
    
End Sub
