VERSION 5.00
Begin VB.Form abk 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Futtat"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "abk.frx":0000
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "SQL utasítás:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "abk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    FSQL Text1.Text
End Sub
