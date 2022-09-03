VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form bemeneti 
   Caption         =   "Bemeneti üzemkönyv"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   12780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox isdatum 
      Alignment       =   1  'Right Justify
      Caption         =   "Szûrés"
      Height          =   495
      Left            =   10440
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton bezar 
      Caption         =   "Bezár"
      Height          =   615
      Left            =   960
      TabIndex        =   11
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Frissítés"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
   Begin MSComctlLib.ListView autok 
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   11880
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Azonosíó"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nyilvántartási szám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Átadás napja"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Jármûtípus és gyártmány"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Alvázszám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Motorszám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Rendszám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Leadó/tulajdonos neve, címe"
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Beszállított roncs önsúlya"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker ig 
      Height          =   375
      Left            =   8280
      TabIndex        =   12
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   57278465
      CurrentDate     =   38561
   End
   Begin MSComCtl2.DTPicker tol 
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   57278465
      CurrentDate     =   38561
   End
   Begin VB.Label Label4 
      Caption         =   "Áthozott az elõzõ idõszakról:"
      Height          =   255
      Left            =   9240
      TabIndex        =   17
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Tól"
      Height          =   255
      Left            =   5040
      TabIndex        =   15
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "Ig"
      Height          =   255
      Left            =   7800
      TabIndex        =   14
      Top             =   240
      Width           =   255
   End
   Begin VB.Label zarolt 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   11880
      TabIndex        =   10
      Top             =   8520
      Width           =   90
   End
   Begin VB.Label elozo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   11880
      TabIndex        =   9
      Top             =   8160
      Width           =   90
   End
   Begin VB.Label Label5 
      Caption         =   "Zárolt:"
      Height          =   255
      Left            =   10800
      TabIndex        =   8
      Top             =   8520
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "kg"
      Height          =   255
      Index           =   2
      Left            =   12120
      TabIndex        =   7
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "kg"
      Height          =   255
      Index           =   1
      Left            =   12120
      TabIndex        =   6
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "kg"
      Height          =   255
      Index           =   0
      Left            =   12120
      TabIndex        =   4
      Top             =   7800
      Width           =   255
   End
   Begin VB.Label ossztomeg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "# ##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1038
         SubFormatType   =   1
      EndProperty
      Height          =   195
      Left            =   11880
      TabIndex        =   3
      Top             =   7800
      Width           =   90
   End
   Begin VB.Label Label2 
      Caption         =   "Összesen:"
      Height          =   255
      Left            =   10560
      TabIndex        =   2
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Dátum:"
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "bemeneti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim elem As ListItem

Private Sub autok_DblClick()
    'MsgBox autok.SelectedItem.Text
    'MsgBox alkatreszek.SelectedItem.ListSubItems(1)
    adatlap.Megnyit autok.SelectedItem.Text, 100
End Sub

Private Sub bezar_Click()
    Unload Me
End Sub

Public Sub menubol_hiv()
    Dim tol As String, ig As String
    If Month(Date) <= 3 Then
        tol = Year(Date) & ".01.01."
        ig = Year(Date) & ".03.31."
    Else
        If Month(Date) <= 6 Then
            tol = Year(Date) & ".04.01."
            ig = Year(Date) & ".06.30."
        Else
            If Month(Date) <= 9 Then
                tol = Year(Date) & ".07.01."
                ig = Year(Date) & ".09.30."
            Else
                tol = Year(Date) & ".10.01."
                ig = Year(Date) & ".12.31."
            End If
        End If
    End If
    
    bemeneti.meghiv CDate(tol), CDate(ig)
End Sub

Private Sub Command1_Click()
    Frissit
End Sub

Public Sub meghiv(Date1 As Date, date2 As Date)
    
    tol.Value = Date1
    ig.Value = date2
    isdatum.Value = 1
    'Frissit
    
    Me.Show
End Sub

Private Sub Form_Initialize()
    'adatmotor.Megnyitas
    'tol.Value = Date
    'ig.Value = Date
    'Frissit
End Sub

Private Sub isdatum_Click()
    Frissit
End Sub

Private Sub tol_Change()
    If isdatum.Value = 1 Then Frissit
End Sub

Private Sub ig_Change()
    If isdatum.Value = 1 Then Frissit
End Sub

Private Sub Frissit()
    Dim Sor As New ADODB.Recordset, p As String, osszsuly As Double, p1 As String, p2 As String
    osszsuly = 0
    autok.ListItems.Clear
    autok.Visible = False
    p = "SELECT autok.id, autok.nyszam, autok.datum, autok.ido, markak.marka, tipusok.tipus, autok.alvaz, autok.motor, autok.rendszam, partnerek.vnev, partnerek.knev, partnerek.irszam, partnerek.varos, partnerek.cim, autok.tomeg "
    p = p & "FROM markak INNER JOIN (tipusok INNER JOIN (partnerek INNER JOIN autok ON partnerek.id = autok.elado) ON tipusok.id = autok.tipus) ON markak.id = tipusok.marka "
    'p = p & "WHERE (((autok.datum)<#8/13/2005#) AND ((autok.datum)>#8/1/2005#))"
    p1 = "1=1"
    p2 = "1=1"
    
    If isdatum.Value = 1 Then p1 = "(autok.datum)>=" & DatumAtir(tol.Value)
    If isdatum.Value = 1 Then p2 = "(autok.datum)<=" & DatumAtir(ig.Value)
    
    p = p & "WHERE ((" & p1 & ") AND (" & p2 & "))"
    'MsgBox p
    SQL_p p, Sor
    If Not Sor.EOF Then Sor.MoveFirst
    Do While Not Sor.EOF
        Set elem = autok.ListItems.Add(, , Sor.Fields(0))
            elem.ListSubItems.Add , , Sor.Fields(1)
            elem.ListSubItems.Add , , Sor.Fields(2) & " " & Sor.Fields(3)
            elem.ListSubItems.Add , , Sor.Fields(4) & " " & Sor.Fields(5)
            elem.ListSubItems.Add , , Nstr(Sor.Fields(6))
            elem.ListSubItems.Add , , Sor.Fields(7)
            elem.ListSubItems.Add , , Nstr(Sor.Fields(8))
            elem.ListSubItems.Add , , Sor.Fields(9) & " " & Sor.Fields(10) & " " & Sor.Fields(11) & " " & Sor.Fields(12) & " " & Sor.Fields(13)
            elem.ListSubItems.Add , , Sor.Fields(14) & " kg"
            osszsuly = osszsuly + Sor.Fields(14)
        Sor.MoveNext
    Loop
    Sor.Close
    
    SQL_p "SELECT Sum(autok.tomeg) AS SumOftomeg From autok WHERE (((autok.datum)<" & DatumAtir(tol.Value) & "))", Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        If Nstr(Sor.Fields(0)) = "" Then
            elozo.Caption = 0
        Else
            elozo.Caption = Sor.Fields(0)
        End If
    End If
    If isdatum.Value = 0 Then elozo.Caption = 0
    Sor.Close
    ossztomeg.Caption = osszsuly
    zarolt.Caption = osszsuly + elozo.Caption
    autok.Visible = True
End Sub
