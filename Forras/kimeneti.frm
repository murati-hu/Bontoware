VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form kimeneti 
   Caption         =   "Újrahasználati napi kimeneti üzemkönyv"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   14205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox arakkal 
      Caption         =   "Nyomtatás árakkal és számlaszámmal"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton nyomtat 
      Caption         =   "Nyomtat"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton bezar 
      Caption         =   "Bezár"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   7800
      Width           =   1215
   End
   Begin MSComctlLib.ListView alkatreszek 
      Height          =   6855
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   12091
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Azonosító"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Alkatrész"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Suly (kg)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ára"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Számla típusa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Számla száma"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Eladó"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker datum 
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   21037057
      CurrentDate     =   38586
   End
   Begin VB.CommandButton frissites 
      Caption         =   "Frissítés"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Ft"
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   8040
      Width           =   255
   End
   Begin VB.Label oar 
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
      Left            =   5400
      TabIndex        =   8
      Top             =   8040
      Width           =   90
   End
   Begin VB.Label Label3 
      Caption         =   "kg"
      Height          =   255
      Left            =   6720
      TabIndex        =   3
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
      Left            =   5400
      TabIndex        =   2
      Top             =   7800
      Width           =   90
   End
   Begin VB.Label Label2 
      Caption         =   "Összesen:"
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Dátum:"
      Height          =   255
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "kimeneti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ido As Date

Private Sub bezar_Click()
    Unload Me
End Sub

Private Sub frissites_Click()
    Frissit
End Sub

Private Sub datum_Change()
    ido = datum.Value
    Frissit
End Sub

Private Sub Form_Initialize()
    'ido = datum.Value
    'adatmotor.Megnyitas
    'Frissit
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        alkatreszek.Move 120, 840, Me.ScaleWidth - 2 * alkatreszek.Left, Me.ScaleHeight - alkatreszek.Top - 1000
        bezar.Move (Me.ScaleWidth - bezar.Width) / 2, Me.ScaleHeight - bezar.Height - 300, bezar.Width, bezar.Height
        nyomtat.Move bezar.Left - 150 - nyomtat.Width, bezar.Top
        Label2.Move Me.ScaleWidth - 2000, Me.ScaleHeight - 700
        ossztomeg.Move Label2.Left + Label2.Width + 100, Me.ScaleHeight - 700
        oar.Move ossztomeg.Left, ossztomeg.Top + ossztomeg.Height + 100
        Label3.Move Me.ScaleWidth - 300, Me.ScaleHeight - 700
        Label4.Move Label3.Left, Label3.Top + Label3.Height + 100
End Sub

Private Sub Frissit()
    Dim osszsuly As Double, p As String, Sor As New ADODB.Recordset
    oar.Caption = 0
    osszsuly = 0
    ossztomeg.Visible = False
    
    'p = "SELECT raktarkeszlet.id, markak.marka, tipusok.tipus, alcsop.nev, alkatresznevek.nev, raktarkeszlet.suly, szamla.szam, szamla.tipus, szamla.vevo, raktarkeszlet.ar, szamla.uid "
    'p = p & "FROM alcsop INNER JOIN ((markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN (autok INNER JOIN ((elkelt INNER JOIN szamla ON elkelt.szamla = szamla.id) INNER JOIN (alkatresznevek INNER JOIN raktarkeszlet ON alkatresznevek.id = raktarkeszlet.alkatresz) ON elkelt.tetel = raktarkeszlet.id) ON autok.id = raktarkeszlet.auto) ON tipusok.id = autok.tipus) ON alcsop.id = alkatresznevek.alcsop "
    'p = p & "WHERE (((szamla.kelt)=" & DatumAtir(ido) & ") AND ((szamla.tipus)=0))"
    
    p = "SELECT raktarkeszlet.id, markak.marka, tipusok.tipus, alcsop.nev, alkatresznevek.nev, raktarkeszlet.suly, szamla.szam, szamla.tipus, szamla.vevo, raktarkeszlet.ar, felhasznalok.nev "
    p = p & "FROM (alcsop INNER JOIN ((markak INNER JOIN tipusok ON markak.id = tipusok.marka) INNER JOIN (autok INNER JOIN ((elkelt INNER JOIN szamla ON elkelt.szamla = szamla.id) INNER JOIN (alkatresznevek INNER JOIN raktarkeszlet ON alkatresznevek.id = raktarkeszlet.alkatresz) ON elkelt.tetel = raktarkeszlet.id) ON autok.id = raktarkeszlet.auto) ON tipusok.id = autok.tipus) ON alcsop.id = alkatresznevek.alcsop) INNER JOIN felhasznalok ON szamla.uid = felhasznalok.id "
    p = p & "WHERE (((szamla.tipus)=0) AND ((szamla.kelt)=" & DatumAtir(ido) & "))"


    
    SQL_p p, Sor
    'MsgBox Sor.RecordCount
    'MsgBox DatumAtir(ido)
    alkatreszek.ListItems.Clear
    'MsgBox p
    If Not Sor.EOF Then Sor.MoveFirst
    alkatreszek.Visible = False
    Do While Not Sor.EOF
        Set elem = alkatreszek.ListItems.Add(, , Sor.Fields(0).Value)
            elem.ListSubItems.Add , , Sor.Fields(1).Value & " " & Sor.Fields(2).Value & " " & Sor.Fields(3).Value & " " & Sor.Fields(4).Value
            elem.ListSubItems.Add , , Sor.Fields(5).Value & " kg"
            elem.ListSubItems.Add , , Sor.Fields(9).Value & " Ft"
            If Sor.Fields(8).Value = 34 Then
                Select Case Sor.Fields(7).Value
                    Case 0
                        elem.ListSubItems.Add , , "Nyugta"
                    Case 1
                        elem.ListSubItems.Add , , "Sztornózott nyugta"
                End Select
            Else
                Select Case Sor.Fields(7).Value
                    Case 0
                        elem.ListSubItems.Add , , "Számla"
                    Case 1
                        elem.ListSubItems.Add , , "Sztornózott számla"
                End Select
            End If
            elem.ListSubItems.Add , , Sor.Fields(6).Value & "/" & datum.Year
            elem.ListSubItems.Add , , Sor.Fields(10).Value
            osszsuly = osszsuly + Sor.Fields(5).Value
            oar.Caption = oar.Caption + Sor.Fields(9).Value
        Sor.MoveNext
    Loop
    Sor.Close
    alkatreszek.Visible = True
    ossztomeg.Caption = osszsuly
    ossztomeg.Visible = True
End Sub

Public Sub meghiv(Optional nap As Date)     '####### A form meghívása
    Form_Initialize                         '## Ha a havi kimenetibõl hívjuk meg, akkor onnan veszi át a napot
    If nap > 0 Then                         '## Ha a menübõl, akkor az aktuális nappal hívod meg
        ido = nap
        datum.Value = nap
    Else
        datum.Value = Date
        ido = datum.Value
    End If
    Frissit
    Me.Show
End Sub

Public Sub TobbOldalasSzamla(Kimenet As String, Forras As String, Optional DB As Integer)
    Dim i As Integer
    Open Kimenet For Output As 1
        Open Forras For Input As 2
                Do While Not EOF(2)
                Line Input #2, Sor
                Select Case Trim(Sor)
                    Case "<#!DATUM!#>"
                        Print #1, datum.Value
                    Case "<#!TARTALOM!#>"
                        For i = 1 To alkatreszek.ListItems.Count
                        With alkatreszek.ListItems(i)
                            Print #1, "<TR>"
                            Print #1, "  <td>" & .ListSubItems(1).Text
                            Print #1, "  </td>"
                            Print #1, "  <td>" & .ListSubItems(2).Text & "</td>"
                            If arakkal.Value = 1 Then
                                Print #1, "  <td>" & .ListSubItems(3).Text & "</td>"
                                Print #1, "  <td>" & .ListSubItems(5).Text & "</td>"
                            Else
                                Print #1, "  <td> - </td>"
                                Print #1, "  <td> - </td>"
                            End If
                            Print #1, "</TR>"
                        End With
                        Next i
                    Case "<#!OAR!#>"
                        If arakkal.Value = 1 Then Print #1, oar.Caption & " Ft"
                    Case "<#!OKG!#>"
                        Print #1, ossztomeg.Caption & " Kg"
                    Case Else
                        Print #1, Sor
                End Select
            Loop
        Close 2
    Close 1
    
    'nyomtatasikep.gombsor.Visible = False
    'nyomtatasikep.Show
    'nyomtatasikep.bongeszo.Navigate2 Kimenet
End Sub

Private Sub nyomtat_Click()
    Dim Sablon As String
    Dim Fajl As String
    
    Frissit
    
    Sablon = Konyvtar & "Sablonok\napikim.htm"
    Fajl = "c:\Windows\Temp\" & TmpGeneral(Sablon)
    
    TobbOldalasSzamla Fajl, Sablon
    nyomtatasikep.szamla Fajl
    
End Sub
