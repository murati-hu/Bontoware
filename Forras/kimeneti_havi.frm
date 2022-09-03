VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form kimeneti_havi 
   Caption         =   "Kimeneti üzemkönyv"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   12780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bezar 
      Caption         =   "Bezár"
      Height          =   615
      Left            =   960
      TabIndex        =   13
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Frissítés"
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
   Begin MSComctlLib.ListView alkatreszek 
      Height          =   6735
      Left            =   120
      TabIndex        =   2
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dátum"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Számlaszám"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Kiszállítás napja"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Hulladék típusa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "EWC kód"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Hulladékot átvevõ cég"
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "További hasznosítás módja"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Célállomás"
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Elszállított hulladék súlya"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker datum 
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "yyyy. MM."
      Format          =   45416451
      CurrentDate     =   38559
   End
   Begin VB.Label zarolt 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   11880
      TabIndex        =   12
      Top             =   8520
      Width           =   90
   End
   Begin VB.Label elozo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   11880
      TabIndex        =   11
      Top             =   8160
      Width           =   90
   End
   Begin VB.Label Label6 
      Caption         =   "Zárolt:"
      Height          =   255
      Left            =   10800
      TabIndex        =   10
      Top             =   8520
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Áthozott az elõzõ hónapról:"
      Height          =   255
      Left            =   9360
      TabIndex        =   9
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "kg"
      Height          =   255
      Index           =   2
      Left            =   12120
      TabIndex        =   8
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "kg"
      Height          =   255
      Index           =   1
      Left            =   12120
      TabIndex        =   7
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "kg"
      Height          =   255
      Index           =   0
      Left            =   12120
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   7800
      Width           =   90
   End
   Begin VB.Label Label2 
      Caption         =   "Összesen:"
      Height          =   255
      Left            =   10560
      TabIndex        =   3
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
Attribute VB_Name = "kimeneti_havi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim elem As ListItem

Private Sub alkatreszek_DblClick()
    'MsgBox alkatreszek.SelectedItem.Text
    'MsgBox alkatreszek.SelectedItem.ListSubItems(1)
    If alkatreszek.SelectedItem.ListSubItems(1) = 0 Then
        kimeneti.meghiv (alkatreszek.SelectedItem.Text)
    Else
        hulladek_szamla.modosit alkatreszek.SelectedItem.ListSubItems(2), 80
    End If
End Sub

Private Sub bezar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Frissit
End Sub

Private Sub datum_Change()
    Frissit
End Sub

Public Sub meghiv(Optional Honap As Date)
    datum.Value = Honap
    Frissit
    
    Me.Show
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        alkatreszek.Move 120, 840, Me.ScaleWidth - 2 * alkatreszek.Left, Me.ScaleHeight - alkatreszek.Top - 1300
        bezar.Move 960, Me.ScaleHeight - 1000, bezar.Width, bezar.Height
        Label2.Move Me.ScaleWidth - 4000, Me.ScaleHeight - 1100
        ossztomeg.Move Me.ScaleWidth - 1100, Me.ScaleHeight - 1100
        Label3(0).Move Me.ScaleWidth - 700, Me.ScaleHeight - 1100
        Label4.Move Me.ScaleWidth - 4000, Me.ScaleHeight - 800
        elozo.Move Me.ScaleWidth - 1100, Me.ScaleHeight - 800
        Label5(1).Move Me.ScaleWidth - 700, Me.ScaleHeight - 800
        Label6.Move Me.ScaleWidth - 4000, Me.ScaleHeight - 400
        zarolt.Move Me.ScaleWidth - 1100, Me.ScaleHeight - 400
        Label7(2).Move Me.ScaleWidth - 700, Me.ScaleHeight - 400
End Sub

Private Sub Form_Initialize()
    'adatmotor.Megnyitas
    'datum.Value = Date
    'Frissit
End Sub

Private Sub Frissit()
    Dim osszsuly As Double, napiossz As Double, p As String, Sor As New ADODB.Recordset, nap As String, meddig As Integer, Seged As Double, seged2 As Double, SID As Integer
    osszsuly = 0
    ossztomeg.Visible = False
    alkatreszek.ListItems.Clear
    alkatreszek.Visible = False
    '################### Beallitom, hogy az aktualis honapban csak az aktualis napig lassuk a napi kimenetiket ############
    If (datum.Year = Year(Date)) And (datum.Month = Month(Date)) Then
        meddig = Day(Date)
    Else
        If (datum.Month = 1) Or (datum.Month = 3) Or (datum.Month = 5) Or (datum.Month = 7) Or (datum.Month = 8) Or (datum.Month = 10) Or (datum.Month = 12) Then
            meddig = 31
        Else
            If (datum.Month = 2) And (datum.Year = (datum.Year \ 4) * 4) Then '###########*******Ez így hülyeség, JAVÍTANDÓ
                meddig = 29
            Else
                If (datum.Month = 2) Then
                    meddig = 28
                Else
                    meddig = 30
                End If
            End If
        End If
    End If
    If ((datum.Year = Year(Date)) And (datum.Month > Month(Date))) Or (datum.Year > Year(Date)) Then meddig = 0
    '#################### Napi osszegzes ################
    For i = 1 To meddig
    
    
        '#################### Alkatresz kimeneti ################
        'p = "SELECT Sum(raktarkeszlet.suly) AS SumOfsuly, szamla.kelt, szamla.id "
        'p = p & "FROM szamla INNER JOIN (raktarkeszlet INNER JOIN elkelt ON raktarkeszlet.id = elkelt.tetel) ON szamla.id = elkelt.szamla "
        'p = p & "GROUP BY szamla.kelt, Year([szamla].[kelt]), Month([szamla].[kelt]), Day([szamla].[kelt]), szamla.id "
        'p = p & "HAVING (((Year([szamla].[kelt]))=" & datum.Year & ") AND ((Month([szamla].[kelt]))=" & datum.Month & ") AND ((Day([szamla].[kelt]))=" & i & "))"
        
        p = "SELECT Sum(raktarkeszlet.suly) AS SumOfsuly, szamla.kelt "
        p = p & "FROM szamla INNER JOIN (raktarkeszlet INNER JOIN elkelt ON raktarkeszlet.id = elkelt.tetel) ON szamla.id = elkelt.szamla "
        p = p & "Where (((Year([szamla].[kelt])) =" & datum.Year & ") And ((Month([szamla].[kelt])) =" & datum.Month & ") And ((Day([szamla].[kelt])) =" & i & ") And ((raktarkeszlet.tipus) = 0)) "
        p = p & "GROUP BY szamla.kelt"

        
        SQL_p p, Sor
        napiossz = 0
        If Not Sor.EOF Then
            Sor.MoveFirst
            napiossz = Sor.Fields(0)
            nap = Sor.Fields(1)
            'SID = sor.Fields(2)
        Else
            nap = datum.Year & "." & datum.Month & "." & i & "."
        End If
        Sor.Close
        Set elem = alkatreszek.ListItems.Add(, , nap)
            elem.ListSubItems.Add , , "0"
            SID = 0
            elem.ListSubItems.Add , , SID
            elem.ListSubItems.Add , , "-"
            elem.ListSubItems.Add , , i & "."
            elem.ListSubItems.Add , , "napi kimeneti üzemkönyv szerint"
            elem.ListSubItems.Add , , "-"
            elem.ListSubItems.Add , , "lakossági"
            elem.ListSubItems.Add , , "újrahasználat"
            elem.ListSubItems.Add , , "-"
            elem.ListSubItems.Add , , napiossz
            osszsuly = osszsuly + napiossz
            
            
        '#################### Hulladek kimeneti ################
     
        'p = "SELECT szamla.kelt, szamla.szam, ewc.nev, ewc.ewc, partnerek.vnev, partnerek.knev, partnerek.cim, partnerek.vhk, partnerek_1.vnev, partnerek_1.knev, partnerek_1.cim, partnerek_1.vhk, Sum(hulladek.suly) AS SumOfsuly, szamla.id "
        'p = p & "FROM partnerek AS partnerek_1 INNER JOIN (partnerek INNER JOIN (((kimenet INNER JOIN szamla ON kimenet.szamla = szamla.id) INNER JOIN hulladek ON kimenet.hulladek = hulladek.id) INNER JOIN ewc ON hulladek.ewc = ewc.id) ON partnerek.id = szamla.szallito) ON partnerek_1.id = szamla.vevo "
        'p = p & "GROUP BY szamla.kelt, szamla.szam, ewc.nev, ewc.ewc, partnerek.vnev, partnerek.knev, partnerek.cim, partnerek.vhk, partnerek_1.vnev, partnerek_1.knev, partnerek_1.cim, partnerek_1.vhk, hulladek.auto, szamla.id "
        'p = p & "HAVING (((szamla.kelt)=#" & datum.Month & "/" & i & "/" & datum.Year & "#) AND ((hulladek.auto)=-1))"
        
        p = "SELECT szamla.kelt, szamla.szam, ewc.nev, ewc.ewc, partnerek.vnev, partnerek.knev, partnerek.cim, partnerek.vhk, partnerek_1.vnev, partnerek_1.knev, partnerek_1.cim, partnerek_1.vhk, Sum(raktarkeszlet.suly) AS SumOfsuly, szamla.id, raktarkeszlet.tipus, partnerek.kuj, partnerek.ktj, partnerek_1.kuj, partnerek_1.ktj "
        p = p & "FROM partnerek AS partnerek_1 INNER JOIN (partnerek INNER JOIN (((elkelt INNER JOIN szamla ON elkelt.szamla = szamla.id) INNER JOIN raktarkeszlet ON elkelt.tetel = raktarkeszlet.id) INNER JOIN ewc ON raktarkeszlet.ewc = ewc.id) ON partnerek.id = szamla.szallito) ON partnerek_1.id = szamla.vevo "
        p = p & "GROUP BY szamla.kelt, szamla.szam, ewc.nev, ewc.ewc, partnerek.vnev, partnerek.knev, partnerek.cim, partnerek.vhk, partnerek_1.vnev, partnerek_1.knev, partnerek_1.cim, partnerek_1.vhk, szamla.id, raktarkeszlet.auto, raktarkeszlet.tipus, partnerek.kuj, partnerek.ktj, partnerek_1.kuj, partnerek_1.ktj "
        p = p & "HAVING (((szamla.kelt)=#" & datum.Month & "/" & i & "/" & datum.Year & "#) AND ((raktarkeszlet.auto)=-1) AND ((raktarkeszlet.tipus)=1))"

        
        SQL_p p, Sor
        If Not Sor.EOF Then
            Sor.MoveFirst
            Do While Not Sor.EOF
                'Ákoskájé
                'MsgBox Sor.Fields(2).
                
                Set elem = alkatreszek.ListItems.Add(, , Sor.Fields(0)) 'Datum, rejtett
                    elem.ListSubItems.Add , , "1"
                    elem.ListSubItems.Add , , Sor.Fields(13)
                    elem.ListSubItems.Add , , Nstr(Sor.Fields(1)) & "/" & datum.Year            'Szamlaszam
                    elem.ListSubItems.Add , , i & "."                                           'Honap hanyadik napja
                    elem.ListSubItems.Add , , Nstr(Sor.Fields(2))                               'Ewc megnevezes
                    elem.ListSubItems.Add , , Nstr(Sor.Fields(3))                               'EWC kod
                    'elem.ListSubItems.Add , , Sor.Fields(4) & " " & Sor.Fields(5)              'Szallito neve
                    'elem.ListSubItems.Add , , Nstr(Sor.Fields(6))                              'Szallito cime
                    'elem.ListSubItems.Add , , Nstr(Sor.Fields(7))                              'Szallito engedely szama
                    elem.ListSubItems.Add , , Sor.Fields(4) & " " & Sor.Fields(5) & ", " & Sor.Fields(6) & ", eng szam: " & Sor.Fields(7) & ", KÜJ: " & Sor.Fields(15) & ", KTJ: " & Sor.Fields(16)
                    elem.ListSubItems.Add , , "megsemmisítés"
                    'elem.ListSubItems.Add , , Sor.Fields(8) & " " & Sor.Fields(9)              'Vevo neve
                    'elem.ListSubItems.Add , , Nstr(Sor.Fields(10))                             'Vevo cime
                    'elem.ListSubItems.Add , , Nstr(Sor.Fields(11))                             'Vevo engedely szama
                    elem.ListSubItems.Add , , Sor.Fields(8) & " " & Sor.Fields(9) & ", " & Sor.Fields(10) & ", eng szam: " & Sor.Fields(11) & ", KÜJ: " & Sor.Fields(17) & ", KTJ: " & Sor.Fields(18)
                    elem.ListSubItems.Add , , Abs(Sor.Fields(12))
                    osszsuly = osszsuly + Abs(Sor.Fields(12))
                Sor.MoveNext
            Loop
        End If
        Sor.Close

    Next i
    '################# Elozo havi zarolt ###############
    p = "SELECT Sum(raktarkeszlet.suly) AS SumOfsuly "
    p = p & "FROM raktarkeszlet INNER JOIN (szamla INNER JOIN elkelt ON szamla.id = elkelt.szamla) ON raktarkeszlet.id = elkelt.tetel "
    p = p & "WHERE (((szamla.kelt)<#" & datum.Month & "/1/" & datum.Year & "#) AND ((raktarkeszlet.elkelt)=True))"
    SQL_p p, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        If Nstr(Sor.Fields(0)) = "" Then
            Seged = 0
        Else
            Seged = Sor.Fields(0)
        End If
        'elozo.Caption = seged 'Sor.Fields(0)
        'zarolt.Caption = seged + osszsuly   'Sor.Fields(0) + osszsuly
    End If
    Sor.Close
    p = "SELECT Sum(raktarkeszlet.suly) AS SumOfsuly "
    p = p & "FROM raktarkeszlet INNER JOIN (szamla INNER JOIN elkelt ON szamla.id = elkelt.szamla) ON raktarkeszlet.id = elkelt.tetel "
    p = p & "WHERE (((szamla.kelt)<#" & datum.Month & "/1/" & datum.Year & "#) AND ((raktarkeszlet.tipus)=1))"

    SQL_p p, Sor
    If Not Sor.EOF Then
        Sor.MoveFirst
        If Nstr(Sor.Fields(0)) = "" Then
            seged2 = 0
        Else
            seged2 = Abs(Sor.Fields(0))
        End If
    End If
    Sor.Close
    
    elozo.Caption = Seged + seged2
    zarolt.Caption = Seged + seged2 + osszsuly
    alkatreszek.Visible = True
    ossztomeg.Caption = osszsuly
    ossztomeg.Visible = True
End Sub
