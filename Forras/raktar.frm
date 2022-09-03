VERSION 5.00
Begin VB.Form raktar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   Caption         =   "Teleprendezés"
   ClientHeight    =   5205
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame vezerlok 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10215
      Begin VB.TextBox kereso 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label mutat 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   195
         Left            =   6000
         TabIndex        =   12
         Top             =   120
         Width           =   90
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Keresés helykód alapján:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1770
      End
   End
   Begin VB.CommandButton gomb 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar jb 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4080
      Width           =   5295
   End
   Begin VB.VScrollBar fl 
      Height          =   4095
      Left            =   5280
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox terulet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   1800
      MousePointer    =   2  'Cross
      ScaleHeight     =   2145
      ScaleWidth      =   3105
      TabIndex        =   0
      Top             =   720
      Width           =   3135
      Begin VB.PictureBox ja 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   2040
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox jf 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   2040
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox bf 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   1200
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox ba 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   1200
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox tarolo 
         Appearance      =   0  'Flat
         FillColor       =   &H8000000F&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   360
         MousePointer    =   1  'Arrow
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape keret 
         Height          =   375
         Left            =   1200
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Menu szerk_mnu 
      Caption         =   "Szerkesztés"
      Visible         =   0   'False
      Begin VB.Menu nev 
         Caption         =   "tulajdonságai"
         Enabled         =   0   'False
      End
      Begin VB.Menu v0 
         Caption         =   "-"
      End
      Begin VB.Menu enged_mnu 
         Caption         =   "Szerkesztés engedélyezése"
      End
      Begin VB.Menu nev_mnu 
         Caption         =   "Tároló átnevezése"
      End
      Begin VB.Menu meretez_mnu 
         Caption         =   "Tároló átméretez"
      End
      Begin VB.Menu szinez_mnu 
         Caption         =   "Színezés"
         Begin VB.Menu szin 
            Caption         =   "szin"
            Checked         =   -1  'True
            Index           =   0
         End
      End
      Begin VB.Menu v4 
         Caption         =   "-"
      End
      Begin VB.Menu klon_mnu 
         Caption         =   "Klónozás"
      End
      Begin VB.Menu v3 
         Caption         =   "-"
      End
      Begin VB.Menu torol_mnu 
         Caption         =   "Tároló törlése"
      End
   End
   Begin VB.Menu szerk2_mnu 
      Caption         =   "Szerk2"
      Visible         =   0   'False
      Begin VB.Menu rak_tul 
         Caption         =   "Raktár tulajdonságai"
         Enabled         =   0   'False
      End
      Begin VB.Menu v1 
         Caption         =   "-"
      End
      Begin VB.Menu enged2_mnu 
         Caption         =   "Szerkesztés engedélyezése"
      End
      Begin VB.Menu uj_tarolo 
         Caption         =   "Új tároló"
      End
      Begin VB.Menu v2 
         Caption         =   "-"
      End
      Begin VB.Menu ment 
         Caption         =   "Raktár állapotának mentése"
      End
   End
End
Attribute VB_Name = "raktar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sajat_vezerlo As Object
Public Konyvtaram As String
Dim dx As Single, dy As Single
Dim aktualis As Integer, meretezes As Boolean
Dim VisszaHivo As CommandButton
Dim szinkodok(1 To 8) As ColorConstants
Dim Beszuras As Boolean, Valasztott As String
Dim szerkesztes As Boolean


Private Sub ba_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        dx = x
        dy = Y
        ba.Drag
    End If
End Sub

Private Sub beilleszt_Click()
    kereso.Text = Clipboard.GetText
End Sub

Private Sub bf_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        dx = x
        dy = Y
        bf.Drag
    End If
End Sub

Private Sub enged2_mnu_Click()
    enged_mnu_Click
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
    Source.Visible = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If meretezes Then
        If KeyAscii = 13 Or KeyAscii = 27 Then
            meretezes = False
            ba.Visible = False
            ja.Visible = False
            bf.Visible = False
            jf.Visible = False
            keret.Visible = False
            tarolo(aktualis).Visible = True
            terulet.Cls
            
            ment_Click
        End If
    Else
        If KeyAscii = 27 Then
            kereso.Text = ""
        Else
            kereso.SetFocus
        End If
    End If
End Sub

'Megmutatja a megadott elemet
Private Sub AlapMutat(Mit As String)
    Form_Initialize
    
    Frissit
    kereso.Text = Mit
    kereso_Change
End Sub
Public Sub Mutasd(Optional Mit As String, Optional Hova As Byte)
    AlapMutat Mit
    
    Me.Show vbModal
End Sub
'Hely választás egy objektumnak
Public Sub HelyValaszt(Hivo As CommandButton)
On Error Resume Next
    AlapMutat Hivo.Caption
    Beszuras = True
    enged_mnu_Click
    szerkesztes = True
    Set VisszaHivo = Hivo
    
    Me.Show vbModal
End Sub

'Alapértelemzések
Private Sub Form_Initialize()
On Error Resume Next
    'Szinek betöltése
    AlapHelyzet
    Beszuras = False
    szerkesztes = False
    
    enged_mnu_Click
    Load szin(szin.Count)
    szin(szin.Count - 1).Caption = "Fekete"
    szin(szin.Count - 1).Checked = False
    szinkodok(szin.Count - 1) = vbBlack
    
    Load szin(szin.Count)
    szin(szin.Count - 1).Caption = "Kék"
    szin(szin.Count - 1).Checked = False
    szinkodok(szin.Count - 1) = vbBlue
    
    Load szin(szin.Count)
    szin(szin.Count - 1).Caption = "Cián"
    szin(szin.Count - 1).Checked = False
    szinkodok(szin.Count - 1) = vbCyan
    
    Load szin(szin.Count)
    szin(szin.Count - 1).Caption = "Zöld"
    szin(szin.Count - 1).Checked = False
    szinkodok(szin.Count - 1) = vbGreen
    
    Load szin(szin.Count)
    szin(szin.Count - 1).Caption = "Lila"
    szin(szin.Count - 1).Checked = False
    szinkodok(szin.Count - 1) = vbMagenta
    
    Load szin(szin.Count)
    szin(szin.Count - 1).Caption = "Piros"
    szin(szin.Count - 1).Checked = False
    szinkodok(szin.Count - 1) = vbRed
    
    Load szin(szin.Count)
    szin(szin.Count - 1).Caption = "Fehér"
    szin(szin.Count - 1).Checked = False
    szinkodok(szin.Count - 1) = vbWhite
    
    Load szin(szin.Count)
    szin(szin.Count - 1).Caption = "Sárga"
    szin(szin.Count - 1).Checked = False
    szinkodok(szin.Count - 1) = vbYellow

    szin(0).Visible = False
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    terulet_MouseMove Button, Shift, x, Y
End Sub

Private Sub Form_Resize()
    Pozicional
End Sub

Public Sub AlapHelyzet()
Dim i As Integer
    For i = 1 To tarolo.Count - 1
        Unload tarolo(i)
    Next i
End Sub
Public Sub Pozicional()
On Error Resume Next
Dim x As Single, Y As Single, pw As Single, ph As Single

If fl.Visible Then
        pw = fl.Width
    Else
        pw = 0
End If

If jb.Visible Then
        ph = jb.Height + vezerlok.Height
    Else
        ph = vezerlok.Height
End If


x = (ScaleWidth - terulet.Width - pw) / 2
Y = (ScaleHeight - terulet.Height - ph) / 2

vezerlok.Move 0, 0, Me.ScaleWidth, vezerlok.Height
fl.Move ScaleWidth - fl.Width, vezerlok.Height, fl.Width, ScaleHeight - fl.Width - vezerlok.Height
jb.Move 0, ScaleHeight - jb.Height, ScaleWidth - jb.Height, jb.Height
terulet.Move x, Y

If ScaleWidth - terulet.Width - pw < 0 Then
        jb.SmallChange = Int(ScaleWidth - terulet.Width / 100)
        jb.LargeChange = Int(ScaleWidth - terulet.Width / 10)
        jb.Max = ScaleWidth - terulet.Width - 2 * fl.Width
        jb.Min = fl.Width
        jb.Visible = True
        jb.Value = terulet.Left
    Else
        jb.Visible = False
End If

If ScaleHeight - terulet.Height - ph < 0 Then
        fl.SmallChange = Int(ScaleHeight - terulet.Height / 100)
        fl.LargeChange = Int(ScaleHeight - terulet.Height / 10)
        fl.Max = ScaleHeight - terulet.Height - 2 * jb.Height
        fl.Min = ph + fl.Width
        fl.Visible = True
        fl.Value = terulet.Top
    Else
        fl.Visible = False
End If

If (fl.Visible Or jb.Visible) Then
        gomb.Move fl.Left, jb.Top
        gomb.Visible = True
Else
        gomb.Visible = False
        
End If
'terulet.Top = fl.Value
'terulet.Left = jb.Value
End Sub
Private Sub fl_Change()
    terulet.Top = fl.Value
End Sub

Private Sub Form_Terminate()
    'ment_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ment_Click
End Sub

Private Sub ja_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        dx = x
        dy = Y
        ja.Drag
    End If
End Sub

Private Sub jb_Change()
    terulet.Left = jb.Value
End Sub

Private Sub enged_mnu_Click()
    enged_mnu.Checked = Not enged_mnu.Checked
    enged2_mnu.Checked = enged_mnu.Checked
    nev_mnu.Enabled = enged_mnu.Checked
    torol_mnu.Enabled = enged_mnu.Checked
    meretez_mnu.Enabled = enged_mnu.Checked
    uj_tarolo.Enabled = enged_mnu.Checked
    ment.Enabled = enged_mnu.Checked
    szinez_mnu.Enabled = enged_mnu.Checked
    klon_mnu.Enabled = enged_mnu.Checked
End Sub

Private Sub jf_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        dx = x
        dy = Y
        jf.Drag
    End If
End Sub

Private Sub kereso_Change()
Dim i As Integer, h As Byte
    kereso.Text = Trim(kereso.Text)
    'kereso.SelStart = Len(kereso.Text)
    h = Len(kereso.Text)
    'If h > 4 Then h = 4
    'If kereso.Text <> "" Then
        For i = 0 To tarolo.Count - 1
            'elem(i).Cls
            'felt: ++ (kereso.Text <> "") And
            If (Mid(UCase(tarolo(i).ToolTipText), 1, h) <> UCase(Trim(Mid(kereso.Text, 1, h)))) Then
                    tarolo(i).BackColor = vbButtonFace
                Else
                    tarolo(i).BackColor = tarolo(i).ForeColor
                    'elem(i).Print "*"
            End If
        Next i
    'End If
End Sub

Private Sub klon_mnu_Click()
    Dim valasz As String
    valasz = InputBox("Mi legyen a klónozott tároló neve?", "Új, klónozott tároló létrehozása", tarolo(aktualis).ToolTipText)
    If valasz = "" Then Exit Sub
    With tarolo(aktualis)
        uj_elem valasz, .Left + 100, .Top + 100, .ForeColor, .Width, .Height
    End With
    
    ment_Click
End Sub

Private Sub ment_Click()
Dim i As Integer
    FSQL "DELETE * FROM tarolo"
    For i = 1 To tarolo.Count - 1
        With tarolo(i)
            FSQL "INSERT INTO tarolo (nev, x, y, szel, mag, szin ) VALUES ('" & .ToolTipText & "', " & .Left & ", " & .Top & ", " & .Width & ", " & .Height & ", " & .ForeColor & ")"
        End With
    Next i
End Sub

Private Sub meretez_mnu_Click()
    meretezes = True
    ba.Visible = True
    ja.Visible = True
    bf.Visible = True
    jf.Visible = True
    keret.Visible = True
    tarolo(aktualis).Visible = False
    passzint (aktualis)
    
    ment_Click
End Sub

Private Sub nev_mnu_Click()
Dim eredeti As String, valasz As String
    eredeti = tarolo(aktualis).ToolTipText
    valasz = InputBox("Adja meg a tároló új nevét:", tarolo(aktualis).ToolTipText & " átnevezése", tarolo(aktualis).ToolTipText)
    If valasz = "" Then
        tarolo(aktualis).ToolTipText = eredeti
    Else
        tarolo(aktualis).ToolTipText = valasz
    End If
    
    ment_Click
End Sub

Private Sub szin_Click(Index As Integer)
    tarolo(aktualis).ForeColor = szinkodok(Index)
    tarolo(aktualis).BackColor = szinkodok(Index)
End Sub

Private Sub tarolo_Click(Index As Integer)
    aktualis = Index
    If Not enged_mnu.Checked Then
        kereso.Text = tarolo(Index).ToolTipText
    End If
End Sub

Private Sub tarolo_DblClick(Index As Integer)
    aktualis = Index
    kereso.Text = tarolo(Index).ToolTipText
    If Beszuras Then
        VisszaHivo.Caption = tarolo(Index).ToolTipText
        Unload Me
    End If
End Sub

Private Sub tarolo_DragOver(Index As Integer, Source As Control, x As Single, Y As Single, State As Integer)
    Form_DragOver Source, x, Y, State
End Sub

Private Sub tarolo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    aktualis = Index
    If Button = 1 And enged_mnu.Checked Then
        tarolo(Index).Visible = False
        dx = x
        dy = Y
        tarolo(Index).Drag
    End If
End Sub


Private Sub tarolo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    mutat.Caption = tarolo(Index).ToolTipText
End Sub

Private Sub tarolo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        aktualis = Index
        nev.Caption = tarolo(Index).ToolTipText & " tulajdonságai"
        PopupMenu szerk_mnu
    End If
End Sub

Private Sub terulet_Click()
    kereso.Text = ""
    kereso_Change
End Sub

Private Sub terulet_DragDrop(Source As Control, x As Single, Y As Single)
On Error GoTo kilep
    Source.Left = Racsra(x - dx)
    Source.Top = Racsra(Y - dy)
    Source.Visible = True
    With tarolo(aktualis)
    Select Case Source.Name
        Case "bf"
            .Move bf.Left, bf.Top, Racsra(jf.Left - bf.Left + jf.Width), Racsra(ba.Top - bf.Top + ba.Height)
            passzint (aktualis)
        Case "ba"
            .Move ba.Left, .Top, Racsra(.Left + .Width - ba.Left), Racsra(ba.Top + ba.Height - .Top)
            passzint (aktualis)
        Case "jf"
            .Move .Left, jf.Top, Racsra(jf.Left + jf.Width - .Left), Racsra(ja.Top + ja.Height - jf.Top)
            passzint (aktualis)
        Case "ja"
            .Move .Left, jf.Top, Racsra(ja.Left + ja.Width - .Left), Racsra(ja.Top + ja.Height - .Top)
            passzint (aktualis)
    End Select
    End With
    If tarolo(aktualis).Visible = False Then
        terulet.Cls
        terulet.CurrentX = tarolo(aktualis).Left + 100
        terulet.CurrentY = tarolo(aktualis).Top + 100
        terulet.Print tarolo(aktualis).Width & "x" & tarolo(aktualis).Height
    End If
kilep:
    passzint (aktualis)
    
    ment_Click
End Sub
Private Sub passzint(Id As Integer)
With tarolo(Id)
    bf.Move .Left, .Top
    ba.Move .Left, .Top + .Height - ja.Height
    jf.Move .Left + .Width - jf.Width, .Top
    ja.Move .Left + .Width - jf.Width, .Top + .Height - ja.Height
    keret.Move .Left, .Top, .Width, .Height
End With
End Sub

Private Sub terulet_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
    If Source.Name = "tarolo" Then Source.Visible = False
    If meretezes Then terulet_DragDrop Source, x, Y
End Sub

Private Sub terulet_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    mutat.Caption = "Mutasson rá valamelyik tárolóra!"
End Sub

Private Sub terulet_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 And szerkesztes Then
        dx = x
        dy = Y
        PopupMenu szerk2_mnu
    End If
End Sub
'Tároló törlése
Private Sub torol_mnu_Click()
 If MsgBox("Biztosan törölni akarja a(z) " & tarolo(aktualis).ToolTipText & " nevû tárolót?", vbQuestion + vbYesNo, "Törlés megerõsítése") = vbNo Then Exit Sub

        With tarolo(aktualis)
            .Left = tarolo(tarolo.Count - 1).Left
            .Top = tarolo(tarolo.Count - 1).Top
            .Width = tarolo(tarolo.Count - 1).Width
            .Height = tarolo(tarolo.Count - 1).Height
            .ToolTipText = tarolo(tarolo.Count - 1).ToolTipText
            .BackColor = tarolo(tarolo.Count - 1).ForeColor
            
            .ForeColor = tarolo(tarolo.Count - 1).ForeColor
        End With
    Unload tarolo(tarolo.Count - 1)
End Sub
'Új tároló
Private Sub uj_tarolo_Click()
Dim valasz As String
    valasz = InputBox("Mi legyen az új tároló neve?", "Új tároló létrehozása", "Tároló" & tarolo.Count)
    If valasz = "" Then Exit Sub
    uj_elem valasz, dx, dy, vbWhite, 200, 200
    
    ment_Click
End Sub

'Új elem léterhozása
Private Sub uj_elem(nev As String, x As Single, Y As Single, szine As ColorConstants, Szel As Single, Mag As Single)
    Dim Id As Integer
    Id = tarolo.Count
    Load tarolo(Id)
    With tarolo(Id)
        .ForeColor = szine
        .BackColor = szine
        .ToolTipText = nev
        .Left = x
        .Top = Y
        .Width = Szel
        .Height = Mag
        
        .Visible = True
    End With
End Sub
'Rácsra illesztés
Private Function Racsra(Szam)
Dim maradek
    maradek = Szam Mod 30
    Racsra = Szam - maradek
End Function

'Tárolók újratöltése
Private Sub Frissit()
    Dim Rek As String, Kulcsszo As String, Parameter As String
    Dim Sor As New ADODB.Recordset
    
    'Minden elem törlése
    AlapHelyzet
    
    'Kép betöltése
    terulet.Picture = LoadPicture(Konyvtar & "Egyeb\raktar.gif")
    
    SQL_p "SELECT * FROM tarolo", Sor
    If Not Sor.EOF Then Sor.MoveFirst
        Do While Not Sor.EOF
            uj_elem Nstr(Sor!nev), Sor!x, Sor!Y, Sor!szin, Sor!Szel, Sor!Mag
            Sor.MoveNext
        Loop
    Sor.Close
End Sub

