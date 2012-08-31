VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Мальчик"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "закончить"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "начать"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar info 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   7
      Top             =   7935
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "Информационная панель"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label vzprikup 
      Caption         =   "Прикуп"
      Height          =   255
      Left            =   8400
      TabIndex        =   8
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Image prikup 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   3
      Left            =   9840
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   975
   End
   Begin VB.Image prikup 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   2
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   975
   End
   Begin VB.Image prikup 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   1
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   975
   End
   Begin VB.Image prikup 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   0
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   975
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Index           =   4
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Index           =   3
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Index           =   2
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Index           =   1
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   4
      Left            =   10440
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   3
      Left            =   9960
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   2
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   1
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   975
   End
   Begin VB.Image karta3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   0
      Left            =   8520
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   975
   End
   Begin VB.Shape kar3 
      Height          =   2175
      Left            =   8400
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   4
      Left            =   10440
      Stretch         =   -1  'True
      Top             =   360
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   3
      Left            =   9960
      Stretch         =   -1  'True
      Top             =   480
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   2
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   600
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   1
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   975
   End
   Begin VB.Image karta2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   0
      Left            =   8520
      Stretch         =   -1  'True
      Top             =   840
      Width           =   975
   End
   Begin VB.Shape kar2 
      Height          =   2175
      Left            =   8400
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image karta1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Index           =   0
      Left            =   120
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Shape kar1 
      Height          =   2175
      Left            =   0
      Top             =   5760
      Width           =   6735
   End
   Begin VB.Image stol 
      BorderStyle     =   1  'Fixed Single
      Height          =   4575
      Left            =   1920
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label slova3 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label slova2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   240
      Width           =   1935
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   7560
      TabIndex        =   3
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   7560
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Козырь"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image kozir 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   240
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1455
   End
   Begin VB.Menu m_file 
      Caption         =   "&Файл"
      Begin VB.Menu start 
         Caption         =   "Начать"
      End
      Begin VB.Menu end 
         Caption         =   "Выход"
      End
   End
   Begin VB.Menu m_nastr 
      Caption         =   "&Настройки"
   End
   Begin VB.Menu m_spravka 
      Caption         =   "&Помощь"
      Begin VB.Menu m_help 
         Caption         =   "Справка"
      End
      Begin VB.Menu m_info 
         Caption         =   "О программе"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ind As Integer
Dim nhod As Integer
Dim chprikup As Integer
Dim vibr1 As Integer
Dim vibr2 As Integer
Dim vibr3 As Integer
Dim nk As Integer
Dim nvibkarti As Integer
Dim nkx As Integer
Dim kolvo As Integer
Dim n As Integer
Dim chhod As Integer
Dim tip As Integer

Private Sub Command1_Click()
Call start_Click
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub end_Click()
If MsgBox("Вы действительно хотите выйти?", vbYesNo, "Выход") = vbYes Then End
End Sub

Private Sub Form_Click()
If vibr1 = -1 Or start.Enabled = True Then GoTo en
ki = 0
ks = 0
If slova2.Caption = "Играю" Then ki = ki + 1
If slova3.Caption = "Играю" Then ki = ki + 1
If kar1.Tag = "y" Then ki = ki + 1
If vibr2 <> -1 Then ks = ks + 1
If vibr3 <> -1 Then ks = ks + 1
If vibr1 <> -1 Then ks = ks + 1
If ki = ks Then
nhod = nhod + 1
stol.Tag = ""
End If
If nhod = 5 Then
MsgBox "Сходили"
nhod = 0
End If
If vibr1 <> -1 Then karta1(vibr1).Move 0, 0
If vibr2 <> -1 Then karta2(vibr2).Move 100, 100
If vibr3 <> -1 Then karta3(vibr3).Move 200, 200
vibr1 = -1
vibr2 = -1
vibr3 = -1
en:
Call games
End Sub

Private Sub Form_Load()
Call razdrub
For ind = 0 To 4
karta1(ind).MouseIcon = LoadPicture(App.Path + "\h_point.cur")
Next ind
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Вы действительно хотите выйти?", vbYesNo, "Выход") = vbNo Then
Cancel = 1
End If
End Sub


Private Sub Image1_Click()

End Sub

Private Sub karta1_Click(Index As Integer)
If vibr1 <> -1 Then GoTo en
vibr1 = Index
karta1(Index).Move 20 + stol.Left + (karta1(Index).Width - 20) * Rnd, 20 + stol.Top + (karta1(Index).Height - 20) * Rnd
karta1(Index).Enabled = False
karta1(Index).ZOrder (0)
stol.Tag = stol.Tag + "1" + Mid(karta1(vibr1).Tag, 1, 2)
chhod = chhod + 1
If chhod = 4 Then chhod = 1
If chhod = 2 Then Call games
If chhod = 3 Then Call games
en:
End Sub

Private Sub karta1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
karta1(Index).MousePointer = 99
End Sub

Private Sub start_Click()
start.Enabled = False
For n = 0 To 4
karta1(n).Move 120 + n * 1320, 5880
karta1(n).ZOrder (0)
karta1(n).Enabled = True
karta2(n).Move 8520 + 480 * n, 840 - 120 * n, prikup(0).Width, prikup(0).Height
karta2(n).ZOrder (0)
karta2(n).Enabled = True
karta3(n).Move 8520 + 480 * n, 3240 - 120 * n, prikup(0).Width, prikup(0).Height
karta3(n).ZOrder (0)
karta3(n).Enabled = True
Next n
vibr1 = -1
vibr2 = -1
vibr3 = -1
stol.ZOrder (0)
slova2.Tag = ""
slova3.Tag = ""
kar1.Tag = ""
info.SimpleText = ""
start.Enabled = False
Randomize
chprikup = Int(Rnd * 3) + 1 'чей прикуп 1-3
If chprikup = 1 Then
vzprikup.Enabled = True
Else
vzprikup.Enabled = False
End If
chhod = chprikup
slova2.Caption = ""
slova3.Caption = ""
Call razd 'раздача карт
Call editkart
Call analiz
Call games
End Sub

Public Sub razd()
nhod = 0
nk = 36
karti = "6б6в6к6ч7б7в7к7ч8б8в8к8ч9б9в9к9чабавакачвбвввквчдбдвдкдчкбквкккчтбтвтктч"
For nkx = 0 To 4
Randomize
nvibkarti = Int(Rnd * nk + 1) * 2 - 1
karta1(nkx).Tag = Mid(karti, nvibkarti, 2)
karta1(nkx).Picture = LoadPicture(App.Path + "\колода\" + Mid(karti, nvibkarti, 2) + ".bmp")
karti = Left(karti, nvibkarti - 1) + Right(karti, Len(karti) - (nvibkarti + 1))
nk = nk - 1
Randomize
nvibkarti = Int(Rnd * nk + 1) * 2 - 1
karta2(nkx).Tag = Mid(karti, nvibkarti, 2)
 karta2(nkx).Picture = LoadPicture(App.Path + "\колода\" + Mid(karti, nvibkarti, 2) + ".bmp")
karti = Left(karti, nvibkarti - 1) + Right(karti, Len(karti) - (nvibkarti + 1))
nk = nk - 1
Randomize
nvibkarti = Int(Rnd * nk + 1) * 2 - 1
karta3(nkx).Tag = Mid(karti, nvibkarti, 2)
 karta3(nkx).Picture = LoadPicture(App.Path + "\колода\" + Mid(karti, nvibkarti, 2) + ".bmp")
karti = Left(karti, nvibkarti - 1) + Right(karti, Len(karti) - (nvibkarti + 1))
nk = nk - 1
If nkx = 4 Then
Randomize
nvibkarti = Int(Rnd * nk + 1) * 2 - 1
kozir.Tag = Mid(karti, nvibkarti, 2)
kozir.Picture = LoadPicture(App.Path + "\колода\" + Mid(karti, nvibkarti, 2) + ".bmp")
karti = Left(karti, nvibkarti - 1) + Right(karti, Len(karti) - (nvibkarti + 1))
nk = nk - 1
Else
Randomize
nvibkarti = Int(Rnd * nk + 1) * 2 - 1
prikup(nkx).Tag = Mid(karti, nvibkarti, 2)
 prikup(nkx).Picture = LoadPicture(App.Path + "\колода\" + Mid(karti, nvibkarti, 2) + ".bmp")
karti = Left(karti, nvibkarti - 1) + Right(karti, Len(karti) - (nvibkarti + 1))
nk = nk - 1
End If
Next nkx
slova2.Caption = ""
slova3.Caption = ""
Select Case chprikup
Case 1
info.SimpleText = "Твой прикуп"
Case 2
info.SimpleText = "Прикуп второго игрока"
Case 3
info.SimpleText = "Прикуп третьего игрока"
End Select
End Sub

Public Sub razdrub()
filerub = "b01_c.bmp" 'временная строка, имя файла текущей рубашки загружать из спец.файла из директории с рубашками
For nkx = 0 To 4
karta1(nkx).Picture = LoadPicture(App.Path + "\рубашка\" + filerub)
karta2(nkx).Picture = LoadPicture(App.Path + "\рубашка\" + filerub)
karta3(nkx).Picture = LoadPicture(App.Path + "\рубашка\" + filerub)
If nkx = 4 Then
kozir.Picture = LoadPicture(App.Path + "\рубашка\" + filerub)
Else
prikup(nkx).Picture = LoadPicture(App.Path + "\рубашка\" + filerub)
End If
Next nkx
End Sub

Public Sub analiz()
Exit Sub
kar2.Tag = 0
kar3.Tag = 0
For n = 0 To 4
bubi2 = ""
vini2 = ""
kresti2 = ""
chervi2 = ""
koziri2 = ""
bubi3 = ""
vini3 = ""
kresti3 = ""
chervi3 = ""
koziri3 = ""

If Mid(karta2(n).Tag, 2, 1) = "а" Then
'tip = 9
'tip = 18
tip = 27
Else
tip = 0
End If
 If Mid(karta2(n).Tag, 1, 2) = "тa" Or (Mid(karta2(n).Tag, 1, 2) = "ка" And kozir.Tag = "та") Then
 tip = 100
 End If
Select Case Mid(karta2(n).Tag, 1, 1)
Case "6"
kar2.Tag = Val(kar2.Tag) + tip
karta2(n).Tag = karta2(n).Tag + Right(Str(tip), 2)
Case "7"
kar2.Tag = Val(kar2.Tag) + tip + 1
karta2(n).Tag = karta2(n).Tag + Right(Str(tip + 1), 2)
Case "8"
kar2.Tag = Val(kar2.Tag) + tip + 2
karta2(n).Tag = karta2(n).Tag + Right(Str(tip + 2), 2)
Case "9"
kar2.Tag = Val(kar2.Tag) + tip + 3
karta2(n).Tag = karta2(n).Tag + Right(Str(tip + 3), 2)
Case "а"
kar2.Tag = Val(kar2.Tag) + tip + 4
karta2(n).Tag = karta2(n).Tag + Right(Str(tip + 4), 2)
Case "в"
kar2.Tag = Val(kar2.Tag) + tip + 5
karta2(n).Tag = karta2(n).Tag + Right(Str(tip + 5), 2)
Case "д"
kar2.Tag = Val(kar2.Tag) + tip + 6
karta2(n).Tag = karta2(n).Tag + Right(Str(tip + 6), 2)
Case "к"
kar2.Tag = Val(kar2.Tag) + tip + 7
karta2(n).Tag = karta2(n).Tag + Right(Str(tip + 7), 2)
Case "т"
kar2.Tag = Val(kar2.Tag) + tip + 8
karta2(n).Tag = karta2(n).Tag + Right(Str(tip + 8), 2)
End Select

If Right(karta3(n).Tag, 1) = "а" Then
tip = 27
Else
tip = 0
End If
 If karta2(n).Tag = "тa" Or (karta2(n).Tag = "ка" And kozir.Tag = "та") Then
 tip = 100
 End If
Select Case Left(karta3(n).Tag, 1)
Case "6"
kar3.Tag = Val(kar3.Tag) + tip
karta3(n).Tag = karta3(n).Tag + Right(Str(tip), 2)
Case "7"
kar3.Tag = Val(kar3.Tag) + tip + 1
karta3(n).Tag = karta3(n).Tag + Right(Str(tip + 1), 2)
Case "8"
kar3.Tag = Val(kar3.Tag) + tip + 2
karta3(n).Tag = karta3(n).Tag + Right(Str(tip + 2), 2)
Case "9"
kar3.Tag = Val(kar3.Tag) + tip + 3
karta3(n).Tag = karta3(n).Tag + Right(Str(tip + 3), 2)
Case "а"
kar3.Tag = Val(kar3.Tag) + tip + 4
karta3(n).Tag = karta3(n).Tag + Right(Str(tip + 4), 2)
Case "в"
kar3.Tag = Val(kar3.Tag) + tip + 5
karta3(n).Tag = karta3(n).Tag + Right(Str(tip + 5), 2)
Case "д"
kar3.Tag = Val(kar3.Tag) + tip + 6
karta3(n).Tag = karta3(n).Tag + Right(Str(tip + 6), 2)
Case "к"
kar3.Tag = Val(kar3.Tag) + tip + 7
karta3(n).Tag = karta3(n).Tag + Right(Str(tip + 7), 2)
Case "т"
kar3.Tag = Val(kar3.Tag) + tip + 8
karta3(n).Tag = karta3(n).Tag + Right(Str(tip + 8), 2)
End Select

Select Case karta2(n).Tag
Case "б"
bubi2 = bubi2 + Trim(Str(n)) + Mid(karta2(n).Tag, 1, 1)
Case "в"
vini2 = vini2 + Trim(Str(n)) + Mid(karta2(n).Tag, 1, 1)
Case "к"
kresti2 = kresti2 + Trim(Str(n)) + Mid(karta2(n).Tag, 1, 1)
Case "ч"
chervi2 = chervi2 + Trim(Str(n)) + Mid(karta2(n).Tag, 1, 1)
Case "а"
koziri2 = koziri2 + Trim(Str(n)) + Mid(karta2(n).Tag, 1, 1)
End Select

Select Case karta3(n).Tag
Case "б"
bubi3 = bubi3 + Trim(Str(n)) + Mid(karta3(n).Tag, 1, 1)
Case "в"
vini3 = vini3 + Trim(Str(n)) + Mid(karta3(n).Tag, 1, 1)
Case "к"
kresti3 = kresti3 + Trim(Str(n)) + Mid(karta3(n).Tag, 1, 1)
Case "ч"
chervi3 = chervi3 + Trim(Str(n)) + Mid(karta3(n).Tag, 1, 1)
Case "а"
koziri3 = koziri3 + Trim(Str(n)) + Mid(karta3(n).Tag, 1, 1)
End Select

Next n
 
 Label8.Caption = kar2.Tag
 Label9.Caption = kar3.Tag
 
End Sub

Public Sub editkart()
mastkozir = Right(kozir.Tag, 1)
For n = 0 To 4
If Mid(karta1(n).Tag, 2, 1) = mastkozir Then karta1(n).Tag = Mid(karta1(n).Tag, 1, 1) + "а" + Mid(karta1(n).Tag, 3, 2)
If Mid(karta2(n).Tag, 2, 1) = mastkozir Then karta2(n).Tag = Mid(karta2(n).Tag, 1, 1) + "а" + Mid(karta2(n).Tag, 3, 2)
If Mid(karta3(n).Tag, 2, 1) = mastkozir Then karta3(n).Tag = Mid(karta3(n).Tag, 1, 1) + "а" + Mid(karta3(n).Tag, 3, 2)
If n < 4 Then
If Mid(prikup(n).Tag, 2, 1) = mastkozir Then prikup(n).Tag = Mid(prikup(n).Tag, 1, 1) + "а"
End If
Next n
End Sub

Public Sub hod2()
If vibr2 <> -1 Then GoTo en
nach:
Call prvibr2
If karta2(vibr2).Enabled = False Then GoTo nach
karta2(vibr2).Move 20 + stol.Left + (karta2(vibr2).Width - 20) * Rnd, 20 + stol.Top + (karta2(vibr2).Height - 20) * Rnd, kozir.Width, kozir.Height
karta2(vibr2).Enabled = False
karta2(vibr2).ZOrder (0)
stol.Tag = stol.Tag + "2" + Mid(karta2(vibr2).Tag, 1, 2)
en:
chhod = chhod + 1
If chhod = 4 Then chhod = 1
Call games
End Sub

Public Sub hod3()
If vibr3 <> -1 Then GoTo en
nach:
Call prvibr3
If karta3(vibr3).Enabled = False Then GoTo nach
karta3(vibr3).Move 20 + stol.Left + (karta3(vibr3).Width - 20) * Rnd, 20 + stol.Top + (karta3(vibr3).Height - 20) * Rnd, kozir.Width, kozir.Height
karta3(vibr3).Enabled = False
karta3(vibr3).ZOrder (0)
stol.Tag = stol.Tag + "3" + Mid(karta3(vibr3).Tag, 1, 2)
en:
chhod = chhod + 1
If chhod = 4 Then chhod = 1
End Sub

Public Sub games()
nach:
Select Case chhod
Case 2
If slova2.Caption = "Играю" Then
If nhod < 5 Then Call hod2
Else
chhod = chhod + 1
If chhod = 4 Then chhod = 1
GoTo nach
End If
Case 3
If slova3.Caption = "Играю" Then
If nhod < 5 Then Call hod3
Else
chhod = chhod + 1
If chhod = 4 Then chhod = 1
GoTo nach
End If
End Select
End Sub

Public Sub prvibr2()

End Sub

Public Sub prvibr3()

End Sub

