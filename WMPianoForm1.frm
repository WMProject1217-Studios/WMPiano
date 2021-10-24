VERSION 5.00
Begin VB.Form WMPianoForm1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   12585
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   12
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   11
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "WMPianoForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Const SND_ASYNC = &H1

Public Sub PianoKey(KeyAscii As Integer)
Me.Caption = KeyAscii
If KeyAscii = 113 Then
PlaySound App.Path & "\Audio\C.wav", 0, SND_ASYNC
End If
If KeyAscii = 119 Then
PlaySound App.Path & "\Audio\D.wav", 0, SND_ASYNC
End If
If KeyAscii = 101 Then
PlaySound App.Path & "\Audio\E.wav", 0, SND_ASYNC
End If
If KeyAscii = 114 Then
PlaySound App.Path & "\Audio\F.wav", 0, SND_ASYNC
End If
If KeyAscii = 116 Then
PlaySound App.Path & "\Audio\G.wav", 0, SND_ASYNC
End If
If KeyAscii = 121 Then
PlaySound App.Path & "\Audio\A.wav", 0, SND_ASYNC
End If
If KeyAscii = 117 Then
PlaySound App.Path & "\Audio\B.wav", 0, SND_ASYNC
End If
If KeyAscii = 97 Then
PlaySound App.Path & "\Audio\C##.wav", 0, SND_ASYNC
End If
If KeyAscii = 115 Then
PlaySound App.Path & "\Audio\D#.wav", 0, SND_ASYNC
End If
If KeyAscii = 100 Then
PlaySound App.Path & "\Audio\F#.wav", 0, SND_ASYNC
End If
If KeyAscii = 102 Then
PlaySound App.Path & "\Audio\G#.wav", 0, SND_ASYNC
End If
If KeyAscii = 103 Then
PlaySound App.Path & "\Audio\A#.wav", 0, SND_ASYNC
End If
If KeyAscii = 104 Then
PlaySound App.Path & "\Audio\C#.wav", 0, SND_ASYNC
End If
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
PianoKey (KeyAscii)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
PianoKey (KeyAscii)
End Sub

Private Sub Label1_Click()
PianoKey (113)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackColor = RGB(102, 204, 255)
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackColor = RGB(240, 240, 240)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = RGB(102, 204, 255)
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = RGB(240, 240, 240)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackColor = RGB(102, 204, 255)
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackColor = RGB(240, 240, 240)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BackColor = RGB(102, 204, 255)
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BackColor = RGB(240, 240, 240)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = RGB(102, 204, 255)
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = RGB(240, 240, 240)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BackColor = RGB(102, 204, 255)
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BackColor = RGB(240, 240, 240)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BackColor = RGB(102, 204, 255)
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BackColor = RGB(240, 240, 240)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackColor = RGB(102, 204, 255)
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackColor = RGB(240, 240, 240)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BackColor = RGB(102, 204, 255)
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BackColor = RGB(240, 240, 240)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.BackColor = RGB(102, 204, 255)
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.BackColor = RGB(240, 240, 240)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.BackColor = RGB(102, 204, 255)
End Sub

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.BackColor = RGB(240, 240, 240)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.BackColor = RGB(102, 204, 255)
End Sub

Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.BackColor = RGB(240, 240, 240)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.BackColor = RGB(102, 204, 255)
End Sub

Private Sub Label13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.BackColor = RGB(240, 240, 240)
End Sub

Private Sub Label2_Click()
PianoKey (119)
End Sub
Private Sub Label3_Click()
PianoKey (101)
End Sub
Private Sub Label4_Click()
PianoKey (114)
End Sub
Private Sub Label5_Click()
PianoKey (116)
End Sub
Private Sub Label6_Click()
PianoKey (121)
End Sub
Private Sub Label7_Click()
PianoKey (117)
End Sub
Private Sub Label8_Click()
PianoKey (97)
End Sub
Private Sub Label9_Click()
PianoKey (115)
End Sub
Private Sub Label10_Click()
PianoKey (100)
End Sub
Private Sub Label11_Click()
PianoKey (102)
End Sub
Private Sub Label12_Click()
PianoKey (103)
End Sub
Private Sub Label13_Click()
PianoKey (104)
End Sub
