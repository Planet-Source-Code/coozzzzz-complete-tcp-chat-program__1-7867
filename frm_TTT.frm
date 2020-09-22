VERSION 5.00
Begin VB.Form frm_TTT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic-Tac-Toe  - Playing..."
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   Icon            =   "frm_TTT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   2160
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   2640
      Top             =   1800
   End
   Begin VB.PictureBox xPic 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   360
      Picture         =   "frm_TTT.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   2400
      Width           =   540
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   1680
      TabIndex        =   14
      Top             =   120
      Width           =   615
      Begin VB.Label Spot 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         MouseIcon       =   "frm_TTT.frx":0884
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.PictureBox oPic 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   960
      Picture         =   "frm_TTT.frx":0CC6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   2400
      Width           =   540
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   960
      TabIndex        =   11
      Top             =   120
      Width           =   615
      Begin VB.Label Spot 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         MouseIcon       =   "frm_TTT.frx":1108
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   615
      Begin VB.Label Spot 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         MouseIcon       =   "frm_TTT.frx":154A
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   1680
      TabIndex        =   9
      Top             =   720
      Width           =   615
      Begin VB.Label Spot 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         MouseIcon       =   "frm_TTT.frx":198C
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   960
      TabIndex        =   8
      Top             =   720
      Width           =   615
      Begin VB.Label Spot 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         MouseIcon       =   "frm_TTT.frx":1DCE
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   615
      Begin VB.Label Spot 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         MouseIcon       =   "frm_TTT.frx":2210
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame12 
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   360
      Width           =   1335
      Begin VB.Label Score1 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame11 
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
      Begin VB.Label Score2 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame9 
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   615
      Begin VB.Label Spot 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         MouseIcon       =   "frm_TTT.frx":2652
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame8 
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   615
      Begin VB.Label Spot 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         MouseIcon       =   "frm_TTT.frx":2A94
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   615
      Begin VB.Label Spot 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         MouseIcon       =   "frm_TTT.frx":2ED6
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_TTT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

If player1TTT = True Then

userX = True
Me.MouseIcon = xPic.Picture

For i = 0 To Spot.Count - 1
Spot(i).MouseIcon = xPic.Picture
Spot(i).Caption = ""
spots(i) = True
Next i

ElseIf player2TTT = True Then

userX = False
Me.MouseIcon = oPic.Picture

For i = 0 To Spot.Count - 1
Spot(i).MouseIcon = oPic.Picture
Spot(i).Caption = ""
spots(i) = True
Next i

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

If player1TTT = True Then
frm_NewServer.Winsock1.SendData "@tttend" & ttt2
Else
frm_NewServer.Winsock1.SendData "@tttend" & ttt1
End If


ttt1 = "": ttt2 = ""
player1TTT = False
player2TTT = False
userX = False
played = True

For i = 0 To 8
spots(i) = False
Next i

Unload Me

End Sub

Private Sub Spot_Click(Index As Integer)

If spots(Index) = False Then Exit Sub
If played = True Then Exit Sub

If player1TTT = True Then
Spot(Index).Caption = "X"
spots(Index) = False
frm_NewServer.Winsock1.SendData "@tttclick" & ttt2 & "@spot" & Index: Pause 0.1
played = True

Else
Spot(Index).Caption = "O"
spots(Index) = False
frm_NewServer.Winsock1.SendData "@tttclick" & ttt1 & "@spot" & Index: Pause 0.1
played = True
End If

End Sub

Sub CheckSpots(who As String)
who = UCase(who)

'Check Right-Left
If Spot(0) = who And Spot(1) = who And Spot(2) = who Then GoTo checkname
If Spot(3) = who And Spot(4) = who And Spot(5) = who Then GoTo checkname
If Spot(6) = who And Spot(7) = who And Spot(8) = who Then GoTo checkname

'Check Up-Bottom
If Spot(0) = who And Spot(3) = who And Spot(6) = who Then GoTo checkname
If Spot(1) = who And Spot(4) = who And Spot(7) = who Then GoTo checkname
If Spot(2) = who And Spot(5) = who And Spot(8) = who Then GoTo checkname

'Check Other
If Spot(0) = who And Spot(4) = who And Spot(8) = who Then GoTo checkname
If Spot(2) = who And Spot(4) = who And Spot(6) = who Then GoTo checkname



If spots(0) = False And spots(1) = False And spots(2) = False And spots(3) = False And spots(4) = False And spots(5) = False And spots(7) = False And spots(8) = False Then
MsgBox "Nobody has won!", vbInformation
Form_Load
If player1TTT = True Then
played = True
Else
played = False
End If
lRet = ""
End If

Exit Sub

checkname:
If who = "X" Then
Score1 = Score1 + 1
played = False
MsgBox Frame12.Caption & " has Won!", vbInformation
Form_Load
Else

Score2 = Score2 + 1
played = True
MsgBox Frame11.Caption & " has Won!", vbInformation
Form_Load
End If

End Sub


Private Sub Timer1_Timer()
CheckSpots "x"
CheckSpots "o"
End Sub
