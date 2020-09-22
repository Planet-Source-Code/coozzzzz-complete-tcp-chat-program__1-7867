VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_CreateAccount 
   Caption         =   "Create New Account"
   ClientHeight    =   2325
   ClientLeft      =   5670
   ClientTop       =   4875
   ClientWidth     =   3990
   Icon            =   "frm_CreateAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "C&hange"
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3000
      Top             =   0
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Create"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   3
      Text            =   "Password"
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frm_CreateAccount.frx":030A
      Top             =   960
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Status :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Connection Closed."
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label14 
      Caption         =   "Server Hostname / IP Address :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label12 
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "Username  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frm_CreateAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connection As Boolean, newa As Boolean, changea As Boolean, averified As Boolean
Private Sub Command3_Click()

If averified = False Then
newa = False: changea = True
If Winsock1.State <> 0 Then Winsock1.Close
Winsock1.Connect Text3, "12340"
Else
Winsock1.SendData "@chng" & Text1 & "@password" & Text2
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
SetVals
End Sub

Private Sub Text1_Change()
txt$ = Text1
Text1 = frm_NewServer.StringChange(txt$)
Text1.SelStart = Len(Text1)
End Sub

Private Sub Text2_Change()
txt$ = Text2
Text2 = frm_NewServer.StringChange(txt$)
Text2.SelStart = Len(Text2)
End Sub

Private Sub winsock1_Connect()

If newa = True Then
data$ = "@name" & Text1 & "@password" & Text2
Else
data$ = "@came" & Text1 & "@password" & Text2
End If

Winsock1.SendData data$
End Sub
Private Sub Command1_Click()

If Winsock1.State <> 0 Then Winsock1.Close
newa = True: changea = False
Winsock1.Connect Text3, "12340"

End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()

If Winsock1.State = 7 And connection = False Then connection = True: Label1.Caption = "Connection Established."

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim data$
Winsock1.GetData data$
If data$ = "dupe" Then Winsock1.Close: MsgBox "Name already exists in database.", vbCritical: Label1.Caption = "Connection Closed.": SetVals
If data$ = "success" Then Winsock1.Close: MsgBox "Account Created.", vbInformation: Label1.Caption = "Connection Closed.": Unload Me
If data$ = "nov" Then Winsock1.Close: MsgBox "Username / Password incorrect.", vbInformation: Label1.Caption = "Connection Closed.": SetVals
If data$ = "v" Then averified = True: MsgBox "Account Verified. Enter new password and click change.", vbInformation: Text1.Enabled = False
If data$ = "@changed" Then Winsock1.Close: MsgBox "Password changed.", vbInformation: Unload Me
If data$ = "nc" Then Winsock1.Close: MsgBox "Error. Password not changed.", vbInformation: Unload Me

End Sub
Sub SetVals()
connection = False: newa = False: changea = False: averified = False
End Sub
