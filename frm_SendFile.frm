VERSION 5.00
Begin VB.Form frm_SendFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Getting File..."
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "frm_SendFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Bytes Remaining :"
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frm_SendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
