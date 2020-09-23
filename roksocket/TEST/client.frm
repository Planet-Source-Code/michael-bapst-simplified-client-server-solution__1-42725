VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2172
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3732
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents RokSock As RokSocket
Attribute RokSock.VB_VarHelpID = -1

Private Sub Form_Load()
Set RokSock = New RokSocket
With RokSock
    .Connect "grimnir", 1001
End With

End Sub


Private Sub Label1_Click()
RokSock.SendData "This is a test"

RokSock.CloseSocket
Set RokSock = Nothing
End

End Sub

Private Sub RokSock_DataArrival(ByVal bytesTotal As Long)
Dim TmpStr As String
RokSock.GetData TmpStr
Label1.Caption = TmpStr
End Sub

