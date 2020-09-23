VERSION 5.00
Begin VB.MDIForm frmMDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Socks!"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8055
      TabIndex        =   0
      Top             =   0
      Width           =   8115
      Begin VB.CommandButton Command3 
         Caption         =   "Manual... FTP Test!"
         Height          =   345
         Left            =   3840
         TabIndex        =   3
         Top             =   60
         Width           =   1605
      End
      Begin VB.CommandButton Command2 
         Caption         =   "New Listener"
         Height          =   345
         Left            =   1980
         TabIndex        =   2
         Top             =   60
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "New Connecter"
         Height          =   345
         Left            =   30
         TabIndex        =   1
         Top             =   60
         Width           =   1905
      End
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SocketForm As frmSocket
Dim ListenForm As frmListen
Private Sub Command1_Click()
Set SocketForm = New frmSocket
    SocketForm.Show

End Sub

Private Sub Command2_Click()


Set ListenForm = New frmListen
    ListenForm.Show

End Sub

Private Sub Command3_Click()

frmFTPTest.Show

End Sub

Private Sub MDIForm_Load()
    Call StartWinsock("")
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Set SocketForm = Nothing
'...
Unload frmSocket
   
Call WSACleanup

End Sub
