VERSION 5.00
Begin VB.Form frmFTPTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manual ... FTP"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLocalIP 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   270
      Width           =   2685
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   2700
      TabIndex        =   4
      Text            =   "1027"
      Top             =   270
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Ftp!"
      Height          =   375
      Left            =   2700
      TabIndex        =   1
      Top             =   630
      Width           =   1095
   End
   Begin VB.TextBox txtlog 
      Height          =   2205
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1350
      Width           =   3795
   End
   Begin VB.Label Label5 
      Caption         =   "Sorry I have to do more code cleen up!!!!"
      Height          =   225
      Left            =   270
      TabIndex        =   8
      Top             =   4920
      Width           =   3075
   End
   Begin VB.Label Label4 
      Caption         =   "The intel FILE is an exe so Just Rename the file tmp.tmp to tmp.exe and execute it to test Download"
      Height          =   465
      Left            =   0
      TabIndex        =   7
      Top             =   4230
      Width           =   3795
   End
   Begin VB.Label Label2 
      Caption         =   "Local IP Info (DUN):"
      Height          =   225
      Index           =   1
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Status!"
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   1050
      Width           =   3765
   End
   Begin VB.Label Label3 
      Caption         =   "This Test Will Download a file FROM: ftp.intel.com TO file tmp.tmp in the App Folder"
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   3795
   End
End
Attribute VB_Name = "frmFTPTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents FormSocket As cFormSocket
Attribute FormSocket.VB_VarHelpID = -1
Private Sub Command1_Click()
    Dim b As New cFormSocket
    b.sckStart
    b.sckListen txtPort
    
        FormSocket.sckConnect "ftp.intel.com", "21"
 
    a = FormSocket.sckSendCommand("USER anonymous", "331")
    a = FormSocket.sckSendCommand("PASS test@test.org", "230")
    a = FormSocket.sckSendCommand("CWD /pub/benchmark/bapco/", "250")
    a = FormSocket.sckSendCommand("TYPE I", "200")
    a = FormSocket.sckSendCommand("PORT " & BuildPortString(txtLocalIP, txtPort), "200")
    a = FormSocket.sckSendCommand("RETR SQLSYS.EXE", "226")

    b.sckStop

    Set b = Nothing
End Sub
Private Sub Form_Load()

 txtLocalIP = GetIPAddress()
                
 Set FormSocket = New cFormSocket
 FormSocket.sckHook Me.hWnd

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
    If Not FormSocket Is Nothing Then
        
        FormSocket.sckClose
        FormSocket.sckUnHook
        Set FormSocket = Nothing
    
    End If

End Sub

Private Function BuildPortString(pIP As String, pPort As String)
Dim pDummy As Variant
Dim pTmp As String

pDummy = Split(pIP, ".")

BuildPortString = pDummy(0) & "," _
                & pDummy(1) & "," _
                & pDummy(2) & "," _
                & pDummy(3) & "," _
                & pPort \ 256 & "," _
                & pPort - (pPort \ 256) * 256


End Function

Private Sub FormSocket_sckLog(pLogEntry As String)
    txtlog = txtlog & pLogEntry & vbCrLf
    txtlog.SelStart = Len(txtlog.Text)
End Sub

