VERSION 5.00
Begin VB.Form frmSocket 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Socket"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Log"
      Height          =   3375
      Left            =   3240
      TabIndex        =   14
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtLog 
         Height          =   2925
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   3015
      Begin VB.CommandButton Command3 
         Caption         =   "Macro"
         Height          =   375
         Left            =   810
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtWaitFor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "SEND >>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtCMD 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Text            =   "USER paolo"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "CMD Status"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Wait For"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "CMD:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtLastAnswer 
      Height          =   1395
      Left            =   0
      TabIndex        =   9
      Top             =   7320
      Width           =   10065
   End
   Begin VB.TextBox txtLastCommand 
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "Connect"
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Text            =   "21"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtHost 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Text            =   "192.168.0.96"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Host"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   7200
      Width           =   5415
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data:"
      Height          =   3375
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   9015
      Begin VB.TextBox txtData 
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Last CMD"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   5160
      Width           =   1545
   End
End
Attribute VB_Name = "frmSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents FormSocket As cFormSocket
Attribute FormSocket.VB_VarHelpID = -1

Private Sub Command1_Click()
Dim aRes As Integer
    
    If Command1.Caption = "Connect" Then
                
                Command1.Enabled = False
                
                Me.Caption = "Server:" & txtServer & " Port:" & txtPort
                
                FormSocket.sckConnect txtHost.Text, txtPort.Text
                
                If FormSocket.State = sckError Then
                        
                        FormSocket.sckClose
                        Command1.Enabled = True
                        Command1.Caption = "Connect"
                 
                Else
                    
                    Command1.Enabled = True
                    Command1.Caption = "Disconnect"
                
                End If
    
    Else
                         
                FormSocket.sckClose
                
                txtData.Text = ""
                txtLOG.Text = ""
                Command1.Caption = "Connect"
                
                
                txtCMD.Enabled = False
                txtWaitFor.Enabled = False
                Command2.Enabled = False
                
    End If


End Sub
Private Sub Command2_Click()
Dim a As Variant

    If txtWaitFor = "" Then
    a = FormSocket.sckSendData(txtCMD.Text)
    Else
    
    a = FormSocket.sckSendCommand(txtCMD.Text, txtWaitFor.Text)
    
        'If FormSocket.LastCommandOK Then
           
        '   lbStatus.Caption = "Command OK"
        
        'Else
           
        '   lbStatus.Caption = "Command Error"
        
        'End If
        
    
    End If

End Sub


Private Sub Command3_Click()
    Dim b As New cFormSocket
    b.sckStart
    b.sckListen 4562
    
    a = FormSocket.sckSendCommand("USER anonymous", "331")
    a = FormSocket.sckSendCommand("PASS dd@dd.it", "230")
    a = FormSocket.sckSendCommand("CWD /pub/benchmark/bapco/", "250")
    a = FormSocket.sckSendCommand("TYPE I", "200")
    a = FormSocket.sckSendCommand("PORT 192,168,0,232,17,210", "200")
    a = FormSocket.sckSendCommand("RETR SQLSYS.EXE", "226")

    b.sckStop

    Set b = Nothing

End Sub

Private Sub Form_Load()
                
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
Private Sub FormSocket_sckBinaryData(pDATA As Variant)

    txtData = txtData & pDATA

End Sub

Private Sub FormSocket_sckConnected()
txtCMD.Enabled = True
txtWaitFor.Enabled = True
Command2.Enabled = True
End Sub

Private Sub FormSocket_sckDisconnectedByPeer()

    Call Command1_Click

End Sub

Private Sub FormSocket_sckLog(pLogEntry As String)
    
    txtLOG = txtLOG & pLogEntry & vbCrLf
    txtLOG.SelStart = Len(txtLOG.Text)

End Sub

