VERSION 5.00
Begin VB.Form frmListen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLOG 
      Height          =   1485
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2250
      Width           =   5895
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   360
      Width           =   5895
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   315
      Left            =   4230
      TabIndex        =   2
      Top             =   30
      Width           =   1695
   End
   Begin VB.TextBox txtListenPort 
      Height          =   315
      Left            =   930
      TabIndex        =   0
      Text            =   "4559"
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   315
      Left            =   2700
      TabIndex        =   5
      Top             =   30
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Port"
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   855
   End
End
Attribute VB_Name = "frmListen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents FormSocket As cFormSocket
Attribute FormSocket.VB_VarHelpID = -1

Private Sub cmdListen_Click()

    If cmdListen.Caption = "Listen" Then
        Me.Caption = "Listening on Port: " & txtListenPort
        
        Set FormSocket = New cFormSocket
        FormSocket.sckHook Me.hWnd
        
        FormSocket.sckListen CLng(txtListenPort.Text)
    
        cmdListen.Caption = "Disconnect"
    
    Else
        
        Me.Caption = "Socket"
        
        FormSocket.sckClose
        FormSocket.sckUnHook
        
        Set FormSocket = Nothing
        
        cmdListen.Caption = "Listen"
        
        txtLOG = ""
    
    End If

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
    Label1 = LenB(pDATA)

End Sub
Private Sub FormSocket_sckLog(pLogEntry As String)

    txtLOG = txtLOG & pLogEntry & vbCrLf
    txtLOG.SelStart = Len(txtLOG.Text)

End Sub

