VERSION 5.00
Begin VB.Form frmTesting 
   Caption         =   "Socket!"
   ClientHeight    =   1170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1170
   ScaleWidth      =   4095
   WindowState     =   1  'Minimized
End
Attribute VB_Name = "frmTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mFormSocket As cFormSocket
Attribute mFormSocket.VB_VarHelpID = -1
Property Set FormSocket(pFSocket As cFormSocket)
    
    Set FormSocket = pFSocket

End Property



Private Sub Form_Load()

End Sub
