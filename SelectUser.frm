VERSION 5.00
Begin VB.Form frmSelectUser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccione Usuario"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   465
      Left            =   1890
      TabIndex        =   2
      Top             =   810
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seleccionar"
      Default         =   -1  'True
      Height          =   465
      Left            =   180
      TabIndex        =   1
      Top             =   810
      Width           =   1185
   End
   Begin VB.ComboBox cboUser 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Width           =   3165
   End
End
Attribute VB_Name = "frmSelectUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strUser As String

Private Sub Command1_Click()
    strUser = cboUser.Text
    Me.Hide
End Sub

Private Sub Command2_Click()
    strUser = ""
    Me.Hide

End Sub

Public Function getuser() As String
Dim recUsers As New ADODB.Recordset
    recUsers.Open "SELECT username from users order by username", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not recUsers.EOF
        cboUser.AddItem recUsers("username").Value & ""
        recUsers.MoveNext
    Loop
    Set recUsers = Nothing
    Me.Show vbModal, fmrMain
    getuser = strUser
    Unload Me
End Function
