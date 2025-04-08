VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFileChoose 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Send File"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4380
   Icon            =   "frmFileChoose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   194
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   292
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog cdgSelect 
      Left            =   240
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取  消"
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "发送文件"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   120
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   4125
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "浏览"
      Height          =   285
      Left            =   3510
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblComments 
      AutoSize        =   -1  'True
      Caption         =   "简介：（最多不超过200个字节）"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2610
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "欲传送文件："
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmFileChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyID        As Long
Dim SendClicked As Boolean

Private Sub cmdBrowse_Click()
  'Setup error handling
  On Error GoTo Err_DetermineErr

  'Set Properties for the Common Dialog Control
  With cdgSelect
    'An error will occur if user clicks Cancel
    .CancelError = True
    
    .DialogTitle = "请选择要传送的文件："
    .Filter = "All Files (*.*)|*.*"
    .ShowOpen
  End With
  
  'Show the user the selected file
  txtFile = cdgSelect.FileName
  Exit Sub

Err_DetermineErr:
  'Cancel Was Selected, Do Nothing
End Sub

Private Sub cmdCancel_Click()
  'unload the form
  Unload Me
End Sub

Private Sub cmdSend_Click()
  With ftSend(MyID)
    .Comment = txtComments
    .FileSize = CDbl(FileLen(txtFile))
    .FileToSend = txtFile
    .frmSend.InitTransfer MyID
  End With
  SendClicked = True
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim MsgRet As VbMsgBoxResult 'Message Box Return Value
  
  'If there is a valid file displayed in txtFile, Prompt
  'the user to verify the cancel command
  If (cmdSend.Enabled = True) And (SendClicked = False) Then
    MsgRet = MsgBox("你已经选择了一个要发送的文件。" & _
                       vbNewLine & vbNewLine & "你确定吗" & _
                       " 你想取消吗?", vbYesNo, _
                       "是否取消")
    'If yes, remove form from memory. else cancel unload
    If MsgRet = vbYes Then Set frmFileChoose = Nothing Else _
                                               Cancel = -1
  Else
    'Remove the form from memory
    Set ftSend(MyID).frmChoose = Nothing
  End If
End Sub

Private Sub txtFile_Change()
  On Error GoTo ErrHandler
  
  'Disable the send command button if no file is selected
  If FileLen(txtFile) <> 0 Then cmdSend.Enabled = True Else _
                                cmdSend.Enabled = False
  Exit Sub

ErrHandler:
  'The file doesnt exist, so disable the send button
  cmdSend.Enabled = False
End Sub

Public Function ChooseSend(ByVal ID As Long)
  MyID = ID
  Me.Visible = True
End Function
