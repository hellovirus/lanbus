VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSending 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sending file to %TO%"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5550
   Icon            =   "frmSending.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5550
   StartUpPosition =   3  '窗口缺省
   Begin MSComCtl2.Animation aniTransfer 
      Height          =   675
      Left            =   180
      TabIndex        =   6
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1191
      _Version        =   393216
      FullWidth       =   101
      FullHeight      =   45
   End
   Begin VB.Timer tmrSpeed 
      Interval        =   1000
      Left            =   4560
      Top             =   120
   End
   Begin MSWinsockLib.Winsock wsSend 
      Left            =   5040
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancelClose 
      Caption         =   "取 消"
      Default         =   -1  'True
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox chkClose 
      Caption         =   "文件传送完毕自动关闭本对话框！"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   3315
   End
   Begin MSComctlLib.ProgressBar pgPercent 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblSending 
      AutoSize        =   -1  'True
      Caption         =   "发送:"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   450
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "%FILENAME% to %TO%"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1725
   End
   Begin VB.Label lblSent 
      Caption         =   "正在等待接收文件..."
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   5250
   End
End
Attribute VB_Name = "frmSending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyID As Long
Dim FileNum As Long
Dim FileName As String
Dim RCVAccept As Boolean
Dim Sentbyt As Long
Dim ByteSec As Long, Speed As Long
Dim Complete As Boolean

Public Function InitTransfer(ByVal ID As Long)
  MyID = ID
  FileName = Mid(ftSend(MyID).FileToSend, InStrRev(ftSend(MyID).FileToSend, "\") + 1)
  Caption = "发送文件给 " & ftSend(MyID).To
  lblInfo = FileName & " 给 " & ftSend(MyID).To
  'Attempt to connect to the Destination
  wsSend.Connect ftSend(MyID).To, FT_USE_PORT
  Me.Visible = True
End Function

Private Sub cmdCancel_Click()
On Error Resume Next
  Complete = True
  Close #FileNum
  If chkClose.Value = vbUnchecked Then Unload Me
End Sub

Private Sub cmdCancelClose_Click()
  On Error Resume Next
  'Close the connection to stop
  Complete = True
  wsSend.Close
  Close #FileNum
  Unload Me
End Sub

Private Sub Form_Load()
If FileExists(App.Path & "\media\filemove.avi") Then
  aniTransfer.Open App.Path & "\media\filemove.avi"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Remove the form from memory
  Set ftSend(MyID).frmSend = Nothing
End Sub

Private Sub tmrSpeed_Timer()
  Speed = Format(ByteSec / 1024, "0.0")
  ByteSec = 0
End Sub

Private Sub wsSend_Close()
  On Error Resume Next
  If Not Complete Then
    MsgBox "File Transfer Ended Unexpectedly!", vbCritical + vbOKOnly, "Error"
    Close #FileNum
    Unload Me
  End If
End Sub

Private Sub wsSend_Connect()
  'Send Information regarding the file
  wsSend.SendData "FILE:" & FileName & ":" & ftSend(MyID).FileSize & ":" & ftSend(MyID).Comment
End Sub

Private Sub wsSend_DataArrival(ByVal bytesTotal As Long)

    Dim Dat As String
    wsSend.GetData Dat, vbString
    If Trim$(Dat$) = "ACCEPT" Then
      Call SendChunk
    ElseIf Trim$(Dat$) = "DENIED" Then
      MsgBox "文件被对方拒绝接收!", vbInformation + vbOKOnly, "File Rejected"
      'Close the connection
      wsSend.Close
      'unload the form
      Unload Me
    End If
    
End Sub

Private Sub wsSend_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Select Case Number
    Case sckConnectionRefused, sckHostNotFound, sckHostNotFoundTryAgain
      'couldnt connect
      MsgBox "Could Not Connect To Remote Host!", vbCritical + vbOKOnly, _
             "Error " & Number
      'Close the form
      Unload Me
  End Select
End Sub

Public Function SendChunk()
  'This is where we send the file data
  Dim ChunkSize As Long
  Dim Chunk() As Byte
  Dim arrHash() As Byte
  
  If wsSend.State <> sckConnected Then Exit Function
  
  ChunkSize = FT_BUFFER_SIZE
  If FileNum = 0 Then 'No data has been sent yet, open the file
    FileNum = FreeFile
    Open ftSend(MyID).FileToSend For Binary As #FileNum
  End If
  
  'determine chunk size
  If (LOF(FileNum) - Loc(FileNum)) < FT_BUFFER_SIZE Then _
     ChunkSize = (LOF(FileNum) - Loc(FileNum))
  'set array size to fit chunk
  ReDim Chunk(0 To ChunkSize - 1)
  'read the chunk
  Get #FileNum, , Chunk
  'Send the data
  wsSend.SendData Chunk
  Sentbyt = Sentbyt + ChunkSize
  ByteSec = ByteSec + ChunkSize
  pgPercent.Value = (100 / ftSend(MyID).FileSize) * Sentbyt
  lblSent = "发送 " & Int(pgPercent.Value) & "% of " & FormatNumber(ftSend(MyID).FileSize / 1024, 2, vbTrue) & _
            "Kb 速度：" & Speed & "Kb/秒"
  
  'See if file is sent
  If Sentbyt = ftSend(MyID).FileSize Then
    Complete = True
    Close #FileNum
    cmdCancelClose.Caption = "关闭"
  End If
End Function

Private Sub wsSend_SendComplete()
  DoEvents
  If FileNum > 0 Then
      If Not Complete Then
      SendChunk
    Else
      If chkClose.Value = Checked Then
        wsSend.Close
        Unload Me
      End If
    End If
  End If
End Sub
