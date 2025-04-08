VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReceiving 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Receiving file from %FROM%"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5535
   Icon            =   "frmReceiving.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   StartUpPosition =   3  '窗口缺省
   Begin MSComCtl2.Animation aniTransfer 
      Height          =   675
      Left            =   60
      TabIndex        =   8
      Top             =   180
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1191
      _Version        =   393216
      FullWidth       =   109
      FullHeight      =   45
   End
   Begin VB.Timer tmrSpeed 
      Interval        =   1000
      Left            =   4560
      Top             =   120
   End
   Begin MSWinsockLib.Winsock wsReceive 
      Left            =   5040
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox chkClose 
      Caption         =   "文件传送完成后，关闭对话框！"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   3075
   End
   Begin VB.CommandButton cmdCancelClose 
      Caption         =   "取 消"
      Default         =   -1  'True
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdFolder 
      Caption         =   "打开文件夹"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "打开文件"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
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
   Begin VB.Label lblDownloaded 
      AutoSize        =   -1  'True
      Caption         =   "下载 %PERCENT%k of %SIZE% @ %SPEED%"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3150
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "%FILENAME% from %FROM%"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2130
   End
   Begin VB.Label lblSaving 
      AutoSize        =   -1  'True
      Caption         =   "保存:"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   450
   End
End
Attribute VB_Name = "frmReceiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyID As Long
Dim GotHeader As Boolean
Dim FileNum As Long
Dim Receivedbyt As Long
Dim ByteSec As Long, Speed As Long
Dim Complete As Boolean

Private Sub cmdCancelClose_Click()
  On Error Resume Next
  'Close the connection to stop
  Complete = True
  wsReceive.Close
  Close #FileNum
  Unload Me
End Sub

Private Sub cmdFolder_Click()
  Shell "explorer " & Left(ftRcv(MyID).Destination, Len(ftRcv(MyID).Destination) - Len(ftRcv(MyID).FileName)), vbNormalFocus
End Sub

Private Sub cmdOpen_Click()
Shell "Explorer " & ftRcv(MyID).Destination, vbNormalFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set ftRcv(MyID).frmReceive = Nothing
End Sub

Private Sub tmrSpeed_Timer()
  Speed = Format(ByteSec / 1024, "0.0")
  ByteSec = 0
End Sub

Private Sub Form_Load()
ReDim ResendChunk(0)
If FileExists(App.Path & "\media\filemove.avi") Then
   aniTransfer.Open App.Path & "\media\filemove.avi"
End If
End Sub

Public Function Prepare(ByVal ID As Long, ByVal requestID As Long)
  MyID = ID
  wsReceive.Accept requestID
End Function

Private Sub wsReceive_Close()
  On Error Resume Next
  If FileNum = 0 Then
    wsReceive.Close
    Unload Me
    Exit Sub
  End If
  If Not Complete Then
    MsgBox "File Transfer Ended Unexpectedly!", vbCritical + vbOKOnly, "Error"
    Close #FileNum
    Unload Me
  End If
End Sub

Private Sub wsReceive_DataArrival(ByVal bytesTotal As Long)
  If Not GotHeader Then
    Dim Dat As String
    wsReceive.GetData Dat$, vbString
    If Left(Dat$, 4) = "FILE" Then
      Dim FirstPos As Long, SecondPos As Long
      FirstPos = InStr(6, Dat, ":")
      SecondPos = InStr(FirstPos + 1, Dat, ":")
      With ftRcv(MyID)
        .FileName = Mid(Dat, 6, (FirstPos - 6))
        .FileSize = CDbl(Mid(Dat, FirstPos + 1, (SecondPos - FirstPos) - 1))
        .Comment = Right(Dat, 200)
        .From = wsReceive.RemoteHostIP
        .frmRcOpt.Prepare MyID
      End With
      GotHeader = True
    End If
  Else
    If FileNum = 0 Then
      FileNum = FreeFile
      On Error Resume Next
      If FileLen(ftRcv(MyID).Destination) > 0 Then Kill ftRcv(MyID).Destination
      Open ftRcv(MyID).Destination For Binary As #FileNum
    End If
    Dim GotDat() As Byte
    Dim Hash As String
    ByteSec = ByteSec + bytesTotal
    Receivedbyt = Receivedbyt + bytesTotal
    pgPercent.Value = (100 / ftRcv(MyID).FileSize) * Receivedbyt
    lblDownloaded = "接收 " & Int(pgPercent.Value) & "% of " & FormatNumber(ftSend(MyID).FileSize / 1024, 2, vbTrue) & _
            "Kb 速度：" & Speed & "Kb\秒"
    ReDim GotDat(0 To bytesTotal - 1)
    wsReceive.GetData GotDat, vbArray + vbByte
    Put #FileNum, , GotDat
    If Receivedbyt = ftRcv(MyID).FileSize Then
      Close #FileNum
      Complete = True
      cmdOpen.Enabled = True: cmdFolder.Enabled = True: cmdCancelClose.Caption = "关闭"
      If chkClose.Value = Checked Then
        wsReceive.Close
        Unload Me
      End If
    End If
  End If
End Sub
