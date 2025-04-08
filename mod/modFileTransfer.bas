Attribute VB_Name = "modFileTransfer"
Option Explicit

Public Type T_FILE_TRANSFER_SEND
  Comment         As String * 200 'The Comment Of The File
  To              As String       'IP/Host to send file
  FileToSend      As String       'The File Were Sending
  FileSize        As Double       'The Size Of The File
  frmChoose       As New frmFileChoose
  frmSend         As New frmSending
End Type

Public Type T_FILE_TRANSFER_RECEIVE
  Comment         As String * 200 'The Comment Of The File
  Destination     As String       'Save File To Here
  From            As String       'IP/Host of sending person
  FileSize        As Double       'The Size Of The File
  FileName        As String
  frmRcOpt        As New frmReceiveOpt
  frmReceive      As New frmReceiving
End Type

Public Const FT_BUFFER_SIZE = 5734  'CHANGE THIS IF YOU NEED TO
Public Const FT_USE_PORT = 361      'CHANGE THIS IF YOU NEED TO

Public ftSend()       As T_FILE_TRANSFER_SEND
Dim SendCount         As Long
Public ftRcv()        As T_FILE_TRANSFER_RECEIVE
Dim RcvCount          As Long

Public Function SendFile(ByVal Destination As String)
  ReDim Preserve ftSend(0 To SendCount)
  
  ftSend(SendCount).To = Destination
  ftSend(SendCount).frmChoose.ChooseSend SendCount
  SendCount = SendCount + 1
End Function

Public Function ConnectReq(ByVal requestID As Long)
  ReDim Preserve ftRcv(0 To RcvCount)
  
  ftRcv(RcvCount).frmReceive.Prepare RcvCount, requestID
  RcvCount = RcvCount + 1
End Function
