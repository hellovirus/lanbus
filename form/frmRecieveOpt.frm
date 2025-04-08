VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReceiveOpt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File Recieve"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5535
   Icon            =   "frmRecieveOpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog cdgSave 
      Left            =   4560
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取  消"
      Height          =   285
      Left            =   4440
      TabIndex        =   10
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "另存为"
      Default         =   -1  'True
      Height          =   285
      Left            =   3240
      TabIndex        =   9
      Top             =   3600
      Width           =   975
   End
   Begin VB.Frame frmProperties 
      Appearance      =   0  'Flat
      Caption         =   " 文件属性 "
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5295
      Begin VB.TextBox txtComments 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   1335
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1080
         Width           =   4125
      End
      Begin VB.TextBox txtFileSize 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "%FILESIZE%"
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "%FILENAME%"
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label lblComment 
         Alignment       =   1  'Right Justify
         Caption         =   "内容简介："
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label lblFileSize 
         Alignment       =   1  'Right Justify
         Caption         =   "文件大小："
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "文件名称："
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label lblChoose 
      AutoSize        =   -1  'True
      Caption         =   "点击“另存为”保存文件，或者点击“取消”拒绝接收文件！"
      Height          =   180
      Left            =   60
      TabIndex        =   8
      Top             =   3180
      Width           =   4860
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      Caption         =   "%FROM% 想传送给你一个文件！"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2430
   End
End
Attribute VB_Name = "frmReceiveOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyID As Long

Public Function Prepare(ByVal ID As Long)
  MyID = ID
  With ftRcv(ID)
    lblFrom = .From & " 想传送给一个文件."
    txtFileName = .FileName
    txtFileSize = FormatNumber(.FileSize / 1024, 2, vbTrue) & "KB"
    txtComments = Trim(.Comment)
    .frmReceive.Caption = "接收文件，来自于 " & .From
    .frmReceive.lblInfo = .FileName & " 来自于 " & .From
  End With
  Me.Visible = True
End Function

Private Sub cmdCancel_Click()
  ftRcv(MyID).frmReceive.wsReceive.SendData "DENIED"
  DoEvents
  Unload ftRcv(MyID).frmReceive
  Unload Me
End Sub

Private Sub cmdSave_Click()
  On Error GoTo ErrHandler
  With cdgSave
    .CancelError = True
    .FileName = ftRcv(MyID).FileName
    .DialogTitle = "Select Destination File"
    .Filter = "All Files (*.*)|*.*"
    .Flags = cdlOFNOverwritePrompt
    .ShowSave
  End With
  
  ftRcv(MyID).Destination = cdgSave.FileName
  ftRcv(MyID).frmReceive.wsReceive.SendData "ACCEPT"
  ftRcv(MyID).frmReceive.Visible = True
  Unload Me
  
ErrHandler:
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set ftRcv(MyID).frmRcOpt = Nothing
End Sub
