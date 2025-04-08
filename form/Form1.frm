VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "局域网聊天大巴 V0.1"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6795
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   6795
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtHost 
      Height          =   270
      Left            =   5220
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "对方的IP地址"
      Top             =   5820
      Width           =   1515
   End
   Begin VB.TextBox txtNick 
      Height          =   270
      Left            =   2580
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "对方的昵称"
      Top             =   5820
      Width           =   1695
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "Θ 广播 Θ"
      Height          =   315
      Left            =   2700
      TabIndex        =   2
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Θ 发送 Θ"
      Default         =   -1  'True
      Height          =   315
      Left            =   5580
      TabIndex        =   1
      Top             =   6180
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "请输入你想说的话！"
      Top             =   6180
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Θ 清空 Θ"
      Height          =   315
      Left            =   4140
      TabIndex        =   3
      Top             =   6600
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Form1.frx":08CA
      Left            =   420
      List            =   "Form1.frx":08DA
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保存记录"
      Height          =   315
      Left            =   5580
      TabIndex        =   4
      ToolTipText     =   "聊天记录保存在程序目录下的Chats.txt中！"
      Top             =   6600
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2820
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   99
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1800
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2204
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2706
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C08
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":310A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":360C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4010
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4512
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4A14
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4F16
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5418
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":591A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":631E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6820
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7224
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7726
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7C28
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":812A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":862C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9030
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9532
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9F36
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A438
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A93A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AE3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B33E
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B840
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BD42
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C244
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C746
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CC48
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D14A
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D64C
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DB4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E050
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E552
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EA54
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EF56
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F458
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F95A
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":FE5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10860
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11264
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11766
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1216A
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1266C
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12B6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13070
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13572
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14478
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1497A
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1537E
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15880
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15D82
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16284
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16786
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1718A
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1768C
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18090
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18592
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18A94
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19498
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1999A
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19E9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A39E
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A8A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1ADA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B2A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B7A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1BCA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C1AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C6AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1CBAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D0B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D5B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DAB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DFB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E4B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E9BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1EEBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F3BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3600
      Top             =   3900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   99
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F8C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":20BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":21EE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":231FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":24510
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":25824
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":26B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":27E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29160
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A474
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2B788
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2CA9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2DDB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2F0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":303D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":316EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":32A00
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":33D14
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":35028
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3633C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":37650
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":38964
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":39C78
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3AF8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3C2A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3D5B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3E8C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3FBDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":40EF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":42204
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":43518
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4482C
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":45B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":46E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":48168
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4947C
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4A790
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4BAA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4CDB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4E0CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4F3E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":506F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":51A08
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":52D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":54030
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":55344
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":56658
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5796C
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":58C80
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":59F94
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5B2A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5C5BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5D8D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5EBE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5FEF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6120C
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":62520
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":63834
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":64B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":65E5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":67170
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":68484
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":69798
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6AAAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6BDC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6D0D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6E3E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6F6FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":70A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":71D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":73038
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7434C
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":75660
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":76974
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":77C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":78F9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7A2B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7B5C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7C8D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7DBEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7EF00
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":80214
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":81528
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8283C
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":83B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":84E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":86178
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8748C
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":887A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":89AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8ADC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8C0DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8D3F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8E704
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8FA18
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":90D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":92040
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":93354
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":94668
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckSend 
      Left            =   6360
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "请点击相应头像，来选择你的私聊的对象！"
      Top             =   1260
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   9340
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "昵称"
         Object.Width           =   2558
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP地址"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "登录时间"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsListen 
      Left            =   5880
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox ChatTxt 
      Height          =   4455
      Left            =   1560
      TabIndex        =   12
      Top             =   1260
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7858
      _Version        =   393217
      BackColor       =   12648447
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"Form1.frx":9597C
   End
   Begin VB.Label Label1 
      Caption         =   "对方IP:"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   5820
      Width           =   675
   End
   Begin VB.Label Label4 
      Caption         =   "对方昵称:"
      Height          =   255
      Left            =   1620
      TabIndex        =   10
      Top             =   5820
      Width           =   915
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1860
      Picture         =   "Form1.frx":95A19
      Top             =   6480
      Width           =   480
   End
   Begin VB.Label Label5 
      Caption         =   "显示方式"
      Height          =   435
      Left            =   0
      TabIndex        =   9
      Top             =   6600
      Width           =   435
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1230
      Left            =   0
      Picture         =   "Form1.frx":95D23
      Top             =   0
      Width           =   6780
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ChatTxt_LostFocus()
ChatTxt.SelStart = Len(ChatTxt)
End Sub

Private Sub cmdC_Click()
Dim DataS As String
Debug.Print Len(Trim(Combo1.Text))
    If Len(Trim(Combo1.Text)) < 1 Then  '必须输入信息
        MsgBox "不能发送空信息!", vbOKOnly, "LanChatBus"
        Combo1.SetFocus '把光标放到信息输入框中
        Exit Sub
    End If

    Combo1.SetFocus
    DataS = "4" & MyNickN & ":" & Combo1.Text
    SendM GBIP, DataS
End Sub

Private Sub cmdSend_Click()
Dim DataS As String
Dim data5 As String
    If Len(Trim(Combo1.Text)) < 1 Then  '必须输入信息
        MsgBox "不能发送空信息!", vbOKOnly, "LanChatBus"
        Combo1.SetFocus '把光标放到信息输入框中
        Exit Sub
    End If
    If Len(Trim(txtHost.Text)) < 1 Then  '必须输入信息
       Combo1.SetFocus
       DataS = "4" & MyNickN & ":" & Combo1.Text
       SendM GBIP, DataS
       Exit Sub
    End If
    
            data5 = "⊙ 私聊 → 【 " & txtNick.Text & " 】 " & Time() & Chr$(13) & Chr$(10) & "  " & Combo1.Text & Chr$(13) & Chr$(10)
            'txtMain.Text = txtMain.Text & data5
            AddText ChatTxt, data5, vbRed
            'txtMain.SelStart = Len(txtMain) '显示最后一行内容
            'ChatTxt.SelStart = Len(ChatTxt)

    Combo1.SetFocus
    DataS = "3" & MyNickN & ":" & Combo1.Text
    SendM txtHost.Text, DataS
End Sub

Private Sub Combo1_GotFocus()
Combo1.SelStart = 0
Combo1.SelLength = Len(Combo1.Text)
End Sub

Private Sub Combo2_Click()
Select Case Combo2.ListIndex
Case 0:
    ListView1.View = lvwIcon
Case 1:
    ListView1.View = lvwSmallIcon
Case 2:
    ListView1.View = lvwList
Case 3:
    ListView1.View = lvwReport
End Select

End Sub
Private Sub Command1_Click()
'txtMain.Text = ""
If MsgBox("在清空之前是否要保存聊天记录？", vbYesNo, "清空聊天记录") = vbYes Then
   Dim ChatFile As String
'Dim sendMsg As String
   Dim iFile As Integer
   iFile = FreeFile
   ChatFile = App.Path & "\Chats.txt"
     Open ChatFile For Append As iFile
     Print #iFile, "Year:" & Date & "|| Time: " & Time
     Print #iFile, ChatTxt.Text
     Close iFile
   MsgBox "聊天记录成功保存在程序目录下的Chats.txt中，请查看！"
   ChatTxt.Text = ""
Else
   ChatTxt.Text = ""
End If
End Sub

Private Sub Command2_Click()
Dim ChatFile As String
'Dim sendMsg As String
Dim iFile As Integer
iFile = FreeFile
ChatFile = App.Path & "\Chats.txt"
    Open ChatFile For Append As iFile
    Print #iFile, "Year:" & Date & "|| Time: " & Time
    Print #iFile, ChatTxt.Text
    Close iFile
MsgBox "聊天记录成功保存在程序目录下的Chats.txt中，请查看！"
End Sub

Private Sub Form_Load()
'以下是不让本程序重复运行，暂时不开，测试无误后再开此功能！
'If App.PrevInstance Then
'   MsgBox "本系统已经加载，请不要重复运行本程序！", 48, "提示"
'   End
'End If
      ListView1.Icons = ImageList2
      ListView1.SmallIcons = ImageList1
      ListView1.BackColor = &HFFE4C7
      Debug.Print MyIP, MyNickN, GBIP, MyInfo, MyFace
      'txtMain.Text = "√ 【 " & MyNickN & " 】成功登录于：" & Date & "日" & Time() & Chr$(13) & Chr$(10)
      'txtMain.SelStart = Len(txtMain) 'scroll that chatroom down
      AddText ChatTxt, "√ 【 " & MyNickN & " 】成功登录于：" & Date & "日" & Time() & Chr$(13) & Chr$(10), &HFF00FF
      ChatTxt.SelStart = Len(ChatTxt.Text)
      
    sckSend.Protocol = sckUDPProtocol 'set protocol. For this type of chat, we are using UDP
    Combo2.ListIndex = 0
    sckSend.RemoteHost = GBIP '设置主机IP
    sckSend.LocalPort = 50431 '设置本机端口
    sckSend.RemotePort = 50431 '设置主机端口
    sckSend.Bind '绑定端口，连接成功

  wsListen.LocalPort = FT_USE_PORT
  wsListen.Listen

    SendM GBIP, "1" & MyInfo
    'MyIP = GetIPAddress()
    'Debug.Print GetIPAddress() '显示出本机IP地址，进而计算出本局域网的广播地址
    'GBIP = Left(MyIP, InStrRev(MyIP, ".")) & "255"
  If FileExists(App.Path & "\chat.txt") Then
    Open App.Path & "\chat.txt" For Input As #1
    Do
      Input #1, Gdata
      Combo1.AddItem Gdata
      If EOF(1) Then Close #1: Exit Sub
    Loop
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '当按下 关闭窗口 后
    SendM GBIP, "5" & MyInfo
    If sckSend.State <> sckClosed Then
       sckSend.Close
    End If
    End
End Sub

Private Sub Image1_Click()
        Dim Sm As String
        Dim Inf As String
         hh0$ = Chr$(13) + Chr$(10)
         Sm$ = "                                  局域网聊天大巴(ChatBus) " + hh0$
         Sm$ = Sm$ + "" + hh0$
         Sm$ = Sm$ + "                                       程序编制：张海瑞" + hh0$
         Sm$ = Sm$ + "                                       界面设计：张海瑞" + hh0$
         Sm$ = Sm$ + "" + hh0$
         Sm$ = Sm$ + "    ※  Hurry ChatBus属免费软件。我自学VB八年余，深知编程之苦乐，有时" + hh0$
         Sm$ = Sm$ + "为某一功能的实现要花费许多时间，概因周围无可交流人员。为使后学者在" + hh0$
         Sm$ = Sm$ + "某些方面少走弯路，特作此软件，并公布源程序，您可以免费传播、使用。" + hh0$
         Sm$ = Sm$ + "    ※  同时也希望更多的程序员公布源代码，促进中国软件事业的发展!" + hh0$
         Sm$ = Sm$ + "    ※  若您有疑问，可写信至：E-mail:zhanghairui@56.com " + hh0$
         Sm$ = Sm$ + "" + hh0$
         Sm$ = Sm$ + "关于本软件：" + hh0$
         Sm$ = Sm$ + "    ※  本聊天程序的出现是由于在暑假期间，学校的校园网不知怎么挂了，" + hh0$
         Sm$ = Sm$ + "大家不能上网聊天了，相互之间的资源也不能方便的共享了，聊天要楼上楼" + hh0$
         Sm$ = Sm$ + "下的跑，为解决这个问题，我编写了本软件，目的是让同学们方便的在局域" + hh0$
         Sm$ = Sm$ + "网内聊天，共享好的资源！(共享功能的实现比较麻烦,V0.1版暂不推出！)" + hh0$
         Sm$ = Sm$ + "" + hh0$
         Sm$ = Sm$ + "                                    海瑞写于06年7月 遵义医学院" + hh0$
         
Inf$ = "" + hh0$
Inf$ = Inf$ + hh0$
'Inf$ = Inf$ + hh0$
Inf$ = Inf$ + "╔━㊣━━━━━━━━━━━╗" + hh0$
Inf$ = Inf$ + "┃☆江湖独孤行★我自仰天笑☆┃" + hh0$
Inf$ = Inf$ + "╚━━━━━━━━━━━㊣━╝" + hh0$
'Inf$ = Inf$ + hh0$
Inf$ = Inf$ + hh0$
Inf$ = Inf$ + "                       Hahaha" + hh0$
'Inf$ = Inf$ + hh0$
         MsgBox Sm, vbOKOnly, "关 于 本 软 件 !"
         MsgBox Inf, vbOKOnly, "关于我自己!"

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  With ListView1
      .SortKey = ColumnHeader.Index - 1
      .SortOrder = Abs(Not .SortOrder = 1)
  End With
End Sub

Private Sub ListView1_DblClick()
SendFile ListView1.SelectedItem.SubItems(1)  'change destinationEnd Sub
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
txtHost.Text = Item.SubItems(1)
txtNick.Text = Item.Text
End Sub

Private Sub sckSend_DataArrival(ByVal bytesTotal As Long)
    '接收到数据以后的处理是本程序最重要的部分！！！！
    '把传送的数据分为以下五种：
    '1.传送“我来了” + 用户名、IP地址、端口、主机名
    '2.传送“欢迎” + 用户名、IP地址、端口、主机名
    '3.传送“单独聊天内容” + 用户名、IP地址、端口、主机名
    '4 传送“共同聊天内容” + 用户名、IP地址、端口、主机名
    '5.传送“再见”
    '
    Dim TheData As String
    
    Dim NickN As String
    Dim IPP As String
    Dim HostN As String
    
    Dim FF As Integer
    Dim SF As Integer
    
    Dim clmX As ColumnHeader
    Dim itmX As ListItem
    
    Dim itmFound  As ListItem
    
    
    On Error GoTo ClearChat
    
    sckSend.GetData TheData, vbString 'extract the data
    
    Select Case Left(TheData, 1)
      Case 1 '提取确认码1后，分解内容记录后显示在listview1中，并回复一个2 ！广播 ！
      'MsgBox "收到1开头的数据！"
      
      FF = InStr(Trim(TheData), "|") '第一个标志位置
      SF = InStrRev(Trim(TheData), "|") '第二个标志位置
      
        NickN = Mid(Trim(TheData), 2, (FF - 2)) '找出昵称
        IPP = Mid(Trim(TheData), (FF + 1), (SF - FF - 1)) '找出IP地址
        HostN = Right(Trim(TheData), (Len(TheData) - SF)) '找出头像号

        Set itmFound = ListView1.FindItem(IPP, lvwSubItem, , lvwPartial) '看看现在的IP列表中有没有这个要加入的IP地址

If itmFound Is Nothing Then ' 如果没有找到，说明列表中没有，需要加入这个IP到列表中
'listview1.ListItems.Count
      Set itmX = ListView1.ListItems.Add(, , NickN) '加入昵称
        itmX.SubItems(1) = IPP '加入IP地址
        itmX.SubItems(2) = CStr(Time())
        itmX.Icon = CInt(HostN) '显示大头像
        itmX.SmallIcon = CInt(HostN) '显示小头像
Else '如果已经有了这个IP，说明已经有了一个，但可能昵称改后重新登录的，所以要删除原来的，再加入新的
        ListView1.ListItems.Remove itmFound.Index '删除掉原来的
        Set itmX = ListView1.ListItems.Add(, , NickN) '再加入更改过的
        itmX.SubItems(1) = IPP '加入IP地址
        itmX.Icon = CInt(HostN) '显示大头像
        itmX.SmallIcon = CInt(HostN) '显示小头像
    Exit Sub '如果已经有了这个IP，说明已经有了
End If
        SendM IPP, "2" & MyInfo '★★★最重要的是再回复一个2以及自己的信息给上线的机子
     
      Case 2 '提取确认码2后，记录后显示在listview1中,确定对方也在线
       'MsgBox "收到2开头的数据！"
       
      FF = InStr(Trim(TheData), "|") '第一个标志位置
      SF = InStrRev(Trim(TheData), "|") '第二个标志位置
      
        NickN = Mid(Trim(TheData), 2, (FF - 2)) '找出昵称
        IPP = Mid(Trim(TheData), (FF + 1), (SF - FF - 1)) '找出IP地址
        HostN = Right(Trim(TheData), (Len(TheData) - SF)) '找出头像号

        Set itmFound = ListView1.FindItem(IPP, lvwSubItem, , lvwPartial) '看看现在的IP列表中有没有这个要加入的IP地址

If itmFound Is Nothing Then ' 如果没有找到，说明列表中没有，需要加入这个IP到列表中
      Set itmX = ListView1.ListItems.Add(, , NickN) '加入昵称
        itmX.SubItems(1) = IPP '加入IP地址
        itmX.SubItems(2) = CStr(Time())
        itmX.Icon = CInt(HostN) '显示大头像
        itmX.SmallIcon = CInt(HostN) '显示小头像
Else '如果已经有了这个IP，说明已经有了一个，但可能昵称改后重新登录的，所以要删除原来的，再加入新的
        ListView1.ListItems.Remove itmFound.Index '删除掉原来的
        Set itmX = ListView1.ListItems.Add(, , NickN) '再加入更改过的
        itmX.SubItems(1) = IPP '加入IP地址
        itmX.SubItems(2) = CStr(Time())
        itmX.Icon = CInt(HostN) '显示大头像
        itmX.SmallIcon = CInt(HostN) '显示小头像
    Exit Sub '如果已经有了这个IP，说明已经有了
End If

        
      Case 3 '提取确认码3后，再加上IP或者主机名显示在聊天对话框中
            'MsgBox "收到3开头的数据！"
            Dim Data3 As String
            'Data3 = "●私聊→" & Right(TheData, Len(TheData) - 1) & "↘" & Time() & Chr$(13) & Chr$(10)
            Data3 = "●私聊 → " & Time() & Chr$(13) & Chr$(10) & "  " & Right(TheData, Len(TheData) - 1) & Chr$(13) & Chr$(10)
            'txtMain.Text = txtMain.Text & Data3
            'txtMain.SelStart = Len(txtMain) '显示最后一行内容
            AddText ChatTxt, Data3, vbBlue
            'ChatTxt.SelStart = Len(ChatTxt.Text)
        
      Case 4 '提取确认码4后，再直接显示在聊天对话框中  ！广播 ！
            'MsgBox "收到4开头的数据！"
            Dim Data4 As String
            'Data4 = "→" & Right(TheData, Len(TheData) - 1) & "↘" & Time() & Chr$(13) & Chr$(10)
            Data4 = "※广播 → " & Time() & Chr$(13) & Chr$(10) & "  " & Right(TheData, Len(TheData) - 1) & Chr$(13) & Chr$(10)
            'txtMain.Text = txtMain.Text & Data4
            'txtMain.SelStart = Len(txtMain) '显示最后一行内容
            AddText ChatTxt, Data4, vbBlack
            'ChatTxt.SelStart = Len(ChatTxt.Text)
            
      Case 5 '提取确认码5后，再删除listview1中的相应数据 ！广播 ！
            'MsgBox "收到5开头的数据！"
            
      FF = InStr(Trim(TheData), "|") '第一个标志位置
      SF = InStrRev(Trim(TheData), "|") '第二个标志位置
      
        NickN = Mid(Trim(TheData), 2, (FF - 2)) '找出昵称
        IPP = Mid(Trim(TheData), (FF + 1), (SF - FF - 1)) '找出IP地址
        HostN = Right(Trim(TheData), (Len(TheData) - SF)) '找出头像号

        Set itmFound = ListView1.FindItem(IPP, lvwSubItem, , lvwPartial) '看看现在的IP列表中有没有这个要加入的IP地址

If itmFound Is Nothing Then ' 如果没有找到，说明列表中没有
   Exit Sub
Else '如果已经有了这个IP，说明已经有了一个，所以要删除原来的，再加入新的
        ListView1.ListItems.Remove itmFound.Index '删除掉原来的
    Exit Sub
End If
        
      Case Else
    End Select
        'txtMain.Text = txtMain.Text & TheData & Chr$(13) & Chr$(10) 'add the data to our chatroom
        'txtMain.SelStart = Len(txtMain) 'scroll that chatroom down
    Exit Sub
ClearChat:
    MsgBox "出现不可知错误！", vbOKOnly, "LanChatBus"
    txtMain.Text = ""
End Sub

Private Sub sckSend_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "winsock连接过程中出现未知错误!", vbOKOnly, "LanChatBus"
    End
End Sub

Function SendM(IPD As String, MSGG As String)
'Dim NickN As String
'Dim IPP As String
'Dim HostN As String

    sckSend.RemoteHost = IPD '设置主机IP
    'sckSend.LocalPort = txtLocalP '设置本机端口
    'sckSend.RemotePort = txtRemoteP '设置主机端口
    'sckSend.Bind '绑定端口，连接成功
    sckSend.SendData MSGG
End Function
Private Sub AddText(ByRef RTFBox As RichTextBox, ByVal strText As String, ByVal tColor As Long)
    Dim lngLength As Long
    Dim lngSelStart As Long
    Dim lngTLength As Long
    
    'lngLength = Len(strText)
    
    lngTLength = Len(strText)
    
    lngSelStart = RTFBox.SelStart '光标开始处
    RTFBox.SelLength = 0 '选择长度
    RTFBox.SelText = strText '选择文本
    RTFBox.SelStart = lngSelStart '开始点
    RTFBox.SelLength = lngTLength '选择长度
    RTFBox.SelColor = tColor '设置颜色
    RTFBox.SelLength = 0 '选择长度
    RTFBox.SelStart = lngSelStart + lngTLength   '
    'RTFBox.SelStart = Len(RTFBox.Text)
End Sub


Private Sub wsListen_ConnectionRequest(ByVal requestID As Long)
  ConnectReq requestID
End Sub

