VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "数独游戏"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9615
   DrawWidth       =   4
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   9615
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      MaxLength       =   2
      TabIndex        =   88
      Text            =   "10"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   87
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "清空"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   86
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "求解"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   85
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "验证"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   84
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "生成"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   83
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   82
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   81
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   80
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   79
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   78
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   77
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   24
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   76
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   25
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   75
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   26
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   74
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   33
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   73
      Text            =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   34
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   72
      Text            =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   35
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   71
      Text            =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   42
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   70
      Text            =   "0"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   43
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   69
      Text            =   "0"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   44
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   68
      Text            =   "0"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   51
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   67
      Text            =   "0"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   52
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   66
      Text            =   "0"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   53
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   65
      Text            =   "0"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   60
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   64
      Text            =   "0"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   61
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   63
      Text            =   "0"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   62
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   62
      Text            =   "0"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   69
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   61
      Text            =   "0"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   70
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   60
      Text            =   "0"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   71
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   59
      Text            =   "0"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   78
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   58
      Text            =   "0"
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   79
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   57
      Text            =   "0"
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   80
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   56
      Text            =   "0"
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   55
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   54
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   53
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   52
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   51
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   50
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   21
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   49
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   22
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   48
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   23
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   47
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   30
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   46
      Text            =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   31
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   45
      Text            =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   32
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   44
      Text            =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   39
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   43
      Text            =   "0"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   40
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   42
      Text            =   "0"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   41
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   41
      Text            =   "0"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   48
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   40
      Text            =   "0"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   49
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   39
      Text            =   "0"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   50
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   38
      Text            =   "0"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   57
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   37
      Text            =   "0"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   58
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   36
      Text            =   "0"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   59
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   35
      Text            =   "0"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   66
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   34
      Text            =   "0"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   67
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   33
      Text            =   "0"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   68
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   32
      Text            =   "0"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   75
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   31
      Text            =   "0"
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   76
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   30
      Text            =   "0"
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   77
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   29
      Text            =   "0"
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   74
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   28
      Text            =   "0"
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   73
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   27
      Text            =   "0"
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   72
      Left            =   360
      MaxLength       =   1
      TabIndex        =   26
      Text            =   "0"
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   65
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   25
      Text            =   "0"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   64
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   24
      Text            =   "0"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   63
      Left            =   360
      MaxLength       =   1
      TabIndex        =   23
      Text            =   "0"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   56
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   22
      Text            =   "0"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   55
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   21
      Text            =   "0"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   54
      Left            =   360
      MaxLength       =   1
      TabIndex        =   20
      Text            =   "0"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   47
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   19
      Text            =   "0"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   46
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   18
      Text            =   "0"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   45
      Left            =   360
      MaxLength       =   1
      TabIndex        =   17
      Text            =   "0"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   38
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   16
      Text            =   "0"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   37
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   15
      Text            =   "0"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   36
      Left            =   360
      MaxLength       =   1
      TabIndex        =   14
      Text            =   "0"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   29
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   28
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   12
      Text            =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   27
      Left            =   360
      MaxLength       =   1
      TabIndex        =   11
      Text            =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   20
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   10
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   19
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   18
      Left            =   360
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   360
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   360
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "空白数量："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7440
      TabIndex        =   89
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label bottom 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "小组成员：许婷 景阳 钟泳诗"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2730
      TabIndex        =   2
      Top             =   7440
      Width           =   3615
   End
   Begin VB.Label TOP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "数独游戏"
      BeginProperty Font 
         Name            =   "方正舒体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim index As Integer
Dim x(81) As Integer

Private Sub Command1_Click()
Dim i, j, a, s, f As Integer

For i = 0 To 80
    x(i) = 0
Next

t = Array( _
9, 7, 6, 8, 5, 4, 2, 3, 1, 2, 5, 3, 6, 9, 1, 7, 8, 4, 4, 1, 8, 3, 7, 2, 5, 6, 9, 7, 6, 1, 4, 2, 3, 9, 5, 8, 5, 8, 4, 9, 6, 7, 1, 2, 3, 3, 2, 9, 1, 8, 5, 4, 7, 6, 6, 3, 5, 7, 4, 9, 8, 1, 2, 1, 4, 2, 5, 3, 8, 6, 9, 7, 8, 9, 7, 2, 1, 6, 3, 4, 5, _
4, 1, 6, 5, 2, 8, 3, 7, 9, 2, 8, 9, 3, 7, 6, 4, 1, 5, 7, 3, 5, 9, 1, 4, 6, 8, 2, 9, 4, 7, 1, 8, 3, 2, 5, 6, 1, 5, 3, 2, 6, 9, 7, 4, 8, 8, 6, 2, 7, 4, 5, 1, 9, 3, 6, 9, 1, 4, 5, 2, 8, 3, 7, 3, 7, 8, 6, 9, 1, 5, 2, 4, 5, 2, 4, 8, 3, 7, 9, 6, 1, _
3, 6, 8, 2, 1, 9, 5, 4, 7, 1, 9, 2, 5, 4, 7, 8, 3, 6, 4, 7, 5, 6, 8, 3, 9, 2, 1, 6, 5, 3, 8, 7, 1, 2, 9, 4, 2, 1, 4, 9, 6, 5, 3, 7, 8, 7, 8, 9, 3, 2, 4, 6, 1, 5, 9, 3, 1, 4, 5, 8, 7, 6, 2, 5, 4, 6, 7, 3, 2, 1, 8, 9, 8, 2, 7, 1, 9, 6, 4, 5, 3, _
5, 9, 8, 6, 1, 7, 2, 4, 3, 1, 3, 7, 8, 2, 4, 6, 9, 5, 2, 6, 4, 9, 3, 5, 7, 1, 8, 4, 5, 2, 7, 6, 8, 9, 3, 1, 3, 8, 1, 4, 9, 2, 5, 7, 6, 6, 7, 9, 1, 5, 3, 4, 8, 2, 9, 1, 3, 2, 4, 6, 8, 5, 7, 8, 2, 5, 3, 7, 9, 1, 6, 4, 7, 4, 6, 5, 8, 1, 3, 2, 9, _
2, 8, 3, 7, 5, 6, 4, 1, 9, 5, 1, 4, 2, 9, 8, 3, 6, 7, 7, 9, 6, 1, 3, 4, 5, 2, 8, 1, 6, 7, 9, 8, 5, 2, 4, 3, 4, 2, 5, 3, 7, 1, 9, 8, 6, 8, 3, 9, 6, 4, 2, 1, 7, 5, 3, 7, 2, 4, 6, 9, 8, 5, 1, 9, 4, 8, 5, 1, 7, 6, 3, 2, 6, 5, 1, 8, 2, 3, 7, 9, 4, _
3, 8, 2, 1, 5, 7, 9, 4, 6, 5, 6, 4, 9, 2, 8, 3, 1, 7, 9, 7, 1, 6, 4, 3, 5, 2, 8, 4, 1, 3, 8, 7, 6, 2, 5, 9, 8, 2, 5, 4, 1, 9, 6, 7, 3, 6, 9, 7, 5, 3, 2, 4, 8, 1, 1, 5, 6, 3, 8, 4, 7, 9, 2, 7, 4, 9, 2, 6, 1, 8, 3, 5, 2, 3, 8, 7, 9, 5, 1, 6, 4, _
4, 6, 5, 1, 2, 3, 7, 9, 8, 2, 8, 3, 9, 7, 6, 5, 4, 1, 9, 7, 1, 8, 4, 5, 3, 6, 2, 8, 2, 4, 5, 9, 7, 1, 3, 6, 5, 1, 7, 6, 3, 2, 9, 8, 4, 6, 3, 9, 4, 1, 8, 2, 5, 7, 1, 9, 2, 3, 6, 4, 8, 7, 5, 3, 4, 8, 7, 5, 1, 6, 2, 9, 7, 5, 6, 2, 8, 9, 4, 1, 3, _
6, 2, 8, 9, 4, 1, 3, 5, 7, 7, 9, 5, 8, 6, 3, 4, 2, 1, 3, 4, 1, 7, 2, 5, 9, 8, 6, 8, 1, 7, 4, 3, 2, 5, 6, 9, 4, 5, 6, 1, 8, 9, 2, 7, 3, 9, 3, 2, 5, 7, 6, 1, 4, 8, 1, 8, 4, 3, 5, 7, 6, 9, 2, 5, 6, 9, 2, 1, 8, 7, 3, 4, 2, 7, 3, 6, 9, 4, 8, 1, 5, _
3, 1, 2, 6, 4, 9, 8, 7, 5, 5, 7, 4, 1, 2, 8, 3, 6, 9, 8, 6, 9, 7, 5, 3, 2, 1, 4, 4, 2, 1, 3, 6, 5, 7, 9, 8, 6, 5, 3, 8, 9, 7, 4, 2, 1, 7, 9, 8, 4, 1, 2, 6, 5, 3, 1, 4, 7, 9, 3, 6, 5, 8, 2, 9, 8, 5, 2, 7, 4, 1, 3, 6, 2, 3, 6, 5, 8, 1, 9, 4, 7, _
4, 9, 5, 1, 2, 8, 7, 6, 3, 3, 1, 7, 9, 5, 6, 8, 4, 2, 8, 6, 2, 4, 3, 7, 1, 9, 5, 2, 8, 4, 6, 9, 3, 5, 7, 1, 9, 7, 1, 2, 8, 5, 6, 3, 4, 6, 5, 3, 7, 1, 4, 2, 8, 9, 7, 4, 9, 5, 6, 1, 3, 2, 8, 5, 2, 8, 3, 7, 9, 4, 1, 6, 1, 3, 6, 8, 4, 2, 9, 5, 7, _
8, 6, 5, 7, 2, 4, 9, 1, 3, 2, 9, 7, 3, 1, 6, 4, 8, 5, 4, 3, 1, 8, 9, 5, 7, 2, 6, 6, 1, 2, 9, 3, 8, 5, 7, 4, 9, 7, 3, 5, 4, 1, 2, 6, 8, 5, 8, 4, 6, 7, 2, 3, 9, 1, 3, 4, 8, 2, 6, 7, 1, 5, 9, 1, 2, 6, 4, 5, 9, 8, 3, 7, 7, 5, 9, 1, 8, 3, 6, 4, 2, _
5, 1, 3, 4, 2, 9, 6, 7, 8, 2, 8, 4, 7, 5, 6, 3, 1, 9, 6, 7, 9, 3, 8, 1, 2, 5, 4, 1, 3, 5, 9, 6, 2, 8, 4, 7, 8, 2, 7, 1, 3, 4, 9, 6, 5, 4, 9, 6, 8, 7, 5, 1, 2, 3, 3, 6, 8, 5, 1, 7, 4, 9, 2, 7, 4, 2, 6, 9, 8, 5, 3, 1, 9, 5, 1, 2, 4, 3, 7, 8, 6, _
3, 8, 2, 5, 9, 7, 4, 1, 6, 6, 4, 9, 2, 1, 8, 7, 3, 5, 7, 1, 5, 6, 3, 4, 2, 8, 9, 1, 7, 3, 4, 8, 9, 5, 6, 2, 4, 5, 6, 3, 2, 1, 8, 9, 7, 9, 2, 8, 7, 6, 5, 3, 4, 1, 8, 6, 7, 1, 5, 3, 9, 2, 4, 2, 9, 4, 8, 7, 6, 1, 5, 3, 5, 3, 1, 9, 4, 2, 6, 7, 8, _
1, 6, 7, 2, 5, 9, 3, 4, 8, 4, 2, 8, 3, 6, 1, 5, 9, 7, 5, 9, 3, 7, 8, 4, 6, 2, 1, 2, 8, 5, 6, 9, 3, 1, 7, 4, 3, 1, 6, 4, 7, 8, 2, 5, 9, 9, 7, 4, 5, 1, 2, 8, 3, 6, 8, 5, 1, 9, 3, 7, 4, 6, 2, 6, 4, 9, 1, 2, 5, 7, 8, 3, 7, 3, 2, 8, 4, 6, 9, 1, 5, _
4, 6, 2, 1, 5, 3, 8, 9, 7, 8, 9, 1, 6, 4, 7, 2, 5, 3, 5, 3, 7, 9, 8, 2, 6, 4, 1, 9, 2, 8, 4, 3, 1, 7, 6, 5, 6, 1, 5, 7, 9, 8, 4, 3, 2, 7, 4, 3, 5, 2, 6, 9, 1, 8, 2, 8, 6, 3, 1, 4, 5, 7, 9, 3, 5, 4, 2, 7, 9, 1, 8, 6, 1, 7, 9, 8, 6, 5, 3, 2, 4, _
7, 1, 2, 9, 5, 3, 8, 6, 4, 3, 9, 8, 6, 7, 4, 1, 2, 5, 5, 6, 4, 8, 1, 2, 7, 3, 9, 2, 7, 3, 5, 4, 8, 9, 1, 6, 9, 4, 5, 3, 6, 1, 2, 8, 7, 6, 8, 1, 2, 9, 7, 5, 4, 3, 1, 5, 7, 4, 2, 6, 3, 9, 8, 4, 3, 9, 1, 8, 5, 6, 7, 2, 8, 2, 6, 7, 3, 9, 4, 5, 1, _
8, 1, 3, 6, 5, 2, 4, 7, 9, 5, 2, 4, 8, 9, 7, 6, 3, 1, 7, 6, 9, 1, 4, 3, 5, 2, 8, 3, 4, 2, 5, 7, 9, 8, 1, 6, 9, 8, 7, 2, 1, 6, 3, 5, 4, 6, 5, 1, 3, 8, 4, 2, 9, 7, 1, 9, 6, 4, 3, 5, 7, 8, 2, 4, 3, 8, 7, 2, 1, 9, 6, 5, 2, 7, 5, 9, 6, 8, 1, 4, 3, _
7, 8, 3, 2, 9, 5, 6, 1, 4, 2, 6, 5, 3, 1, 4, 9, 7, 8, 9, 4, 1, 7, 6, 8, 3, 2, 5, 8, 7, 6, 4, 5, 2, 1, 3, 9, 5, 1, 9, 6, 3, 7, 8, 4, 2, 4, 3, 2, 1, 8, 9, 5, 6, 7, 1, 2, 8, 9, 4, 3, 7, 5, 6, 3, 9, 7, 5, 2, 6, 4, 8, 1, 6, 5, 4, 8, 7, 1, 2, 9, 3, _
1, 7, 3, 5, 4, 2, 8, 6, 9, 5, 9, 6, 7, 1, 8, 2, 4, 3, 4, 2, 8, 3, 6, 9, 7, 5, 1, 2, 4, 1, 9, 5, 6, 3, 7, 8, 8, 6, 5, 1, 7, 3, 4, 9, 2, 9, 3, 7, 2, 8, 4, 5, 1, 6, 7, 5, 2, 6, 3, 1, 9, 8, 4, 6, 8, 9, 4, 2, 5, 1, 3, 7, 3, 1, 4, 8, 9, 7, 6, 2, 5, _
9, 4, 5, 1, 3, 8, 2, 6, 7, 3, 8, 6, 5, 2, 7, 1, 4, 9, 7, 1, 2, 6, 9, 4, 8, 5, 3, 8, 7, 9, 2, 6, 3, 4, 1, 5, 4, 2, 3, 8, 1, 5, 7, 9, 6, 5, 6, 1, 7, 4, 9, 3, 2, 8, 1, 5, 8, 4, 7, 6, 9, 3, 2, 6, 9, 4, 3, 8, 2, 5, 7, 1, 2, 3, 7, 9, 5, 1, 6, 8, 4)

s = 0
Randomize
For i = 1 To Val(Text2.Text)
    While (s = 0)
        a = Int(Rnd * (80)) + 1
        If (x(a) = 0) Then
            x(a) = 1
            s = 1
        End If
    Wend
    s = 0
Next

index = index + 1

For i = 0 To 8
    For j = 0 To 8
        If (x(i * 9 + j) = 0) Then
            Text1(i * 9 + j).Text = t((i * 9 + j) + (index Mod 20) * 81)
            Text1(i * 9 + j).ForeColor = &H80&
            Text1(i * 9 + j).Locked = True
        Else
            Text1(i * 9 + j).Text = ""
            Text1(i * 9 + j).Locked = False
            Text1(i * 9 + j).ForeColor = &H0&
        End If
    Next
Next
End Sub

Private Sub Command2_Click()
Dim win, temp1, temp2, flag As Integer
Dim i, j, k As Integer
win = 1

For i = 0 To 8
    temp1 = 0
    temp2 = 1
    For j = 0 To 8
        temp1 = temp1 + Val(Text1(i * 9 + j).Text)
        temp2 = temp2 * Val(Text1(i * 9 + j).Text)
    Next
    If (temp1 = 45 And temp2 = 362880) Then
        win = win
    Else
        win = 0
    End If
Next

For j = 0 To 8
    temp1 = 0
    temp2 = 1
    For i = 0 To 8
        temp1 = temp1 + Val(Text1(i * 9 + j).Text)
        temp2 = temp2 * Val(Text1(i * 9 + j).Text)
    Next
    If (temp1 = 45 And temp2 = 362880) Then
        win = win
    Else
        win = 0
    End If
Next

For i = 0 To 8
    temp1 = 0
    temp2 = 1
    For j = 0 To 2
        For k = 0 To 2
            temp1 = temp1 + Val(Text1(i * 9 + j * 3 + k).Text)
            temp2 = temp2 * Val(Text1(i * 9 + j * 3 + k).Text)
        Next
    Next
    If (temp1 = 45 And temp2 = 362880) Then
        win = win
    Else
        win = 0
    End If
Next

If (win = 1) Then
    For i = 0 To 8
        For j = 0 To 8
            Text1(i * 9 + j).ForeColor = &HFF00FF
            Text1(i * 9 + j).Locked = True
        Next
    Next
    MsgBox "恭喜您，答案正确！", 0, "答案正确"
Else
    MsgBox "您的答案不正确，请重试！", 16, "答案不正确"
End If
End Sub

Private Sub Command3_Click()
t = Array( _
9, 7, 6, 8, 5, 4, 2, 3, 1, 2, 5, 3, 6, 9, 1, 7, 8, 4, 4, 1, 8, 3, 7, 2, 5, 6, 9, 7, 6, 1, 4, 2, 3, 9, 5, 8, 5, 8, 4, 9, 6, 7, 1, 2, 3, 3, 2, 9, 1, 8, 5, 4, 7, 6, 6, 3, 5, 7, 4, 9, 8, 1, 2, 1, 4, 2, 5, 3, 8, 6, 9, 7, 8, 9, 7, 2, 1, 6, 3, 4, 5, _
4, 1, 6, 5, 2, 8, 3, 7, 9, 2, 8, 9, 3, 7, 6, 4, 1, 5, 7, 3, 5, 9, 1, 4, 6, 8, 2, 9, 4, 7, 1, 8, 3, 2, 5, 6, 1, 5, 3, 2, 6, 9, 7, 4, 8, 8, 6, 2, 7, 4, 5, 1, 9, 3, 6, 9, 1, 4, 5, 2, 8, 3, 7, 3, 7, 8, 6, 9, 1, 5, 2, 4, 5, 2, 4, 8, 3, 7, 9, 6, 1, _
3, 6, 8, 2, 1, 9, 5, 4, 7, 1, 9, 2, 5, 4, 7, 8, 3, 6, 4, 7, 5, 6, 8, 3, 9, 2, 1, 6, 5, 3, 8, 7, 1, 2, 9, 4, 2, 1, 4, 9, 6, 5, 3, 7, 8, 7, 8, 9, 3, 2, 4, 6, 1, 5, 9, 3, 1, 4, 5, 8, 7, 6, 2, 5, 4, 6, 7, 3, 2, 1, 8, 9, 8, 2, 7, 1, 9, 6, 4, 5, 3, _
5, 9, 8, 6, 1, 7, 2, 4, 3, 1, 3, 7, 8, 2, 4, 6, 9, 5, 2, 6, 4, 9, 3, 5, 7, 1, 8, 4, 5, 2, 7, 6, 8, 9, 3, 1, 3, 8, 1, 4, 9, 2, 5, 7, 6, 6, 7, 9, 1, 5, 3, 4, 8, 2, 9, 1, 3, 2, 4, 6, 8, 5, 7, 8, 2, 5, 3, 7, 9, 1, 6, 4, 7, 4, 6, 5, 8, 1, 3, 2, 9, _
2, 8, 3, 7, 5, 6, 4, 1, 9, 5, 1, 4, 2, 9, 8, 3, 6, 7, 7, 9, 6, 1, 3, 4, 5, 2, 8, 1, 6, 7, 9, 8, 5, 2, 4, 3, 4, 2, 5, 3, 7, 1, 9, 8, 6, 8, 3, 9, 6, 4, 2, 1, 7, 5, 3, 7, 2, 4, 6, 9, 8, 5, 1, 9, 4, 8, 5, 1, 7, 6, 3, 2, 6, 5, 1, 8, 2, 3, 7, 9, 4, _
3, 8, 2, 1, 5, 7, 9, 4, 6, 5, 6, 4, 9, 2, 8, 3, 1, 7, 9, 7, 1, 6, 4, 3, 5, 2, 8, 4, 1, 3, 8, 7, 6, 2, 5, 9, 8, 2, 5, 4, 1, 9, 6, 7, 3, 6, 9, 7, 5, 3, 2, 4, 8, 1, 1, 5, 6, 3, 8, 4, 7, 9, 2, 7, 4, 9, 2, 6, 1, 8, 3, 5, 2, 3, 8, 7, 9, 5, 1, 6, 4, _
4, 6, 5, 1, 2, 3, 7, 9, 8, 2, 8, 3, 9, 7, 6, 5, 4, 1, 9, 7, 1, 8, 4, 5, 3, 6, 2, 8, 2, 4, 5, 9, 7, 1, 3, 6, 5, 1, 7, 6, 3, 2, 9, 8, 4, 6, 3, 9, 4, 1, 8, 2, 5, 7, 1, 9, 2, 3, 6, 4, 8, 7, 5, 3, 4, 8, 7, 5, 1, 6, 2, 9, 7, 5, 6, 2, 8, 9, 4, 1, 3, _
6, 2, 8, 9, 4, 1, 3, 5, 7, 7, 9, 5, 8, 6, 3, 4, 2, 1, 3, 4, 1, 7, 2, 5, 9, 8, 6, 8, 1, 7, 4, 3, 2, 5, 6, 9, 4, 5, 6, 1, 8, 9, 2, 7, 3, 9, 3, 2, 5, 7, 6, 1, 4, 8, 1, 8, 4, 3, 5, 7, 6, 9, 2, 5, 6, 9, 2, 1, 8, 7, 3, 4, 2, 7, 3, 6, 9, 4, 8, 1, 5, _
3, 1, 2, 6, 4, 9, 8, 7, 5, 5, 7, 4, 1, 2, 8, 3, 6, 9, 8, 6, 9, 7, 5, 3, 2, 1, 4, 4, 2, 1, 3, 6, 5, 7, 9, 8, 6, 5, 3, 8, 9, 7, 4, 2, 1, 7, 9, 8, 4, 1, 2, 6, 5, 3, 1, 4, 7, 9, 3, 6, 5, 8, 2, 9, 8, 5, 2, 7, 4, 1, 3, 6, 2, 3, 6, 5, 8, 1, 9, 4, 7, _
4, 9, 5, 1, 2, 8, 7, 6, 3, 3, 1, 7, 9, 5, 6, 8, 4, 2, 8, 6, 2, 4, 3, 7, 1, 9, 5, 2, 8, 4, 6, 9, 3, 5, 7, 1, 9, 7, 1, 2, 8, 5, 6, 3, 4, 6, 5, 3, 7, 1, 4, 2, 8, 9, 7, 4, 9, 5, 6, 1, 3, 2, 8, 5, 2, 8, 3, 7, 9, 4, 1, 6, 1, 3, 6, 8, 4, 2, 9, 5, 7, _
8, 6, 5, 7, 2, 4, 9, 1, 3, 2, 9, 7, 3, 1, 6, 4, 8, 5, 4, 3, 1, 8, 9, 5, 7, 2, 6, 6, 1, 2, 9, 3, 8, 5, 7, 4, 9, 7, 3, 5, 4, 1, 2, 6, 8, 5, 8, 4, 6, 7, 2, 3, 9, 1, 3, 4, 8, 2, 6, 7, 1, 5, 9, 1, 2, 6, 4, 5, 9, 8, 3, 7, 7, 5, 9, 1, 8, 3, 6, 4, 2, _
5, 1, 3, 4, 2, 9, 6, 7, 8, 2, 8, 4, 7, 5, 6, 3, 1, 9, 6, 7, 9, 3, 8, 1, 2, 5, 4, 1, 3, 5, 9, 6, 2, 8, 4, 7, 8, 2, 7, 1, 3, 4, 9, 6, 5, 4, 9, 6, 8, 7, 5, 1, 2, 3, 3, 6, 8, 5, 1, 7, 4, 9, 2, 7, 4, 2, 6, 9, 8, 5, 3, 1, 9, 5, 1, 2, 4, 3, 7, 8, 6, _
3, 8, 2, 5, 9, 7, 4, 1, 6, 6, 4, 9, 2, 1, 8, 7, 3, 5, 7, 1, 5, 6, 3, 4, 2, 8, 9, 1, 7, 3, 4, 8, 9, 5, 6, 2, 4, 5, 6, 3, 2, 1, 8, 9, 7, 9, 2, 8, 7, 6, 5, 3, 4, 1, 8, 6, 7, 1, 5, 3, 9, 2, 4, 2, 9, 4, 8, 7, 6, 1, 5, 3, 5, 3, 1, 9, 4, 2, 6, 7, 8, _
1, 6, 7, 2, 5, 9, 3, 4, 8, 4, 2, 8, 3, 6, 1, 5, 9, 7, 5, 9, 3, 7, 8, 4, 6, 2, 1, 2, 8, 5, 6, 9, 3, 1, 7, 4, 3, 1, 6, 4, 7, 8, 2, 5, 9, 9, 7, 4, 5, 1, 2, 8, 3, 6, 8, 5, 1, 9, 3, 7, 4, 6, 2, 6, 4, 9, 1, 2, 5, 7, 8, 3, 7, 3, 2, 8, 4, 6, 9, 1, 5, _
4, 6, 2, 1, 5, 3, 8, 9, 7, 8, 9, 1, 6, 4, 7, 2, 5, 3, 5, 3, 7, 9, 8, 2, 6, 4, 1, 9, 2, 8, 4, 3, 1, 7, 6, 5, 6, 1, 5, 7, 9, 8, 4, 3, 2, 7, 4, 3, 5, 2, 6, 9, 1, 8, 2, 8, 6, 3, 1, 4, 5, 7, 9, 3, 5, 4, 2, 7, 9, 1, 8, 6, 1, 7, 9, 8, 6, 5, 3, 2, 4, _
7, 1, 2, 9, 5, 3, 8, 6, 4, 3, 9, 8, 6, 7, 4, 1, 2, 5, 5, 6, 4, 8, 1, 2, 7, 3, 9, 2, 7, 3, 5, 4, 8, 9, 1, 6, 9, 4, 5, 3, 6, 1, 2, 8, 7, 6, 8, 1, 2, 9, 7, 5, 4, 3, 1, 5, 7, 4, 2, 6, 3, 9, 8, 4, 3, 9, 1, 8, 5, 6, 7, 2, 8, 2, 6, 7, 3, 9, 4, 5, 1, _
8, 1, 3, 6, 5, 2, 4, 7, 9, 5, 2, 4, 8, 9, 7, 6, 3, 1, 7, 6, 9, 1, 4, 3, 5, 2, 8, 3, 4, 2, 5, 7, 9, 8, 1, 6, 9, 8, 7, 2, 1, 6, 3, 5, 4, 6, 5, 1, 3, 8, 4, 2, 9, 7, 1, 9, 6, 4, 3, 5, 7, 8, 2, 4, 3, 8, 7, 2, 1, 9, 6, 5, 2, 7, 5, 9, 6, 8, 1, 4, 3, _
7, 8, 3, 2, 9, 5, 6, 1, 4, 2, 6, 5, 3, 1, 4, 9, 7, 8, 9, 4, 1, 7, 6, 8, 3, 2, 5, 8, 7, 6, 4, 5, 2, 1, 3, 9, 5, 1, 9, 6, 3, 7, 8, 4, 2, 4, 3, 2, 1, 8, 9, 5, 6, 7, 1, 2, 8, 9, 4, 3, 7, 5, 6, 3, 9, 7, 5, 2, 6, 4, 8, 1, 6, 5, 4, 8, 7, 1, 2, 9, 3, _
1, 7, 3, 5, 4, 2, 8, 6, 9, 5, 9, 6, 7, 1, 8, 2, 4, 3, 4, 2, 8, 3, 6, 9, 7, 5, 1, 2, 4, 1, 9, 5, 6, 3, 7, 8, 8, 6, 5, 1, 7, 3, 4, 9, 2, 9, 3, 7, 2, 8, 4, 5, 1, 6, 7, 5, 2, 6, 3, 1, 9, 8, 4, 6, 8, 9, 4, 2, 5, 1, 3, 7, 3, 1, 4, 8, 9, 7, 6, 2, 5, _
9, 4, 5, 1, 3, 8, 2, 6, 7, 3, 8, 6, 5, 2, 7, 1, 4, 9, 7, 1, 2, 6, 9, 4, 8, 5, 3, 8, 7, 9, 2, 6, 3, 4, 1, 5, 4, 2, 3, 8, 1, 5, 7, 9, 6, 5, 6, 1, 7, 4, 9, 3, 2, 8, 1, 5, 8, 4, 7, 6, 9, 3, 2, 6, 9, 4, 3, 8, 2, 5, 7, 1, 2, 3, 7, 9, 5, 1, 6, 8, 4)

For i = 0 To 8
    For j = 0 To 8
        If (x(i * 9 + j) = 1) Then
            Text1(i * 9 + j).Text = t((i * 9 + j) + (index Mod 20) * 81)
            Text1(i * 9 + j).ForeColor = &HC000&
            Text1(i * 9 + j).Locked = True
        End If
    Next
Next
End Sub

Private Sub Command4_Click()
For i = 0 To 8
    For j = 0 To 8
        Text1(i * 9 + j).Text = ""
        Text1(i * 9 + j).Locked = True
    Next
Next
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Form_Load()
index = -1
For i = 0 To 8
    For j = 0 To 8
        Text1(i * 9 + j).Text = ""
        Text1(i * 9 + j).Locked = True
    Next
Next
End Sub
