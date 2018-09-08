VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "okba Ver.110"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   1665
   ClientWidth     =   23565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   10185
   ScaleWidth      =   23565
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command31 
      Caption         =   "ango"
      Height          =   495
      Left            =   9960
      TabIndex        =   179
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command30 
      Caption         =   "dia"
      Height          =   495
      Left            =   2160
      TabIndex        =   178
      Top             =   1920
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1860
      Left            =   4200
      TabIndex        =   177
      Top             =   2760
      Width           =   5295
   End
   Begin VB.CommandButton Command29 
      Caption         =   "test"
      Height          =   405
      Left            =   3240
      TabIndex        =   176
      Top             =   1920
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picpic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   7000
      Left            =   14880
      ScaleHeight     =   6945
      ScaleWidth      =   8355
      TabIndex        =   130
      Top             =   240
      Width           =   8415
   End
   Begin VB.CheckBox chkNankan 
      Caption         =   "Check1"
      Height          =   375
      Left            =   3600
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   175
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command28 
      Caption         =   "renpai"
      Height          =   495
      Left            =   10320
      TabIndex        =   174
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command27 
      Caption         =   "rate"
      Height          =   495
      Left            =   10320
      TabIndex        =   173
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command26 
      Caption         =   "recipe2"
      Height          =   525
      Left            =   3240
      TabIndex        =   171
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtSelRecipe 
      Height          =   270
      Left            =   0
      TabIndex        =   169
      Text            =   "01-66_05-59"
      Top             =   9840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtRecipe 
      Height          =   375
      Left            =   15720
      TabIndex        =   167
      Text            =   "_01-66_05-59"
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command25 
      Caption         =   "recipe"
      Height          =   405
      Left            =   1560
      TabIndex        =   166
      Top             =   9720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command24 
      Caption         =   "today"
      Height          =   525
      Left            =   2280
      TabIndex        =   165
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command23 
      Caption         =   "DB2"
      Height          =   525
      Left            =   1560
      TabIndex        =   164
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox chkTF 
      Caption         =   "Check1"
      Height          =   375
      Left            =   840
      TabIndex        =   163
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAll 
      Height          =   270
      Left            =   3960
      TabIndex        =   162
      Text            =   "Text2"
      Top             =   9840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox areaMD 
      Height          =   285
      Left            =   3240
      TabIndex        =   160
      Text            =   "0716"
      Top             =   4560
      Width           =   675
   End
   Begin VB.CommandButton Command22 
      Caption         =   "IE run"
      Height          =   645
      Left            =   10800
      TabIndex        =   157
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton Command21 
      Caption         =   "CLR"
      Height          =   2295
      Left            =   9840
      TabIndex        =   156
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command20 
      Caption         =   "check"
      Height          =   645
      Left            =   13080
      TabIndex        =   155
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtChk 
      Height          =   375
      Left            =   11760
      TabIndex        =   154
      Text            =   "01-66_05-59"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtVol 
      Height          =   375
      Left            =   11880
      TabIndex        =   153
      Text            =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command19 
      Caption         =   "harai"
      Height          =   615
      Left            =   7680
      TabIndex        =   152
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00FFFF00&
      Caption         =   "XP"
      Height          =   615
      Left            =   13320
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   151
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtStart 
      Height          =   375
      Left            =   11880
      TabIndex        =   150
      Text            =   "0"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtRes 
      Height          =   2445
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   3  '両方
      TabIndex        =   149
      Top             =   120
      Width           =   4395
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FFFF00&
      Caption         =   "tblAnal"
      Height          =   615
      Left            =   10800
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   148
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Caption         =   "bin2DB2all"
      Height          =   615
      Left            =   8880
      TabIndex        =   147
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00008080&
      Caption         =   "bin2DB2"
      Height          =   615
      Left            =   8880
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   146
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCnd 
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   145
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtCnd 
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   143
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtCnd 
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   141
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command14 
      Caption         =   "choice"
      Height          =   1335
      Left            =   4200
      TabIndex        =   139
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtBin 
      Height          =   375
      Left            =   360
      TabIndex        =   138
      Text            =   "ptn_cmb02001-03000.dat"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command13 
      Caption         =   "bin2DB"
      Height          =   615
      Left            =   8880
      TabIndex        =   137
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "read bin"
      Height          =   615
      Left            =   6600
      TabIndex        =   136
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "makePtn"
      Height          =   615
      Left            =   5520
      TabIndex        =   135
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "getPtn"
      Height          =   615
      Left            =   4680
      TabIndex        =   134
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "cmb"
      Height          =   375
      Left            =   2280
      TabIndex        =   133
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ptn"
      Height          =   615
      Left            =   2280
      TabIndex        =   132
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "LINE"
      Height          =   615
      Left            =   2280
      TabIndex        =   131
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "cmpiListCheck"
      Height          =   525
      Left            =   2400
      TabIndex        =   129
      Top             =   9600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "csv2db"
      Height          =   525
      Left            =   600
      TabIndex        =   128
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "html2csv"
      Height          =   525
      Left            =   600
      TabIndex        =   127
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   2895
   End
   Begin VB.TextBox areaY 
      Height          =   285
      Left            =   2400
      TabIndex        =   126
      Text            =   "2017"
      Top             =   4560
      Width           =   675
   End
   Begin VB.CommandButton Command3 
      Caption         =   "スクレイピング"
      Height          =   525
      Left            =   600
      TabIndex        =   125
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IE run"
      Height          =   645
      Left            =   120
      TabIndex        =   124
      Top             =   6000
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   2085
      Left            =   9720
      MultiLine       =   -1  'True
      ScrollBars      =   3  '両方
      TabIndex        =   23
      Top             =   2760
      Width           =   4875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "複勝回収率１２０％"
      Enabled         =   0   'False
      Height          =   525
      Left            =   120
      TabIndex        =   22
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmd 
      Caption         =   "on off"
      Height          =   1335
      Index           =   2
      Left            =   4440
      TabIndex        =   4
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   16455
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   17
         Left            =   12120
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   21
         Top             =   360
         Width           =   315
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   16
         Left            =   11400
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   15
         Left            =   10680
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   14
         Left            =   10080
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   18
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   13
         Left            =   9360
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   12
         Left            =   8640
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   11
         Left            =   7920
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   10
         Left            =   7200
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   9
         Left            =   6480
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   8
         Left            =   5880
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   7
         Left            =   5160
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   6
         Left            =   4560
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   5
         Left            =   3720
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   4
         Left            =   3000
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   3
         Left            =   2280
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   2
         Left            =   1560
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   1
         Left            =   840
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000000FF&
         Caption         =   "99"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   0
         Left            =   240
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   99
         Left            =   3960
         TabIndex        =   123
         Top             =   120
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   98
         Left            =   4440
         TabIndex        =   122
         Top             =   60
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   97
         Left            =   4860
         TabIndex        =   121
         Top             =   60
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   96
         Left            =   5340
         TabIndex        =   120
         Top             =   0
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   95
         Left            =   5820
         TabIndex        =   119
         Top             =   90
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   94
         Left            =   6300
         TabIndex        =   118
         Top             =   30
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   93
         Left            =   6840
         TabIndex        =   117
         Top             =   90
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   92
         Left            =   7320
         TabIndex        =   116
         Top             =   30
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   91
         Left            =   7830
         TabIndex        =   115
         Top             =   90
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   90
         Left            =   8310
         TabIndex        =   114
         Top             =   30
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   89
         Left            =   3360
         TabIndex        =   113
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   88
         Left            =   3840
         TabIndex        =   112
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   87
         Left            =   4350
         TabIndex        =   111
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   86
         Left            =   4830
         TabIndex        =   110
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   85
         Left            =   5280
         TabIndex        =   109
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   84
         Left            =   5760
         TabIndex        =   108
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   83
         Left            =   6360
         TabIndex        =   107
         Top             =   1710
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   82
         Left            =   6840
         TabIndex        =   106
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   81
         Left            =   7320
         TabIndex        =   105
         Top             =   1770
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   80
         Left            =   7800
         TabIndex        =   104
         Top             =   1710
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   79
         Left            =   8220
         TabIndex        =   103
         Top             =   1710
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   78
         Left            =   8700
         TabIndex        =   102
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   77
         Left            =   9180
         TabIndex        =   101
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   76
         Left            =   9660
         TabIndex        =   100
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   75
         Left            =   10200
         TabIndex        =   99
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   74
         Left            =   10680
         TabIndex        =   98
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   73
         Left            =   11190
         TabIndex        =   97
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   72
         Left            =   11670
         TabIndex        =   96
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   71
         Left            =   1950
         TabIndex        =   95
         Top             =   1710
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   70
         Left            =   2430
         TabIndex        =   94
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   69
         Left            =   2940
         TabIndex        =   93
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   68
         Left            =   3420
         TabIndex        =   92
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   67
         Left            =   3870
         TabIndex        =   91
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   66
         Left            =   4350
         TabIndex        =   90
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   65
         Left            =   4950
         TabIndex        =   89
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   64
         Left            =   5430
         TabIndex        =   88
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   63
         Left            =   5910
         TabIndex        =   87
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   62
         Left            =   6390
         TabIndex        =   86
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   61
         Left            =   6810
         TabIndex        =   85
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   60
         Left            =   7290
         TabIndex        =   84
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   59
         Left            =   7770
         TabIndex        =   83
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   58
         Left            =   8250
         TabIndex        =   82
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   57
         Left            =   8790
         TabIndex        =   81
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   56
         Left            =   9270
         TabIndex        =   80
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   55
         Left            =   9780
         TabIndex        =   79
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   54
         Left            =   10260
         TabIndex        =   78
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   53
         Left            =   1020
         TabIndex        =   77
         Top             =   1710
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   52
         Left            =   1500
         TabIndex        =   76
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   51
         Left            =   2010
         TabIndex        =   75
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   50
         Left            =   2490
         TabIndex        =   74
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   49
         Left            =   2940
         TabIndex        =   73
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   48
         Left            =   3420
         TabIndex        =   72
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   47
         Left            =   4020
         TabIndex        =   71
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   46
         Left            =   4500
         TabIndex        =   70
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   45
         Left            =   4980
         TabIndex        =   69
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   44
         Left            =   5460
         TabIndex        =   68
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   43
         Left            =   5880
         TabIndex        =   67
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   42
         Left            =   6360
         TabIndex        =   66
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   41
         Left            =   6840
         TabIndex        =   65
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   40
         Left            =   7320
         TabIndex        =   64
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   39
         Left            =   7860
         TabIndex        =   63
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   38
         Left            =   8340
         TabIndex        =   62
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   37
         Left            =   8850
         TabIndex        =   61
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   36
         Left            =   9330
         TabIndex        =   60
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   35
         Left            =   420
         TabIndex        =   59
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   34
         Left            =   900
         TabIndex        =   58
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   33
         Left            =   1410
         TabIndex        =   57
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   32
         Left            =   1890
         TabIndex        =   56
         Top             =   1530
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   31
         Left            =   2340
         TabIndex        =   55
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   30
         Left            =   2820
         TabIndex        =   54
         Top             =   1530
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   29
         Left            =   3420
         TabIndex        =   53
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   28
         Left            =   3900
         TabIndex        =   52
         Top             =   1500
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   27
         Left            =   4380
         TabIndex        =   51
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   26
         Left            =   4860
         TabIndex        =   50
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   25
         Left            =   5280
         TabIndex        =   49
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   24
         Left            =   5760
         TabIndex        =   48
         Top             =   1500
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   23
         Left            =   6240
         TabIndex        =   47
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   22
         Left            =   6720
         TabIndex        =   46
         Top             =   1530
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   21
         Left            =   7260
         TabIndex        =   45
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   20
         Left            =   7740
         TabIndex        =   44
         Top             =   1530
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   19
         Left            =   8250
         TabIndex        =   43
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   18
         Left            =   8730
         TabIndex        =   42
         Top             =   1530
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   17
         Left            =   8610
         TabIndex        =   41
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   16
         Left            =   8130
         TabIndex        =   40
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   15
         Left            =   7620
         TabIndex        =   39
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   14
         Left            =   7140
         TabIndex        =   38
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   13
         Left            =   6600
         TabIndex        =   37
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   12
         Left            =   6120
         TabIndex        =   36
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   11
         Left            =   5640
         TabIndex        =   35
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   5160
         TabIndex        =   34
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   4740
         TabIndex        =   33
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   4260
         TabIndex        =   32
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   3780
         TabIndex        =   31
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   3300
         TabIndex        =   30
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   2700
         TabIndex        =   29
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   2220
         TabIndex        =   28
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   1770
         TabIndex        =   27
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   1290
         TabIndex        =   26
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   780
         TabIndex        =   25
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BorderStyle     =   1  '実線
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   24
         Top             =   1710
         Width           =   435
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "->"
      Height          =   1335
      Index           =   1
      Left            =   5880
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "<-"
      Height          =   1335
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "recipe"
      Height          =   255
      Index           =   8
      Left            =   9720
      TabIndex        =   172
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "recipe"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   170
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "List1"
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   168
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "monthday"
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   161
      Top             =   7800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "year"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   159
      Top             =   7440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "pattern"
      Height          =   375
      Index           =   3
      Left            =   10920
      TabIndex        =   158
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "race count"
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   144
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "hit rate"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   142
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "return rate"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   140
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objNonCode As Object    ' 文字コード判定/変換オブジェクト
'Private Const PATH_DB = ".\data__\dmc.mdb"        'テスト版
Private Const PATH_DB = ".\dmcAnal.mdb"        '半製品版
Private ie As SHDocVw.InternetExplorer

Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim mSel As Integer         '1(コンピ人気順位)～               賭ける馬
Dim mPos(99) As Integer     'mPos(0)がコンピ９０
Dim mChoise(17) As Integer  'mChoise(aNinki - 1)        選択されてる
Private myURL         As String
Private gYmd As String

Private gYmdPlaceRace As String
Private gStr As String
Private gYears() As String
Private gUrlYear() As String
Private gDnmYear() As String
Private gDnmUrlYear() As String
Private gDay() As String
Private gDayFmt() As String
Private gUrlDay() As String
Private gPosDay() As String     '開催場所
Private gPosDayCd() As String     '開催場所コード 楽天
Private gPosDayDbCd() As String     '開催場所コード データベース
Private gCmpDay() As String     'コンピ指数
Private gWk As String
Private gRace() As String
Private gDenmaRace() As String  '出走表
Private gResRace() As String    '結果
Private gUmaban() As String
Private gBamei() As String
Private gUmaCD() As String
Private gCmp() As String
Private gFukuMny() As String
Private gFukuNum() As String

Private gDnmDay() As String
Private gDnmDayFmt() As String
Private gDnmUrlDay() As String
Private gDnmPosDay() As String     '開催場所
Private gDnmPosDayCd() As String     '開催場所コード 楽天
Private gDnmPosDayDbCd() As String     '開催場所コード データベース

Const M_HABA = 315

Private Sub Compi2TxtNankan(src As String, file As String, wfile As String)
Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
Dim strResult As String '置換後の文字列
Dim Matches
Dim Match
Dim fn As Integer
Dim wfn As Integer
Dim lCnt As Integer
Dim data() As String
Dim wk As String
Dim wk2 As String
Dim wkPrt As String
Dim pos As Long
Dim phase As Long
Dim raceNo As String
Dim wkRaceNo As String
Dim retstr As String '
Dim nen As String
Dim gatu As String
Dim niti As String
Dim basho As String
Dim cmpininki As String
Dim cmpidata(12, 20) As String      'ninki,value
Dim umaban As Integer
Dim value As String
Dim idx As Integer
Dim backup As String
Dim wakCnt As Integer
Dim wakD() As String
Dim plc As Integer
Dim kire As String
Dim smpl As String
Dim wban As String
Dim cmpV As String
Dim wkstr As String
Dim dptr As Integer
Dim cutstr As String
Dim tmp As String

'cmdCmpi.Enabled = False

Dim wkwk As String
'    If optMode(0).value = True Then
         wkwk = "馬番"
'    Else
'         wkwk = "wakuNum"
'    End If


'HTMLファイル(param.)をメモリーに展開
'<<ファイル 開>>
fn = FreeFile
Open src For Input As #fn

'<<ファイル 読>>
lCnt = 0
Do Until EOF(fn)
    Line Input #fn, wk
    ReDim Preserve data(lCnt)
    data(lCnt) = wk
    lCnt = lCnt + 1
Loop

'<<ファイル 閉>>
Close #fn

wfn = FreeFile
Open wfile For Append As #wfn


'<<データ解析>>
'正規表現オブジェクトの宣言
Set objRegExp = New RegExp

With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    
    phase = 0
    For lCnt = 0 To UBound(data)
        
        Select Case phase
        '開催場所、年月日を検索
        '<TH BGCOLOR="#F56403" COLSPAN=31><FONT SIZE=+2>馬番コンピ　　　　</FONT><FONT SIZE=+2>2008年1月19日 1回中山5日目</FONT><FONT SIZE=+2>　　　　枠番コンピ</FONT></TH>
        '<h2 id="contentTit">2012年1月5日　コンピ指数－1回中山1日目</h2>
        Case 0
'             .Pattern = "<FONT SIZE=\+2>20.+日目"
             .Pattern = "<h2 id=""contentTit"">.+日目"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
               pos = Match.FirstIndex       '一致する文字列が見つかった位置
               retstr = Match.value
            Next
            If pos = 0 Then
'                .Pattern = "<font size=""\+2"">20.+日目"
                .Pattern = "contentTit""\>20.+日目"
                Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                '<FONT SIZE=+2>2008年1月20日 1回中山6日目
                '年
                 .Pattern = ">.+年"
                pos = 0
                Set Matches = .Execute(retstr)   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   wk = Match.value
                Next
                nen = Mid$(wk, 2, 4)
                '月
                 .Pattern = "年.+月"
                pos = 0
                Set Matches = .Execute(retstr)   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   wk = Match.value
                Next
                If Len(wk) = 3 Then
                    gatu = Mid$(wk, 2, 1)
                Else
                    gatu = Mid$(wk, 2, 2)
                End If
                '日
                 .Pattern = "月.+日.+コンピ指数"
                pos = 0
                Set Matches = .Execute(retstr)   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   wk = Match.value
                Next
                If Len(wk) = 9 Then
                    niti = Mid$(wk, 2, 1)
                Else
                    niti = Mid$(wk, 2, 2)
                End If
                '開催場所
                 .Pattern = "回.+日目"
                pos = 0
                Set Matches = .Execute(retstr)   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   wk = Match.value
                Next
                Select Case Mid$(wk, 2, 2)
                Case "川崎"
                    basho = "45"
                Case "船橋"
                    basho = "43"
                Case "浦和"
                    basho = "42"
                Case "大井"
                    basho = "44"
                End Select
                
                phase = 1
            End If
        Case 1
            Do
                If InStr(data(lCnt), "</table>") > 0 Then '"【枠番コンピ】") > 0 Then
                    Close #wfn
                    
'                    cmdCmpi.Enabled = True
                    
                    Exit Sub
                End If
                
                .Pattern = "race""\>.+R\<\/span\>"
                 
                pos = 0
                Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   retstr = Match.value
                Next
                
                If pos <> 0 Then
                    If Mid$(retstr, 8, 1) = "R" Then
                        tmp = Mid$(retstr, 7, 1)
                    Else
                        tmp = Mid$(retstr, 7, 2)
                    End If
                    raceNo = tmp
                    
                    wkstr = nen & Format$(gatu, "00") & Format$(niti, "00") & basho & Format$(raceNo, "00")
                    
                    'データ取得
                    pos = InStr(data(lCnt), "<br>指数")
                    cutstr = Mid$(data(lCnt), pos + 6)
'                    cutstr = data(lCnt)
                    Do
                        dptr = InStr(cutstr, "<br>") '"<br />")
                        If dptr = 0 Then
                            Print #wfn, wkstr
                            Exit Do
                        End If
                        
                        If Mid$(cutstr, dptr - 2, 1) = ">" Then
                            tmp = Mid$(cutstr, dptr - 1, 1)
                        Else
                            tmp = Mid$(cutstr, dptr - 2, 2)
                        End If
                        wkstr = wkstr & "," & tmp
                        If Mid$(cutstr, dptr + 5, 1) = "<" Then
                            tmp = Mid$(cutstr, dptr + 4, 1)
                        Else
                            tmp = Mid$(cutstr, dptr + 4, 2)
                        End If
                        
                        wkstr = wkstr & "," & tmp
                        
                        cutstr = Mid$(cutstr, dptr + 10)
                    Loop
                    
                    
                End If
                lCnt = lCnt + 1
            Loop
        End Select
        
    Next lCnt
    
End With

Set objRegExp = Nothing

''テキストファイル(param.)へ出力
'src = file
'fn = FreeFile
'Open src For Append As #fn
'
''<<ファイル 書>>
'
'wk = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho
'
'For idx = 1 To 12
'    wk2 = ""
'    For lCnt = 1 To 20
'        wk2 = wk2 & "," & cmpidata(idx, lCnt)
'    Next lCnt
'
'    wk2 = wk & Format$(idx, "00") & wk2
'    Print #fn, wk2
'
'Next idx
'
''<<ファイル 閉>>
'Close #fn
Close #wfn

'cmdCmpi.Enabled = True

End Sub

Private Function getRate(pArg As String) As Integer
    Dim aDrawDate As String
    Dim aStartDate As Date
    Dim aDays As Long
    Dim win As Long
    Dim total As Long
    Dim mny As Long
    Dim betTotal As Long
    Dim betWin As Long
    Dim resUma(41) As String
    Dim resPay(41) As String
    Dim umauma(5) As String
    Dim aSelUma As String
    Dim aChoiceFlg As Boolean
    Dim aNinki As Integer
    Dim aKariWin As Long
    Dim aKariMny As Long
    Dim aKariBetTotal As Long
    Dim aMsg As String
    Dim aWebMsg As String
    Dim aWk As String
    
    aStartDate = "2007/1/1"
    
    total = 0
    win = 0
    mny = 0
    
    aSelUma = Format$(mSel, "00")
    
    '座標ファイル仮を開く
    Dim aPath As String
    Dim aZa As String
    Dim aXY(1) As Long
    
    aPath = App.Path & "\zahyo\"
    aZa = "kari" & aSelUma & pArg & ".dat"
    
    fn = 3
    Open aPath & aZa For Binary Access Write As #fn
    
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "RACE.SyussoTosu,race.HassoTime, UMA_RACE.Year, UMA_RACE.MonthDay, UMA_RACE.JyoCD,UMA_RACE.racenum,  "
    gstrSql = gstrSql + "UMA_RACE.KakuteiJyuni,  UMA_RACE.Ninki, UMA_RACE.TanOdds5, UMA_RACE.TanNinki5, UMA_RACE.TanNinki1, UMA_RACE.fukuninki1, UMA_RACE.FOdds, UMA_RACE.Umaban, UMA_RACE.DMJyuni, UMA_RACE.CmpiNinki,UMA_RACE.hensaCmpi,UMA_RACE.CmpiValue "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "RACE INNER JOIN "
    gstrSql = gstrSql + "UMA_RACE ON "
    gstrSql = gstrSql + "(RACE.RaceNum = UMA_RACE.RaceNum) AND "
    gstrSql = gstrSql + "(RACE.JyoCD = UMA_RACE.JyoCD) AND "
    gstrSql = gstrSql + "(RACE.MonthDay = UMA_RACE.MonthDay) AND "
    gstrSql = gstrSql + "(RACE.Year = UMA_RACE.Year) "
    gstrSql = gstrSql + "WHERE "
    gstrSql = gstrSql + "(UMA_RACE.year<='2016') and "
    gstrSql = gstrSql + "(UMA_RACE.JyoCD<='10')  "
    gstrSql = gstrSql + "ORDER BY "
    gstrSql = gstrSql + "UMA_RACE.Year, UMA_RACE.MonthDay, UMA_RACE.JyoCD, UMA_RACE.RaceNum, UMA_RACE.CmpiNinki"
    ' テーブル名を指定してレコードセットを作成する
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)

    Do
        If Rs.EOF = False Then
            aYear = Rs("year")
            aMonthday = Rs("MonthDay")
            aJyoCd = Rs("JyoCD")
            aRaceNum = Rs("RaceNum")
            gYear = aYear
            gMonthDay = aMonthday
            gJyoCD = aJyoCd
            gRaceNum = aRaceNum
            
            Select Case aJyoCd
            Case "01"
                jyoName = "札幌"
            Case "02"
                jyoName = "函館"
            Case "03"
                jyoName = "福島"
            Case "04"
                jyoName = "新潟"
            Case "05"
                jyoName = "東京"
            Case "06"
                jyoName = "中山"
            Case "07"
                jyoName = "中京"
            Case "08"
                jyoName = "京都"
            Case "09"
                jyoName = "阪神"
            Case "10"
                jyoName = "小倉"
            End Select
            
            cmpininki = ""
            If IsNull(Rs("CmpiNinki")) Then
            Else
                cmpininki = Rs("CmpiNinki")
                CmpiValue = Rs("CmpiValue")
                
            End If
            
            If cmpininki = "01" Then
                If aChoiceFlg = True Then
                    '清算
                    total = total + 1
                    win = win + aKariWin
                    mny = mny + aKariMny
                    
                    '描画 DateDiff("d", "2012/1/2", "2012/4/2")
                    aDays = DateDiff("d", aStartDate, aDrawDate)
'                    Debug.Print aDays
                    'picpic.PSet (aDays, mny), vbBlack
                    Call sDot(aDays * 4, 1050 - aKariMny)
                    aXY(0) = aDays * 4
                    aXY(1) = 1050 - aKariMny
                    
                    '座標書き込み
                    Put #fn, , aXY
                    
                End If
                
                aKariWin = 0
                aKariMny = 0
                aChoiceFlg = True
            End If
            
            aDrawDate = aYear & "/" & Left$(aMonthday, 2) & "/" & Right$(aMonthday, 2)
            
            If cmpininki <> "" And aChoiceFlg = True Then
                umauma(0) = Rs("umaban")
                
                'パターン選択合致チェック
                aNinki = CInt(cmpininki)
'                If opt(aNinki - 1).BackColor = vbBlue Then
                If mChoise(aNinki - 1) = 1 Then
                    For ii = 0 To 99
                        If mPos(ii) <> 0 Then
                            If mPos(ii) = aNinki Then
                                If CInt(CmpiValue) <> (90 - ii) Then
                                    aChoiceFlg = False
                                End If
                                
                                Exit For
                            End If
                        End If
                    Next ii
                End If
                
                If aChoiceFlg = True Then
                    If cmpininki = aSelUma Then
                        '払い戻し
                        ret = getHarai(resUma, resPay)
                        If ret <> 0 Then
Debug.Print aYear & ", " & aMonthday & ", " & aJyoCd & ", " & aRaceNum
                            aChoiceFlg = False
                        Else
                            For idx2 = 3 To 7
                                If resUma(idx2) = umauma(0) Then
                                    '次のレースのデータの時と最後に清算
                                    aKariWin = 1
                                    aKariMny = CLng(resPay(idx2))
'                                    Debug.Print aYear & ", " & aMonthday & ", " & aJyoCD & ", " & aRaceNum & ", " & aKariMny
                                End If
                            Next idx2
                        End If

'Debug.Print aYear & ", " & aMonthday & ", " & aJyoCD & ", " & aRaceNum
                    End If
                End If
            End If
            
        Else
            Exit Do
        End If
        
        Rs.MoveNext
    
    Loop
    
    Rs.Close
    
    If aChoiceFlg = True Then
        '清算
        total = total + 1
        win = win + aKariWin
        mny = mny + aKariMny
    End If
    
    Close #fn
    
    Dim aRcnt As String
    Dim aRtRate As String
    Dim aHtRate As String
    Dim aCmpiPtn As String
    Dim aLookNinki As String
    
    If total > 0 Then
        'データベース登録
        aRcnt = Format$(total, "000000")
        aRtRate = Format$((mny / total), "000000.000")
        aHtRate = Format$((win / total) * 100, "000.000")
        aCmpiPtn = pArg
        aLookNinki = aSelUma
        
        gstrSql = ""
        gstrSql = gstrSql + "SELECT "
        gstrSql = gstrSql + "* "
        gstrSql = gstrSql + "FROM "
        gstrSql = gstrSql + "analCmpi "
        gstrSql = gstrSql + "where "
        gstrSql = gstrSql + "CmpiPtn ='" & aCmpiPtn & "' and "
        gstrSql = gstrSql + "LookNinki ='" & aLookNinki & "' "
        
        Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
        
        If Rs.EOF = True Then
            'データベースに追加
            gstrSql = ""
            gstrSql = gstrSql + "insert into analCmpi (Rcnt, RtRate, HtRate, CmpiPtn, LookNinki"
            gstrSql = gstrSql + ") values ("
            
            gstrSql = gstrSql + "'" & aRcnt & "', "
            gstrSql = gstrSql + "'" & aRtRate & "', "
            gstrSql = gstrSql + "'" & aHtRate & "', "
            gstrSql = gstrSql + "'" & aCmpiPtn & "', "
            gstrSql = gstrSql + "'" & aLookNinki & "')"
            
            db.Execute gstrSql, dbFailOnError
        End If
        
        Rs.Close
        
        
        '座標ファイルを仮をリネーム
        
    Else
        '仮を削除
        
    End If
    
    
    
    
    aWk = "total=> " & total
    Debug.Print aWk
    aMsg = aMsg & vbCr & vbLf & aWk
    aWebMsg = aWebMsg & "</br></br>" & aWk
    
    aWk = "win=> " & win
    Debug.Print aWk
    aMsg = aMsg & vbCr & vbLf & aWk
    aWebMsg = aWebMsg & "</br></br>" & aWk
    
    aWk = "mny=> " & mny
    Debug.Print aWk
    aMsg = aMsg & vbCr & vbLf & aWk
    aWebMsg = aWebMsg & "</br></br>" & aWk
    
    If total > 0 Then
        aWk = "hit rate=> " & Format$((win / total) * 100, "0.000")
        Debug.Print aWk
        aMsg = aMsg & vbCr & vbLf & aWk
        aWebMsg = aWebMsg & "</br></br>" & aWk
        
        aWk = "ret rate=> " & Format$((mny / total), "0.000")
        Debug.Print aWk
        aMsg = aMsg & vbCr & vbLf & aWk
        aWebMsg = aWebMsg & "</br></br>" & aWk & "</br></br>"
    End If
    Text1.Text = aMsg & vbCr & vbLf & Text1.Text

End Function

Private Function getRate_Empty(pArg As String) As Integer
    Dim aSelUma As String
    
    aSelUma = Format$(mSel, "00")
    
    
    Dim aRcnt As String
    Dim aRtRate As String
    Dim aHtRate As String
    Dim aCmpiPtn As String
    Dim aLookNinki As String
    
    'データベース登録
    aRcnt = "" 'Format$(total, "000000")
    aRtRate = "" 'Format$((mny / total), "000000.000")
    aHtRate = "" 'Format$((win / total) * 100, "000.000")
    aCmpiPtn = pArg
    aLookNinki = aSelUma
    
    '_xx-xx_xx-xx このパターンで、選択馬の条件
'    If Len(aCmpiPtn) = 12 And Mid$(aCmpiPtn, 2, 2) <> aSelUma Then
    If Len(aCmpiPtn) = 18 And Mid$(aCmpiPtn, 2, 2) <> aSelUma Then
    
        gstrSql = ""
        gstrSql = gstrSql + "SELECT "
        gstrSql = gstrSql + "* "
        gstrSql = gstrSql + "FROM "
        gstrSql = gstrSql + "analCmpi "
        gstrSql = gstrSql + "where "
        gstrSql = gstrSql + "CmpiPtn ='" & aCmpiPtn & "' and "
        gstrSql = gstrSql + "LookNinki ='" & aLookNinki & "' "
        
        Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
        
        If Rs.EOF = True Then
            'データベースに追加
            gstrSql = ""
            gstrSql = gstrSql + "insert into analCmpi (Rcnt, RtRate, HtRate, CmpiPtn, LookNinki"
            gstrSql = gstrSql + ") values ("
            
            gstrSql = gstrSql + "'" & aRcnt & "', "
            gstrSql = gstrSql + "'" & aRtRate & "', "
            gstrSql = gstrSql + "'" & aHtRate & "', "
            gstrSql = gstrSql + "'" & aCmpiPtn & "', "
            gstrSql = gstrSql + "'" & aLookNinki & "')"
            
            db.Execute gstrSql, dbFailOnError
        End If
        
        Rs.Close
    End If

End Function


Private Function setParam(pArg As String) As Integer
    Dim aDat() As String
    Dim aSel() As String
    Dim aVal() As String
    Dim aWk() As String
    Dim ret As Integer
    
    '_01-70_02-67
    aDat = Split(pArg, "_")
    ReDim aSel(UBound(aDat) - 1)
    ReDim aVal(UBound(aDat) - 1)
    
    For ii = 1 To UBound(aDat)
        aWk = Split(aDat(ii), "-")
        aSel(ii - 1) = aWk(0)
        aVal(ii - 1) = aWk(1)
    Next ii
    
    For ii = 0 To UBound(aSel)
        mSel = aSel(ii)
        For jj = 0 To UBound(mPos)
            mPos(jj) = 0
        Next jj
        For jj = 0 To UBound(mChoise)
            mChoise(jj) = 0
        Next jj
        
        For jj = 0 To UBound(aSel)
            mPos(90 - aVal(jj)) = aSel(jj)
            mChoise(aSel(jj) - 1) = 1
        Next jj
        
        '分析
        ret = getRate(pArg)
    Next ii
    
End Function

Private Function getAnal(pArg As String, pSel As String) As String
    Dim aDrawDate As String
    Dim aStartDate As Date
    Dim aDays As Long
    Dim win As Long
    Dim total As Long
    Dim mny As Long
    Dim betTotal As Long
    Dim betWin As Long
    Dim resUma(41) As String
    Dim resPay(41) As String
    Dim umauma(5) As String
    Dim aSelUma As String
    Dim aChoiceFlg As Integer
    Dim aChoiceMax As Integer
    Dim aNinki As Integer
    Dim aKariWin As Long
    Dim aKariMny As Long
    Dim aKariBetTotal As Long
    Dim aMsg As String
    Dim aWebMsg As String
    Dim aWk As String
    
    aStartDate = "2007/1/1"
    
    For ii = 0 To 17
        If mChoise(ii) <> 0 Then                   'arg
            aChoiceMax = aChoiceMax + 1
        End If
    Next ii
    
    total = 0
    win = 0
    mny = 0
    
    aSelUma = Format$(pSel, "00")
    
'    '座標ファイル仮を開く
'    Dim aPath As String
'    Dim aZa As String
'    Dim aXY(1) As Long
'
'    aPath = App.Path & "\zahyo\"
'    aZa = "kari" & aSelUma & pArg & ".dat"
'
'    fn = 3
'    Open aPath & aZa For Binary Access Write As #fn
    
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "RACE.SyussoTosu,race.HassoTime, UMA_RACE.Year, UMA_RACE.MonthDay, UMA_RACE.JyoCD,UMA_RACE.racenum,  "
    gstrSql = gstrSql + "UMA_RACE.KakuteiJyuni,  UMA_RACE.Ninki, UMA_RACE.TanOdds5, UMA_RACE.TanNinki5, UMA_RACE.TanNinki1, UMA_RACE.fukuninki1, UMA_RACE.FOdds, UMA_RACE.Umaban, UMA_RACE.DMJyuni, UMA_RACE.CmpiNinki,UMA_RACE.hensaCmpi,UMA_RACE.CmpiValue "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "RACE INNER JOIN "
    gstrSql = gstrSql + "UMA_RACE ON "
    gstrSql = gstrSql + "(RACE.RaceNum = UMA_RACE.RaceNum) AND "
    gstrSql = gstrSql + "(RACE.JyoCD = UMA_RACE.JyoCD) AND "
    gstrSql = gstrSql + "(RACE.MonthDay = UMA_RACE.MonthDay) AND "
    gstrSql = gstrSql + "(RACE.Year = UMA_RACE.Year) "
    gstrSql = gstrSql + "WHERE "
    gstrSql = gstrSql + "(UMA_RACE.year<='2016') and "
    gstrSql = gstrSql + "(UMA_RACE.JyoCD<='10')  "
    gstrSql = gstrSql + "ORDER BY "
    gstrSql = gstrSql + "UMA_RACE.Year, UMA_RACE.MonthDay, UMA_RACE.JyoCD, UMA_RACE.RaceNum, UMA_RACE.CmpiNinki"
    ' テーブル名を指定してレコードセットを作成する
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)

    Do
        If Rs.EOF = False Then
            
            cmpininki = ""
            If IsNull(Rs("CmpiNinki")) Then
            Else
                cmpininki = Rs("CmpiNinki")
                CmpiValue = Rs("CmpiValue")
                
            End If
            
            If cmpininki = "01" Then
                If aChoiceFlg = aChoiceMax Then
                    '清算
'Debug.Print aYear & ", " & aMonthday & ", " & aJyoCD & ", " & aRaceNum & ", " & aKariMny
                    total = total + 1
                    win = win + aKariWin
                    mny = mny + aKariMny
                    
'                    '描画 DateDiff("d", "2012/1/2", "2012/4/2")
'                    aDays = DateDiff("d", aStartDate, aDrawDate)
''                    Debug.Print aDays
'                    'picpic.PSet (aDays, mny), vbBlack
'                    Call sDot(aDays * 4, 1050 - aKariMny)
'                    aXY(0) = aDays * 4
'                    aXY(1) = 1050 - aKariMny
'
'                    '座標書き込み
'                    Put #fn, , aXY
                    
                End If
                
                umauma(0) = ""
                aKariWin = 0
                aKariMny = 0
                aChoiceFlg = 0
            End If
            
            aYear = Rs("year")
            aMonthday = Rs("MonthDay")
            aJyoCd = Rs("JyoCD")
            aRaceNum = Rs("RaceNum")
            gYear = aYear
            gMonthDay = aMonthday
            gJyoCD = aJyoCd
            gRaceNum = aRaceNum
            
'            Select Case aJyoCD
'            Case "01"
'                jyoName = "札幌"
'            Case "02"
'                jyoName = "函館"
'            Case "03"
'                jyoName = "福島"
'            Case "04"
'                jyoName = "新潟"
'            Case "05"
'                jyoName = "東京"
'            Case "06"
'                jyoName = "中山"
'            Case "07"
'                jyoName = "中京"
'            Case "08"
'                jyoName = "京都"
'            Case "09"
'                jyoName = "阪神"
'            Case "10"
'                jyoName = "小倉"
'            End Select
            
'            aDrawDate = aYear & "/" & Left$(aMonthday, 2) & "/" & Right$(aMonthday, 2)
            
            If aChoiceFlg < aChoiceMax Then
                If cmpininki <> "" Then
                    If aSelUma = cmpininki Then     '0702
                        umauma(0) = Rs("umaban")
                    End If
                    
                    'パターン選択合致チェック
                    aNinki = CInt(cmpininki)
    '                If opt(aNinki - 1).BackColor = vbBlue Then
                    If mChoise(aNinki - 1) = 1 Then
                        For ii = 0 To 99
                            If mPos(ii) <> 0 Then
                                If mPos(ii) = aNinki Then
                                    If CInt(CmpiValue) <> (90 - ii) Then
                                    Else
                                        aChoiceFlg = aChoiceFlg + 1
                                        If aChoiceFlg = aChoiceMax Then
                                            aChoiceFlg = aChoiceFlg
                                        End If
                                    End If
                                    
                                    Exit For
                                End If
                            End If
                        Next ii
                    End If
                    
                    If aChoiceFlg > 0 Then
                        If cmpininki = aSelUma Then
                            '払い戻し
                            ret = getHarai(resUma, resPay)
                            If ret <> 0 Then
    'Debug.Print aYear & ", " & aMonthday & ", " & aJyoCD & ", " & aRaceNum
    '                            aChoiceFlg = False
                            Else
                                For idx2 = 3 To 7
                                    If resUma(idx2) <> "" Then
                                        If resUma(idx2) = umauma(0) Then
                                            '次のレースのデータの時と最後に清算
                                            aKariWin = 1
                                            aKariMny = CLng(resPay(idx2))
                                            Exit For
    '                                        Debug.Print aYear & ", " & aMonthday & ", " & aJyoCD & ", " & aRaceNum & ", " & aKariMny
                                        End If
                                    End If
                                Next idx2
                            End If
    
    'Debug.Print aYear & ", " & aMonthday & ", " & aJyoCD & ", " & aRaceNum
                        End If
                    End If
                End If
            End If
            
        Else
            Exit Do
        End If
        
        Rs.MoveNext
    
    Loop
    
    Rs.Close
    
    If aChoiceFlg = aChoiceMax Then
        '清算
'Debug.Print aYear & ", " & aMonthday & ", " & aJyoCD & ", " & aRaceNum & ", " & aKariMny
        total = total + 1
        win = win + aKariWin
        mny = mny + aKariMny
    End If
    
'    Close #fn
    
    Dim aRcnt As String
    Dim aRtRate As String
    Dim aHtRate As String
    Dim aCmpiPtn As String
    Dim aLookNinki As String
    
    If total > 0 Then
        'データベース登録
        aRcnt = Format$(total, "000000")
        aRtRate = Format$((mny / total), "000000.000")
        aHtRate = Format$((win / total) * 100, "000.000")
        aCmpiPtn = pArg
        aLookNinki = aSelUma
        
        getAnal = aRcnt & "," & aRtRate & "," & aHtRate
        
        '座標ファイルを仮をリネーム
        
    Else
        '仮を削除
        getAnal = ""
    End If

End Function


Private Function setParam_Empty(pArg As String) As Integer
    Dim aDat() As String
    Dim aSel() As String
    Dim aVal() As String
    Dim aWk() As String
    Dim ret As Integer
    
    '_01-70_02-67
    aDat = Split(pArg, "_")
    ReDim aSel(UBound(aDat) - 1)
    ReDim aVal(UBound(aDat) - 1)
    
    For ii = 1 To UBound(aDat)
        aWk = Split(aDat(ii), "-")
        aSel(ii - 1) = aWk(0)
        aVal(ii - 1) = aWk(1)
    Next ii
    
    For ii = 0 To UBound(aSel)
        mSel = aSel(ii)
        For jj = 0 To UBound(mPos)
            mPos(jj) = 0
        Next jj
        For jj = 0 To UBound(mChoise)
            mChoise(jj) = 0
        Next jj
        
        For jj = 0 To UBound(aSel)
            mPos(90 - aVal(jj)) = aSel(jj)
            mChoise(aSel(jj) - 1) = 1
        Next jj
        
        '分析
        ret = getRate_Empty(pArg)
    Next ii
    
End Function


Private Sub cmd_Click(Index As Integer)
    Select Case Index
    Case 0
        'left
        Call fMove(-1)
    Case 1
        'right
        Call fMove(1)
    Case 2
        'on off
        If opt(mSel - 1).BackColor <> vbBlue Then
            opt(mSel - 1).BackColor = vbBlue
            mChoise(mSel - 1) = 1
            Command1.Enabled = True
        Else
            opt(mSel - 1).BackColor = vbRed
            mChoise(mSel - 1) = 0
'            Command1.Enabled = False
        End If
    End Select
    
    If Index < 2 Then
        '再描画
        Call sDraw
    End If
End Sub

Private Sub fMove(aLR As Integer)
    Dim ii As Integer
    Dim aPos As Integer
    Dim aNow As Integer
    
    '現在位置検索
    For ii = 0 To 99
        If mPos(ii) = mSel Then
            aPos = ii
            Exit For
        End If
    Next ii
    
    '空きチェック
    aNow = aPos
    Do
        aNow = aNow + aLR
        If aNow = -1 Or aNow = 60 Then
            Exit Do
        End If
        
        If mPos(aNow) = 0 Then
            Do
                mPos(aNow) = mPos(aNow - aLR)
                aNow = aNow - aLR
                If aNow = aPos Then
                    mPos(aNow) = 0
                    Exit Do
                End If
            Loop
            
            Exit Do
        End If
    Loop
    
End Sub

Private Sub sDot(X As Long, Y As Long)
    picpic.PSet (X, Y), vbYellow
'    picpic.PSet (x, y + 1), vbYellow
'    picpic.PSet (x + 1, y), vbYellow
'    picpic.PSet (x + 1, y + 1), vbYellow
End Sub
Private Sub sLine(X As Long, Y As Long)
    picpic.Line (X, Y)-(X, Y - 100), vbWhite
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    
    Dim aDrawDate As String
    Dim aStartDate As Date
    Dim aDays As Long
    Dim win As Long
    Dim total As Long
    Dim mny As Long
    Dim betTotal As Long
    Dim betWin As Long
    Dim resUma(41) As String
    Dim resPay(41) As String
    Dim umauma(5) As String
    Dim aSelUma As String
    Dim aChoiceFlg As Integer
    Dim aChoiceMax As Integer
    Dim aNinki As Integer
    Dim aKariWin As Long
    Dim aKariMny As Long
    Dim aKariBetTotal As Long
    Dim aMsg As String
    Dim aWebMsg As String
    Dim aWk As String
    Dim aRaceNum As String
    
    aStartDate = "2007/1/1"
'    aStartDate = "2017/1/1"
    
    For ii = 0 To 17
        If mChoise(ii) <> 0 Then                   'arg
            aChoiceMax = aChoiceMax + 1
        End If
    Next ii
    
    fn = FreeFile
    Open App.Path & "\_today.txt" For Output As #fn
    
    total = 0
    win = 0
    mny = 0
    aSelUma = Format$(mSel, "00")       'arg
    
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "RACE.SyussoTosu,race.HassoTime, UMA_RACE.Year, UMA_RACE.MonthDay, UMA_RACE.JyoCD,UMA_RACE.racenum,  "
    gstrSql = gstrSql + "UMA_RACE.KakuteiJyuni,  UMA_RACE.Ninki, UMA_RACE.TanOdds5, UMA_RACE.TanNinki5, UMA_RACE.TanNinki1, UMA_RACE.fukuninki1, UMA_RACE.FOdds, UMA_RACE.Umaban, UMA_RACE.DMJyuni, UMA_RACE.CmpiNinki,UMA_RACE.hensaCmpi,UMA_RACE.CmpiValue "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "RACE INNER JOIN "
    gstrSql = gstrSql + "UMA_RACE ON "
    gstrSql = gstrSql + "(RACE.RaceNum = UMA_RACE.RaceNum) AND "
    gstrSql = gstrSql + "(RACE.JyoCD = UMA_RACE.JyoCD) AND "
    gstrSql = gstrSql + "(RACE.MonthDay = UMA_RACE.MonthDay) AND "
    gstrSql = gstrSql + "(RACE.Year = UMA_RACE.Year) "
    gstrSql = gstrSql + "WHERE "
'        gstrSql = gstrSql + "(RACE.SyussoTosu='16') AND "

'        gstrSql = gstrSql + "(UMA_RACE.Year='2007') AND "
'    gstrSql = gstrSql + "UMA_RACE.MonthDay <= '0331' and "

'        gstrSql = gstrSql + "(UMA_RACE.Year='2014') AND "

    gstrSql = gstrSql + "(UMA_RACE.year<='2016') and "
    gstrSql = gstrSql + "(UMA_RACE.JyoCD<='10')  "
'    gstrSql = gstrSql + "(UMA_RACE.JyoCD>'40') AND "

'    gstrSql = gstrSql + "(UMA_RACE.Year>='" & Format$(Now, "yyyy") & "') AND "
'    gstrSql = gstrSql + "UMA_RACE.MonthDay='" & Format$(Now, "mmdd") & "' and "

'''    gstrSql = gstrSql + "RACE.DataKubun='7' and "

'''    gstrSql = gstrSql + " and "
'''    gstrSql = gstrSql + "TrackCD >='10' and "       '障害除く
'''    gstrSql = gstrSql + "TrackCD <='29' and "
'''    gstrSql = gstrSql + "JyokenCD5 <> '701'  "

    gstrSql = gstrSql + "ORDER BY "
    gstrSql = gstrSql + "UMA_RACE.Year, UMA_RACE.MonthDay, UMA_RACE.JyoCD, UMA_RACE.RaceNum, UMA_RACE.CmpiNinki"
'        gstrSql = gstrSql + "UMA_RACE.Year, UMA_RACE.MonthDay, race.HassoTime"
'        gstrSql = gstrSql + "race.HassoTime"
    ' テーブル名を指定してレコードセットを作成する
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)

    Do
        If Rs.EOF = False Then
            
'            Select Case aJyoCD
'            Case "01"
'                jyoName = "札幌"
'            Case "02"
'                jyoName = "函館"
'            Case "03"
'                jyoName = "福島"
'            Case "04"
'                jyoName = "新潟"
'            Case "05"
'                jyoName = "東京"
'            Case "06"
'                jyoName = "中山"
'            Case "07"
'                jyoName = "中京"
'            Case "08"
'                jyoName = "京都"
'            Case "09"
'                jyoName = "阪神"
'            Case "10"
'                jyoName = "小倉"
'            End Select
            
            cmpininki = ""
            If IsNull(Rs("CmpiNinki")) Then
            Else
                cmpininki = Rs("CmpiNinki")
                CmpiValue = Rs("CmpiValue")
                
            End If
            
            If cmpininki = "01" Then
                If aChoiceFlg = aChoiceMax Then
'Debug.Print aYear & ", " & aMonthday & ", " & aJyoCD & ", " & aRaceNum & ", " & aKariMny
                    '清算
                    Print #fn, aKariMny
                    total = total + 1
                    win = win + aKariWin
                    mny = mny + aKariMny
Debug.Print aKariWin & "," & aKariMny
                    
                    '描画 DateDiff("d", "2012/1/2", "2012/4/2")
                    aDays = DateDiff("d", aStartDate, aDrawDate)
                    
                    Debug.Print aDays
                   'picpic.PSet (aDays, mny), vbBlack
                   If aKariMny = 0 Then
                    Call sDot(aDays * 2, 7000 - 500 + 100)
                   Else
                    Call sDot(aDays * 2, 7000 - 500 - aKariMny)
                   End If
                End If
                
                umauma(0) = ""
                aKariWin = 0
                aKariMny = 0
                aChoiceFlg = 0
            End If
            
            aYear = Rs("year")
            aMonthday = Rs("MonthDay")
            aJyoCd = Rs("JyoCD")
            aRaceNum = Rs("RaceNum")
            gYear = aYear
            gMonthDay = aMonthday
            gJyoCD = aJyoCd
            gRaceNum = aRaceNum
            
            aDrawDate = aYear & "/" & Left$(aMonthday, 2) & "/" & Right$(aMonthday, 2)
            
            If aChoiceFlg < aChoiceMax Then
                If cmpininki <> "" Then
                    If aSelUma = cmpininki Then     '0702
                        umauma(0) = Rs("umaban")
                    End If
                    
                    'パターン選択合致チェック
                    aNinki = CInt(cmpininki)
    '                If opt(aNinki - 1).BackColor = vbBlue Then
                    If mChoise(aNinki - 1) = 1 Then                 'arg
                        For ii = 0 To 99
                            If mPos(ii) <> 0 Then                   'arg
                                If mPos(ii) = aNinki Then
                                    If CInt(CmpiValue) <> (90 - ii) Then
                                        aNinki = aNinki
                                    Else
                                        aChoiceFlg = aChoiceFlg + 1
                                    End If
                                    
                                    Exit For
                                End If
                            End If
                        Next ii
                    End If
                    
                    If aChoiceFlg > 0 Then
                        If cmpininki = aSelUma Then
                            '払い戻し
                            ret = getHarai(resUma, resPay)
                            If ret <> 0 Then
    'Debug.Print aYear & ", " & aMonthday & ", " & aJyoCD & ", " & aRaceNum
    '                            aChoiceFlg = False
                            Else
                                For idx2 = 3 To 7
                                    If resUma(idx2) = umauma(0) Then
                                        '次のレースのデータの時と最後に清算
                                        aKariWin = 1
                                        aKariMny = CLng(resPay(idx2))
                                        Exit For
    '                                    Debug.Print aYear & ", " & aMonthday & ", " & aJyoCD & ", " & aRaceNum & ", " & aKariMny
                                    End If
                                Next idx2
                            End If
    
    'Debug.Print aYear & ", " & aMonthday & ", " & aJyoCD & ", " & aRaceNum
                        End If
                    End If
                End If
            End If
            
        Else
            Exit Do
        End If
        
        Rs.MoveNext
    
    Loop
    
    Rs.Close
    
    If aChoiceFlg = aChoiceMax Then
        '清算
        Print #fn, aKariMny
        total = total + 1
        win = win + aKariWin
        mny = mny + aKariMny
Debug.Print aKariWin & "," & aKariMny
    End If
    
    aWk = "total=> " & total
    Debug.Print aWk
    aMsg = aMsg & vbCr & vbLf & aWk
    aWebMsg = aWebMsg & "</br></br>" & aWk
    
    aWk = "win=> " & win
    Debug.Print aWk
    aMsg = aMsg & vbCr & vbLf & aWk
    aWebMsg = aWebMsg & "</br></br>" & aWk
    
    aWk = "mny=> " & mny
    Debug.Print aWk
    aMsg = aMsg & vbCr & vbLf & aWk
    aWebMsg = aWebMsg & "</br></br>" & aWk
    
    If total > 0 Then
        aWk = "hit rate=> " & Format$((win / total) * 100, "0.000")
        Debug.Print aWk
        aMsg = aMsg & vbCr & vbLf & aWk
        aWebMsg = aWebMsg & "</br></br>" & aWk
        
        aWk = "ret rate=> " & Format$((mny / total), "0.000")
        Debug.Print aWk
        aMsg = aMsg & vbCr & vbLf & aWk
        aWebMsg = aWebMsg & "</br></br>" & aWk & "</br></br>"
    End If
    Text1.Text = aMsg & vbCr & vbLf & Text1.Text
    
    Dim mail As String
    Dim aTitle As String
    Dim aBody As String
    
    aTitle = GC_APLI_NAME & GC_THANKS
    aBody = aWebMsg
    
'    mail = sendMail(aTitle, aBody, GC_BLOG_MAIL)
    
    Close #fn
    
    MsgBox "finish!"
    
    Command1.Enabled = True
End Sub

Private Function getHarai(pUma() As String, pPay() As String) As Integer
On Error GoTo err
    getHarai = 1
    
    '払戻チェック
    gstrSql = ""
    gstrSql = gstrSql + "select * "
    gstrSql = gstrSql + "from HARAI where "
    gstrSql = gstrSql + "Year = '" & gYear & "' and "
    gstrSql = gstrSql + "MonthDay = '" & gMonthDay & "' and "
    gstrSql = gstrSql + "JyoCD = '" & gJyoCD & "' and "
    gstrSql = gstrSql + "RaceNum = '" & gRaceNum & "'"
    ' テーブル名を指定してレコードセットを作成する
    Set RsWk = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
    If RsWk.EOF = True Then
        RsWk.Close
        
        For ii = 3 To 7
            pUma(ii) = ""
            pPay(ii) = ""
        Next ii
        
        getHarai = 1
        Exit Function
    End If
    
    
'    pUma(0) = RsWk("PayTansyoUmaban1")
'    pPay(0) = RsWk("PayTansyoPay1")
'    pUma(1) = RsWk("PayTansyoUmaban2")
'    pPay(1) = RsWk("PayTansyoPay2")
'    pUma(2) = RsWk("PayTansyoUmaban3")
'    pPay(2) = RsWk("PayTansyoPay3")
    If IsNull(RsWk("PayFukusyoUmaban1")) = False Then
        pUma(3) = RsWk("PayFukusyoUmaban1")
        pPay(3) = RsWk("PayFukusyoPay1")
    Else
        pUma(3) = ""
        pPay(3) = ""
    End If
    If IsNull(RsWk("PayFukusyoUmaban2")) = False Then
        pUma(4) = RsWk("PayFukusyoUmaban2")
        pPay(4) = RsWk("PayFukusyoPay2")
    Else
        pUma(4) = ""
        pPay(4) = ""
    End If
    If IsNull(RsWk("PayFukusyoUmaban3")) = False Then
        pUma(5) = RsWk("PayFukusyoUmaban3")
        pPay(5) = RsWk("PayFukusyoPay3")
    Else
        pUma(5) = ""
        pPay(5) = ""
    End If
    If IsNull(RsWk("PayFukusyoUmaban4")) = False Then
        pUma(6) = RsWk("PayFukusyoUmaban4")
        pPay(6) = RsWk("PayFukusyoPay4")
    Else
        pUma(6) = ""
        pPay(6) = ""
    End If
    If IsNull(RsWk("PayFukusyoUmaban5")) = False Then
        pUma(7) = RsWk("PayFukusyoUmaban5")
        pPay(7) = RsWk("PayFukusyoPay5")
    Else
        pUma(7) = ""
        pPay(7) = ""
    End If
'    pUma(8) = RsWk("PayUmarenKumi1")
'    pPay(8) = RsWk("PayUmarenPay1")
'    pUma(9) = RsWk("PayUmarenKumi2")
'    pPay(9) = RsWk("PayUmarenPay2")
'    pUma(10) = RsWk("PayUmarenKumi3")
'    pPay(10) = RsWk("PayUmarenPay3")
'    pUma(11) = RsWk("PayWideKumi1")
'    pPay(11) = RsWk("PayWidePay1")
'    pUma(12) = RsWk("PayWideKumi2")
'    pPay(12) = RsWk("PayWidePay2")
'    pUma(13) = RsWk("PayWideKumi3")
'    pPay(13) = RsWk("PayWidePay3")
'    pUma(14) = RsWk("PayWideKumi4")
'    pPay(14) = RsWk("PayWidePay4")
'    pUma(15) = RsWk("PayWideKumi5")
'    pPay(15) = RsWk("PayWidePay5")
'    pUma(16) = RsWk("PayWideKumi6")
'    pPay(16) = RsWk("PayWidePay6")
'    pUma(17) = RsWk("PayWideKumi7")
'    pPay(17) = RsWk("PayWidePay7")
'    pUma(18) = RsWk("PaySanrenpukuKumi1")
'    pPay(18) = RsWk("PaySanrenpukuPay1")
'    pUma(19) = RsWk("PaySanrenpukuKumi2")
'    pPay(19) = RsWk("PaySanrenpukuPay2")
'    pUma(20) = RsWk("PaySanrenpukuKumi3")
'    pPay(20) = RsWk("PaySanrenpukuPay3")
'
'    pUma(21) = RsWk("PayWakurenKumi1")
'    pPay(21) = RsWk("PayWakurenPay1")
'    pUma(22) = RsWk("PayWakurenKumi2")
'    pPay(22) = RsWk("PayWakurenPay2")
'    pUma(23) = RsWk("PayWakurenKumi3")
'    pPay(23) = RsWk("PayWakurenPay3")
'
'    pUma(24) = RsWk("PaySanrentanKumi1")
'    pPay(24) = RsWk("PaySanrentanPay1")
'    pUma(25) = RsWk("PaySanrentanKumi2")
'    pPay(25) = RsWk("PaySanrentanPay2")
'    pUma(26) = RsWk("PaySanrentanKumi3")
'    pPay(26) = RsWk("PaySanrentanPay3")
'    pUma(27) = RsWk("PaySanrentanKumi4")
'    pPay(27) = RsWk("PaySanrentanPay4")
'    pUma(28) = RsWk("PaySanrentanKumi5")
'    pPay(28) = RsWk("PaySanrentanPay5")
'    pUma(29) = RsWk("PaySanrentanKumi6")
'    pPay(29) = RsWk("PaySanrentanPay6")
'
'    pUma(30) = RsWk("PayUmatanKumi1")
'    pPay(30) = RsWk("PayUmatanPay1")
'    pUma(31) = RsWk("PayUmatanKumi2")
'    pPay(31) = RsWk("PayUmatanPay2")
'    pUma(32) = RsWk("PayUmatanKumi3")
'    pPay(32) = RsWk("PayUmatanPay3")
'    pUma(33) = RsWk("PayUmatanKumi4")
'    pPay(33) = RsWk("PayUmatanPay4")
'    pUma(34) = RsWk("PayUmatanKumi5")
'    pPay(34) = RsWk("PayUmatanPay5")
'    pUma(35) = RsWk("PayUmatanKumi6")
'    pPay(35) = RsWk("PayUmatanPay6")
    
    RsWk.Close
    
    getHarai = 0
    Exit Function
err:
'    Debug.Print err.Description
    If err.Number = 3021 Then
        getHarai = -1
    End If
    Exit Function
End Function

Private Sub Command10_Click()
    Dim aRaw As String
    
    Command10.Enabled = False
    
    fn = FreeFile
    Open App.Path & "\ptn_all.txt" For Output As #fn
        
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "UMA_RACE.Year, UMA_RACE.MonthDay, UMA_RACE.JyoCD,UMA_RACE.racenum,  "
    gstrSql = gstrSql + "UMA_RACE.Ninki, UMA_RACE.Umaban, UMA_RACE.CmpiNinki, UMA_RACE.CmpiValue "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "RACE INNER JOIN "
    gstrSql = gstrSql + "UMA_RACE ON "
    gstrSql = gstrSql + "(RACE.RaceNum = UMA_RACE.RaceNum) AND "
    gstrSql = gstrSql + "(RACE.JyoCD = UMA_RACE.JyoCD) AND "
    gstrSql = gstrSql + "(RACE.MonthDay = UMA_RACE.MonthDay) AND "
    gstrSql = gstrSql + "(RACE.Year = UMA_RACE.Year) "
    gstrSql = gstrSql + "WHERE "
    gstrSql = gstrSql + "(UMA_RACE.year<='2016') and "
    gstrSql = gstrSql + "(UMA_RACE.JyoCD<='10') and "
    gstrSql = gstrSql + "TrackCD >='10' and "       '障害除く
    gstrSql = gstrSql + "TrackCD <='29' and "
    gstrSql = gstrSql + "JyokenCD5 <> '701'  "

    gstrSql = gstrSql + "ORDER BY "
    gstrSql = gstrSql + "UMA_RACE.Year, UMA_RACE.MonthDay, UMA_RACE.JyoCD, UMA_RACE.RaceNum, UMA_RACE.CmpiNinki"
    ' テーブル名を指定してレコードセットを作成する
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)

    Do
        If Rs.EOF = False Then
            aYear = Rs("year")
            aMonthday = Rs("MonthDay")
            aJyoCd = Rs("JyoCD")
            aRaceNum = Rs("RaceNum")
            gYear = aYear
            gMonthDay = aMonthday
            gJyoCD = aJyoCd
            gRaceNum = aRaceNum
            
            cmpininki = ""
            If IsNull(Rs("CmpiNinki")) Then
            Else
                cmpininki = Rs("CmpiNinki")
                CmpiValue = Rs("CmpiValue")
                
            End If
            
            If cmpininki = "01" Then
                Print #fn, aRaw
                aRaw = ""
            End If
            
            If CmpiValue <> "00" And CmpiValue <> "" Then
                aRaw = aRaw & "_" & CmpiValue
            End If
            
        Else
            Exit Do
        End If
        
        Rs.MoveNext
    
    Loop
    
    Rs.Close
        
    Close #fn
    
    Command10.Enabled = True
    
    MsgBox "end"
End Sub

Private Sub Command11_Click()
    Dim StartTime  As Long
    Dim StopTime  As Long

    StartTime = GetTickCount
    
    Dim lCnt As Long
    Dim aArr() As String
    
    Command11.Enabled = False
    
    fn = FreeFile
    Open App.Path & "\ptn_all_unique.txt" For Input As #fn
    
    lCnt = 0
    Do Until EOF(fn)
        Line Input #fn, wk
        lCnt = lCnt + 1
    Loop
    
    Close #fn
    
    ReDim aArr(lCnt - 1)
    
    fn = FreeFile
    Open App.Path & "\ptn_all_unique.txt" For Input As #fn
    
    lCnt = 0
    Do Until EOF(fn)
        Line Input #fn, wk
        aArr(lCnt) = wk
        lCnt = lCnt + 1
    Loop
    
    Close #fn
    
    Dim aVal() As String
    Dim aCmb() As String
    Dim aCmbs() As String
    Dim ii As Long
    Dim jj As Integer
    Dim kk As Long
    Dim aPut As String
    Dim aBin() As Byte
    Dim aDt(1) As Byte
    Dim aDmy As Byte
    Dim aBinC As Integer
    Dim aStart As Long
    Dim aEnd As Long
'For tt = 26 To 30
    aStart = tt * 1000 + 1 ' 1001
    aEnd = aStart + 999 ' 2000
    fn = 2
'    Open App.Path & "\ptn_cmb.txt" For Output As #fn
'    Open App.Path & "\ptn_cmb" & Format$(CStr(aStart), "00000") & "-" & Format$(CStr(aEnd), "00000") & ".dat" For Binary Access Write As #fn
    Open App.Path & "\ptn_cmb_Last.dat" For Binary Access Write As #fn
    
    '0-10 de 12948  10.6MB
    For ii = 31001 To UBound(aArr)
        aVal = Split(aArr(ii), "_")     'コンピ指数
        
        '全パターン抽出
        For jj = 1 To UBound(aVal)      '抽出数
            wk = (Combination(UBound(aVal), jj))
            aCmb = Split(wk, vbCrLf)
            
            For kk = 0 To UBound(aCmb) - 1
                aCmbs = Split(aCmb(kk), " ")
                
                
                
'                For mm = 1 To UBound(aCmbs)
'                    aDt(0) = CInt(aCmbs(mm))
'                    aDt(1) = CInt(aVal(aCmbs(mm)))
'                    Put #fn, , aDt
'                Next mm
'
'                aDmy = 0
'                Put #fn, , aDmy
                
                
                
                
                aBinC = 0
                For mm = 1 To UBound(aCmbs)
                    aBinC = aBinC + 2
                Next mm

                ReDim aBin(aBinC)
                For mm = 1 To UBound(aCmbs)
                    aBin((mm - 1) * 2) = CInt(aCmbs(mm))
                    aBin((mm - 1) * 2 + 1) = CInt(aVal(aCmbs(mm)))
                Next mm

                aBin(aBinC) = 0
                Put #fn, , aBin
                
                
                
'                aPut = ""
'                For mm = 1 To UBound(aCmbs)
'                    aPut = aPut & "_" & Format$(aCmbs(mm), "00") & "-" & aVal(aCmbs(mm))
'                Next mm
'                Print #fn, aPut
            Next kk
            
        Next jj
    Next ii
    
    Close #fn
'Next tt
    Command11.Enabled = True
    
    StopTime = GetTickCount
    Debug.Print StopTime - StartTime
    
    MsgBox "end"
    
End Sub

Private Sub Command12_Click()
    Dim fName As String
    Dim fileNum As Long
    Dim fName2 As String
    Dim fileNum2 As Long
    Dim b() As Byte
    Dim i As Long
    Dim HexStr As String
    Dim aa As Long
    Dim wk As String
    
    fName = App.Path & "\ptn_cmb.dat"
    fName2 = App.Path & "\ptn_cmb_binread.txt"
    aa = FileLen(fName)
    ReDim b(aa - 1)
    fileNum = 1
    fileNum2 = 2
    Open fName For Binary Access Read As #fileNum
    Open fName2 For Output As #fileNum2
    
'    For i = 0 To 9
'        If EOF(fileNum) Then Exit For
        Get #fileNum, , b
'        HexStr = HexStr & Right$("0" & Hex$(b), 2)
'    Next
    
    Dim cnt As Integer
    
    wk = ""
    cnt = 0
    For ii = 0 To (aa - 1)
        If b(ii) = 0 Then
            '区切り
            Print #fileNum2, "_" & Mid$(wk, 2)
            wk = ""
            cnt = 0
        Else
            If cnt Mod 2 = 0 Then
                wk = wk & "_" & Format$(b(ii), "00")
            Else
                wk = wk & "-" & Format$(b(ii), "00")
            End If
            cnt = cnt + 1
        End If
        
    Next ii
    
    
    Close #fileNum2 'Closeは後述
    Close #fileNum 'Closeは後述
    
    MsgBox "end"
End Sub


Private Sub Command13_Click()
    Dim fName As String
    Dim fileNum As Long
    Dim fName2 As String
    Dim fileNum2 As Long
    Dim b() As Byte
    Dim i As Long
    Dim HexStr As String
    Dim aa As Long
    Dim wk As String
    Dim ret As Integer
    
    fName = App.Path & "\" & txtBin.Text
    aa = FileLen(fName)
    ReDim b(aa - 1)
    fileNum = 1
    fileNum2 = 2
    Open fName For Binary Access Read As #fileNum
    
    Get #fileNum, , b
    
    Dim cnt As Integer
    
    wk = ""
    cnt = 0
    For ii = 0 To (aa - 1)
        If b(ii) = 0 Then
            '区切り DB登録
Debug.Print wk
            
            ret = setParam(wk)
            
            
'            Print #fileNum2, "_" & Mid$(wk, 2)
            wk = ""
            cnt = 0
        Else
            If cnt Mod 2 = 0 Then
                wk = wk & "_" & Format$(b(ii), "00")
            Else
                wk = wk & "-" & Format$(b(ii), "00")
            End If
            cnt = cnt + 1
        End If
        
    Next ii
    
    
    Close #fileNum 'Closeは後述
    
    MsgBox "end"
End Sub


Private Sub Command14_Click()
    On Error GoTo Err1
    
    Dim mail As String
    Dim aTitle As String
    Dim aBody As String
    Dim aMac As String
    
    Dim aRcnt As String
    Dim aRtRate As String
    Dim aHtRate As String
    
    Command14.Enabled = False
    
    If txtCnd(0).Text <> "" Then
        aRtRate = Format$(txtCnd(0).Text, "000000.000")
    End If
    
    If txtCnd(1).Text <> "" Then
        aHtRate = Format$(txtCnd(1).Text, "000.000")
    End If
    
    If txtCnd(2).Text <> "" Then
        aRcnt = Format$(txtCnd(2).Text, "000000")
    End If
    
    Dim aCnd As String
    
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "analCmpi "
    If txtCnd(0).Text <> "" Or txtCnd(1).Text <> "" Or txtCnd(2).Text <> "" Then
        gstrSql = gstrSql + "where "
        
        aCnd = ""
        If txtCnd(0).Text <> "" Then
            aCnd = aCnd + "RtRate >='" & aRtRate & "' "
        End If
        
        
        If txtCnd(1).Text <> "" Then
            If aCnd <> "" Then
                aCnd = aCnd + "and "
            End If
            aCnd = aCnd + "HtRate >='" & aHtRate & "' "
        End If
        
        If txtCnd(2).Text <> "" Then
            If aCnd <> "" Then
                aCnd = aCnd + "and "
            End If
            aCnd = aCnd + "Rcnt >='" & aRcnt & "' "
        End If
        
        gstrSql = gstrSql + aCnd
    Else
        Exit Sub
    End If
    
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
    Do
        If Rs.EOF = False Then
            aLookNinki = Rs("LookNinki")
            aCmpiPtn = Rs("CmpiPtn")
            aRcnt = Rs("Rcnt")
            aRtRate = Rs("RtRate")
            aHtRate = Rs("HtRate")
            
'            Debug.Print aLookNinki & ":" & aCmpiPtn
            txtRes.Text = txtRes.Text & "Ptn:" & aCmpiPtn & vbCrLf
            txtRes.Text = txtRes.Text & "Nin:" & aLookNinki & vbCrLf
            txtRes.Text = txtRes.Text & "Rac:" & aRcnt & vbCrLf
            txtRes.Text = txtRes.Text & "Ret:" & aRtRate & vbCrLf
            txtRes.Text = txtRes.Text & "Hit:" & aHtRate & vbCrLf
            
            DoEvents
        Else
            Exit Do
        End If
        
        Rs.MoveNext
    Loop
    
    Rs.Close

    Command14.Enabled = True
    
    Exit Sub
Err1:
    
    aTitle = "okba_error info."
    mail = sendMail(aTitle, err.Number & err.Description, GC_MAC_MAIL)
    
    Command14.Enabled = True
    
    Exit Sub
End Sub

Private Sub Command15_Click()
    Dim fName As String
    Dim fileNum As Long
    Dim fName2 As String
    Dim fileNum2 As Long
    Dim b(100) As Byte
    Dim i As Long
    Dim HexStr As String
    Dim aa As Long
    Dim wk As String
    Dim ret As Integer
    Dim aBinCnt As Integer
    Dim aDatCnt As Long
    
    Dim StartTime  As Long
    Dim StopTime  As Long
    Dim cnt As Integer
    Dim aFile(31) As String

    aFile(30) = "ptn_cmb01001-02000.dat"
    aFile(31) = "ptn_cmb00001-01000.dat"
    
    aFile(0) = "ptn_cmb02001-03000.dat"
    aFile(0) = "ptn_cmb02001-03000.dat"
    aFile(1) = "ptn_cmb03001-04000.dat"
    aFile(2) = "ptn_cmb04001-05000.dat"
    aFile(3) = "ptn_cmb05001-06000.dat"
    aFile(4) = "ptn_cmb06001-07000.dat"
    aFile(5) = "ptn_cmb07001-08000.dat"
    aFile(6) = "ptn_cmb08001-09000.dat"
    aFile(7) = "ptn_cmb09001-10000.dat"
    aFile(8) = "ptn_cmb10001-11000.dat"
    aFile(9) = "ptn_cmb11001-12000.dat"
    aFile(10) = "ptn_cmb12001-13000.dat"
    aFile(11) = "ptn_cmb13001-14000.dat"
    aFile(12) = "ptn_cmb14001-15000.dat"
    aFile(13) = "ptn_cmb15001-16000.dat"
    aFile(14) = "ptn_cmb16001-17000.dat"
    aFile(15) = "ptn_cmb17001-18000.dat"
    aFile(16) = "ptn_cmb18001-19000.dat"
    aFile(17) = "ptn_cmb19001-20000.dat"
    aFile(18) = "ptn_cmb20001-21000.dat"
    aFile(19) = "ptn_cmb21001-22000.dat"
    aFile(20) = "ptn_cmb22001-23000.dat"
    aFile(21) = "ptn_cmb23001-24000.dat"
    aFile(22) = "ptn_cmb24001-25000.dat"
    aFile(23) = "ptn_cmb25001-26000.dat"
    aFile(24) = "ptn_cmb26001-27000.dat"
    aFile(25) = "ptn_cmb27001-28000.dat"
    aFile(26) = "ptn_cmb28001-29000.dat"
    aFile(27) = "ptn_cmb29001-30000.dat"
    aFile(28) = "ptn_cmb30001-31000.dat"
    aFile(29) = "ptn_cmb30001-31000_Last.dat"
    
    
'    Dim c(160000000) As Byte
'    fName = App.Path & "\" & aFile(0)
'
'    fileNum = 1
'    fileNum2 = 2
'    Open fName For Binary Access Read As #fileNum
'    Get #fileNum, , c
'    Get #fileNum, , c
'    Get #fileNum, , c
'    Get #fileNum, , c
'    Get #fileNum, , c
'    Get #fileNum, , c
'
'    Close #fileNum 'Closeは後述



For kk = 0 To 0

    StartTime = GetTickCount
    
    
    fName = App.Path & "\" & aFile(kk)
    aa = FileLen(fName)
    
    fileNum = 1
    fileNum2 = 2
    Open fName For Binary Access Read As #fileNum
    
    wk = ""
    cnt = 0
    aDatCnt = 1
    
    Do
        aBinCnt = 0
        If (aDatCnt - 1) = aa Then
            Exit Do
        End If
        Do
            Get #fileNum, aDatCnt, b(aBinCnt)
            aDatCnt = aDatCnt + 1
            If b(aBinCnt) = 0 Then
                Exit Do
            End If
            aBinCnt = aBinCnt + 1
        Loop
        
        For ii = 0 To aBinCnt
            
            If b(ii) = 0 Then
                '区切り DB登録
'                If Len(wk) = 12 Then
                If Len(wk) = 18 Then
                    ret = setParam_Empty(wk)
                Else
                End If
                
                wk = ""
                cnt = 0
            Else
                If cnt Mod 2 = 0 Then
                    wk = wk & "_" & Format$(b(ii), "00")
                Else
                    wk = wk & "-" & Format$(b(ii), "00")
                End If
                cnt = cnt + 1
            End If
            
        Next ii
    Loop
    
    
    Close #fileNum 'Closeは後述
    
    StopTime = GetTickCount
    Me.Caption = kk & ":" & StopTime - StartTime
    Text1.Text = kk & ":" & StopTime - StartTime & vbCrLf
    Me.Refresh
    DoEvents
Next kk
    
    MsgBox "end"

End Sub

Private Sub Command16_Click()
    Dim fName As String
    Dim fileNum As Long
    Dim fName2 As String
    Dim fileNum2 As Long
    Dim b() As Byte
    Dim i As Long
    Dim HexStr As String
    Dim aa As Long
    Dim wk As String
    Dim ret As Integer
    Dim aBinCnt As Integer
    Dim aDatCnt As Long
    
    Dim StartTime  As Long
    Dim StopTime  As Long

    StartTime = GetTickCount
    
    fName = App.Path & "\" & txtBin.Text
    aa = FileLen(fName)
    ReDim b(aa - 1)
    fileNum = 1
    fileNum2 = 2
    Open fName For Binary Access Read As #fileNum
    
    Get #fileNum, , b
    
    Dim cnt As Integer
    
    wk = ""
    cnt = 0
    For ii = 0 To (aa - 1)
        If b(ii) = 0 Then
            '区切り DB登録
            ret = setParam_Empty(wk)
            
            wk = ""
            cnt = 0
        Else
            If cnt Mod 2 = 0 Then
                wk = wk & "_" & Format$(b(ii), "00")
            Else
                wk = wk & "-" & Format$(b(ii), "00")
            End If
            cnt = cnt + 1
        End If
        
    Next ii
    
    Close #fileNum 'Closeは後述
    
    StopTime = GetTickCount
    Debug.Print StopTime - StartTime
    
    MsgBox "end"

End Sub

Private Sub Command17_Click()
    Dim StartTime  As Long
    Dim StopTime  As Long
    Dim aRet As String
    Dim aLookNinki As String
    Dim aCmpiPtn As String
    Dim aRcnt As String
    Dim aUpd() As String
    Dim aCnt As Long
    Dim aTop As Long
    
    Dim aaa As Long
    Dim bbb As Long
    
    aaa = CLng(txtStart.Text)
    bbb = CLng(txtVol.Text)
    
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "analCmpi "
    gstrSql = gstrSql + "where rcnt = '' "
    gstrSql = gstrSql + "order by CmpiPtn, LookNinki "
    
    Set Rs2 = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
'    Do
'        If aaa = aTop Then
'            Exit Do
'        End If
'
'        aTop = aTop + 1
'        Rs2.MoveNext
'    Loop
    
    
    Do
        aRcnt = ""
        If Rs2.EOF = False Then
            aLookNinki = Rs2("LookNinki")
            aCmpiPtn = Rs2("CmpiPtn")
            If IsNull(Rs2("Rcnt")) = False Then
                aRcnt = Rs2("Rcnt")
            Else
                aRcnt = ""
            End If
        Else
            Exit Do
        End If
        
        If aRcnt = "" Then
'StartTime = GetTickCount
            'anal
            For jj = 0 To UBound(mPos)
                mPos(jj) = 0
            Next jj
            For jj = 0 To UBound(mChoise)
                mChoise(jj) = 0
            Next jj
            
            
            '修正必要
            mPos(90 - Mid$(aCmpiPtn, 5, 2)) = CInt(Mid$(aCmpiPtn, 2, 2))
            mPos(90 - Mid$(aCmpiPtn, 11)) = CInt(Mid$(aCmpiPtn, 8, 2))
            mChoise(CInt(Mid$(aCmpiPtn, 2, 2)) - 1) = 1
            mChoise(CInt(Mid$(aCmpiPtn, 8, 2)) - 1) = 1
            
            aRet = ""
            aRet = getAnal(aCmpiPtn, aLookNinki)
            If aRet <> "" Then
                aUpd = Split(aRet, ",")
                
                Debug.Print aUpd(0) & " , " & aUpd(1) & " , " & aUpd(2)
                
                gstrSql = ""
                gstrSql = gstrSql + "update analCmpi set "
                gstrSql = gstrSql + "Rcnt='" & aUpd(0) & "' , "
                gstrSql = gstrSql + "RtRate='" & aUpd(1) & "' , "
                gstrSql = gstrSql + "HtRate='" & aUpd(2) & "' "
                gstrSql = gstrSql + "where "
                gstrSql = gstrSql + "CmpiPtn ='" & aCmpiPtn & "' and "
                gstrSql = gstrSql + "LookNinki ='" & aLookNinki & "' "
                    
                db.Execute gstrSql, dbFailOnError
                
                aCnt = aCnt + 1
                Me.Caption = CStr(aCnt)
                Debug.Print Format$(Now, "hhnnss")
                
'                If aCnt = bbb Then
'                    Exit Do
'                End If
                
                Me.Refresh
                DoEvents
            End If
'StopTime = GetTickCount
'Debug.Print StopTime - StartTime
        End If
        
        Rs2.MoveNext
    Loop
    
    Rs2.Close
    
    MsgBox "finish"
    
End Sub

Private Sub Command18_Click()
    Dim aRtRate As String
    Dim aHtRate As String
    Dim aCnt As Long
    
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "analCmpiT1 "
    gstrSql = gstrSql + "where rcnt <> '' "
    gstrSql = gstrSql + "order by CmpiPtn, LookNinki "
    
    Set Rs2 = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
    Do
        If Rs2.EOF = False Then
            aRcnt = Rs2("Rcnt")
            aLookNinki = Rs2("LookNinki")
            aCmpiPtn = Rs2("CmpiPtn")
            aHtRate = Rs2("HtRate")
            aRtRate = Rs2("RtRate")
        Else
            Exit Do
        End If
        
        gstrSql = ""
        gstrSql = gstrSql + "update analCmpi set "
        gstrSql = gstrSql + "Rcnt='" & aRcnt & "' , "
        gstrSql = gstrSql + "RtRate='" & aRtRate & "' , "
        gstrSql = gstrSql + "HtRate='" & aHtRate & "' "
        gstrSql = gstrSql + "where "
        gstrSql = gstrSql + "CmpiPtn ='" & aCmpiPtn & "' and "
        gstrSql = gstrSql + "LookNinki ='" & aLookNinki & "' "
            
        db.Execute gstrSql, dbFailOnError
        
        Rs2.MoveNext
        
        aCnt = aCnt + 1
        Caption = aCnt
        Me.Refresh
        DoEvents
    Loop
    
    Rs2.Close
    
    MsgBox "finish"

End Sub


Private Sub Command19_Click()
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "HARAIX "
    ' テーブル名を指定してレコードセットを作成する
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
    Do
        If Rs.EOF = True Then
            Exit Do
        End If
        
        gstrSql = ""
        gstrSql = gstrSql + "SELECT "
        gstrSql = gstrSql + "* "
        gstrSql = gstrSql + "FROM "
        gstrSql = gstrSql + "HARAI "
        gstrSql = gstrSql + "where "
        gstrSql = gstrSql + "year='" & Rs("year") & "' and "
        gstrSql = gstrSql + "monthday='" & Rs("Monthday") & "' and "
        gstrSql = gstrSql + "JyoCD='" & Rs("JyoCD") & "' and "
        gstrSql = gstrSql + "RaceNum='" & Rs("RaceNum") & "' "
        ' テーブル名を指定してレコードセットを作成する
        Set Rs2 = db.OpenRecordset(gstrSql, dbOpenDynaset)
        
        If Rs2.EOF = False Then
        Else
            '追加
            If Rs("year") <> "" Then
            
                gstrSql = ""
                gstrSql = gstrSql + "insert into harai (Year, monthday, jyocd, racenum, "
                gstrSql = gstrSql + "PayFukusyoUmaban1, "
                gstrSql = gstrSql + "PayFukusyoPay1, "
                gstrSql = gstrSql + "PayFukusyoUmaban2, "
                gstrSql = gstrSql + "PayFukusyoPay2, "
                gstrSql = gstrSql + "PayFukusyoUmaban3, "
                gstrSql = gstrSql + "PayFukusyoPay3, "
                gstrSql = gstrSql + "PayFukusyoUmaban4, "
                gstrSql = gstrSql + "PayFukusyoPay4, "
                gstrSql = gstrSql + "PayFukusyoUmaban5, "
                gstrSql = gstrSql + "PayFukusyoPay5 "
                gstrSql = gstrSql + ") values ("
                
                gstrSql = gstrSql + "'" & Rs("year") & "', "
                gstrSql = gstrSql + "'" & Rs("monthday") & "', "
                gstrSql = gstrSql + "'" & Rs("jyocd") & "', "
                gstrSql = gstrSql + "'" & Rs("racenum") & "', "
                
                gstrSql = gstrSql + "'" & Rs("PayFukusyoUmaban1") & "', "
                gstrSql = gstrSql + "'" & Rs("PayFukusyoPay1") & "', "
                
                gstrSql = gstrSql + "'" & Rs("PayFukusyoUmaban2") & "', "
                gstrSql = gstrSql + "'" & Rs("PayFukusyoPay2") & "', "
                
                gstrSql = gstrSql + "'" & Rs("PayFukusyoUmaban3") & "', "
                gstrSql = gstrSql + "'" & Rs("PayFukusyoPay3") & "', "
                
                gstrSql = gstrSql + "'" & Rs("PayFukusyoUmaban4") & "', "
                gstrSql = gstrSql + "'" & Rs("PayFukusyoPay4") & "', "
                
                gstrSql = gstrSql + "'" & Rs("PayFukusyoUmaban5") & "', "
                gstrSql = gstrSql + "'" & Rs("PayFukusyoPay5") & "')"
                
                db.Execute gstrSql, dbFailOnError
            End If
        End If
        
        Rs2.Close

        Rs.MoveNext
    Loop

    Rs.Close

End Sub

Private Sub Command2_Click()
    myURL = "https://id.nikkansports.com/u/member/login/?guid=on&cid=53&premium=true&backurl=http://p.nikkansports.com/premium%2fj_spring_nikkan_security_check%3Fcurl%3dhttp%253A%252F%252Fp%2enikkansports%2ecom%252Fgoku-uma%252Fmember%252Findex%2ezpl&level=1"
   '起動中のIEを閉じる場合
   If Not ie Is Nothing Then
      ie.Quit
      Set ie = Nothing
   End If
   Set ie = New SHDocVw.InternetExplorer
   '指定のURLを表示
   ie.Navigate2 myURL
'    If chkD.Value = 1 And chkDL.Value = 0 Then
        ie.Visible = True    'IE を表示
'    End If
    
    Me.Caption = "Login start"
    Me.Refresh
    
    Do While ie.Busy = True Or ie.readyState <> 4
        DoEvents
    Loop

    Me.Caption = "Login comp"
    Me.Refresh

End Sub

Private Sub Command20_Click()
    Dim aWk(3) As String
    Dim ii As Integer
    Dim aCmpi As String
    
    'check value
    If Len(txtChk.Text) < 11 Then
        MsgBox "keyword is short." & vbCrLf & "example : 01-66_05-59"
        Exit Sub
    End If
    
    '01-66_05-59
    aWk(0) = Left$(txtChk.Text, 2)
    aWk(1) = Mid$(txtChk.Text, 4, 2)
    aWk(2) = Mid$(txtChk.Text, 7, 2)
    aWk(3) = Right$(txtChk.Text, 2)
    For ii = 0 To 3
        If IsNumeric(aWk(ii)) = False Then
            MsgBox "keyword is misstake." & vbCrLf & "example : 01-66_05-59"
            Exit Sub
        End If
        aWk(ii) = CStr(CInt(aWk(ii)))
    Next ii
    
    myURL = "http://p.nikkansports.com/goku-uma/member/compi/compi_db.zpl#/index/?CompiMin1=40&CompiMax1=90&CompiMin2=40&CompiMax2=90&CompiMin3=40&CompiMax3=90&"
    
    For ii = 1 To 18
        If ii = CInt(aWk(0)) Then
            aCmpi = aWk(1)
        ElseIf ii = CInt(aWk(2)) Then
            aCmpi = aWk(3)
        Else
            aCmpi = ""
        End If
        
        myURL = myURL & "Compi" & CStr(ii) & "=" & aCmpi & "&"
    Next ii
    
    myURL = myURL & "StartYear=2007&StartMonth=1&StartDay=1&EndYear=2016&EndMonth=12&EndDay=31&DistanceMin=0&DistanceMax=3600&BettingType=1&PayoffMin=0&PayoffMax=100000000&DiffCompiRankMin=&DiffCompiRankMax=&DiffMin=&DiffMax=&HeadsMin=&HeadsMax="
   
   '指定のURLを表示
   ie.Navigate2 myURL
    Do While ie.Busy = True Or ie.readyState <> 4
        DoEvents
    Loop
   
    For Each objTag In ie.Document.getElementsByTagName("input")

        If InStr(objTag.outerHTML, "検索開始") > 0 Then

            '送信ボタンクリック
            objTag.Click

            Do While ie.Busy = True Or ie.readyState <> 4
                DoEvents
            Loop

            'ループ脱出
            Exit For
              
        End If
    Next

End Sub

Private Sub Command21_Click()
    txtRes.Text = ""
End Sub

Private Sub Command22_Click()
    myURL = "https://id.nikkansports.com/u/member/login/?guid=on&cid=53&premium=true&backurl=http://p.nikkansports.com/premium%2fj_spring_nikkan_security_check%3Fcurl%3dhttp%253A%252F%252Fp%2enikkansports%2ecom%252Fgoku-uma%252Fmember%252Findex%2ezpl&level=1"
   '起動中のIEを閉じる場合
   If Not ie Is Nothing Then
      ie.Quit
      Set ie = Nothing
   End If
   Set ie = New SHDocVw.InternetExplorer
   '指定のURLを表示
   ie.Navigate2 myURL
'    If chkD.Value = 1 And chkDL.Value = 0 Then
        ie.Visible = True    'IE を表示
'    End If
    
    Me.Caption = "Login start"
    Me.Refresh
    
    Do While ie.Busy = True Or ie.readyState <> 4
        DoEvents
    Loop

    Me.Caption = "Login comp"
    Me.Refresh

End Sub


Private Sub Command23_Click()
    Dim resUma(41) As String
    Dim resPay(41) As String
    Dim aCmpiValue As String
    Dim aUmaban As String
    
    aYear = Format$(Now, "yyyy")
    aMonthday = Format$(Now, "mmdd")
    
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "uma_RACE "
    gstrSql = gstrSql + "where CmpiNinki <> '' and "
    gstrSql = gstrSql + "Year ='" & aYear & "' and "
    gstrSql = gstrSql + "Monthday ='" & aMonthday & "' "
    gstrSql = gstrSql + "ORDER BY "
    gstrSql = gstrSql + "Year, MonthDay, JyoCD, RaceNum, CmpiNinki"
    
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
    Do
        If Rs.EOF = False Then
            aCmpiNinki = Rs("CmpiNinki")
            aCmpiValue = Rs("CmpiValue")
            aUmaban = Rs("Umaban")
            
            If aCmpiNinki = "01" Then
                aYear = Rs("year")
                aMonthday = Rs("MonthDay")
                aJyoCd = Rs("JyoCD")
                aRaceNum = Rs("RaceNum")
                gYear = aYear
                gMonthDay = aMonthday
                gJyoCD = aJyoCd
                gRaceNum = aRaceNum
                ret = getHarai(resUma, resPay)
                'データベースに追加
                gstrSql = ""
                gstrSql = gstrSql + "insert into uma_cmpi (Year, monthday, jyocd, racenum, "
                gstrSql = gstrSql + "C01, U01, "
                gstrSql = gstrSql + "HF01, HF02, HF03, HF04, HF05, "
                gstrSql = gstrSql + "UF01, UF02, UF03, UF04, UF05 "
                gstrSql = gstrSql + ") values ("
                
                gstrSql = gstrSql + "'" & aYear & "', "
                gstrSql = gstrSql + "'" & aMonthday & "', "
                gstrSql = gstrSql + "'" & aJyoCd & "', "
                gstrSql = gstrSql + "'" & aRaceNum & "', "
                gstrSql = gstrSql + "'" & aCmpiValue & "', "
                gstrSql = gstrSql + "'" & aUmaban & "', "
                gstrSql = gstrSql + "'" & resPay(3) & "', "
                gstrSql = gstrSql + "'" & resPay(4) & "', "
                gstrSql = gstrSql + "'" & resPay(5) & "', "
                gstrSql = gstrSql + "'" & resPay(6) & "', "
                gstrSql = gstrSql + "'" & resPay(7) & "', "
                gstrSql = gstrSql + "'" & resUma(3) & "', "
                gstrSql = gstrSql + "'" & resUma(4) & "', "
                gstrSql = gstrSql + "'" & resUma(5) & "', "
                gstrSql = gstrSql + "'" & resUma(6) & "', "
                gstrSql = gstrSql + "'" & resUma(7) & "')"
                
                db.Execute gstrSql, dbFailOnError
                
            Else
                'データベース更新
                gstrSql = ""
                gstrSql = gstrSql + "update uma_cmpi set "
                gstrSql = gstrSql + "C" & aCmpiNinki & "='" & aCmpiValue & "' , "
                gstrSql = gstrSql + "U" & aCmpiNinki & "='" & aUmaban & "' "
                gstrSql = gstrSql + "where "
                gstrSql = gstrSql + "Year ='" & aYear & "' and "
                gstrSql = gstrSql + "Monthday ='" & aMonthday & "' and "
                gstrSql = gstrSql + "JyoCD ='" & aJyoCd & "' and "
                gstrSql = gstrSql + "RaceNum ='" & aRaceNum & "' "
                    
                db.Execute gstrSql, dbFailOnError
            End If
            
            Rs.MoveNext
        Else
            Exit Do
        End If
    Loop
    
    Rs.Close
    
    MsgBox "job finish!"

End Sub

Private Sub Command24_Click()
    Command24.Enabled = False
    Dim aNinki As String
    Dim aNinki2 As String
    Dim aVal As String
    Dim aVal2 As String
    Dim aHit As Boolean
    
    aYear = Format$(Now, "yyyy")
    aMonthday = Format$(Now, "mmdd")
    
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "uma_cmpi "
    gstrSql = gstrSql + "where "
    gstrSql = gstrSql + "Year ='" & aYear & "' and "
    gstrSql = gstrSql + "Monthday ='" & aMonthday & "' "
    gstrSql = gstrSql + "order by JyoCD, racenum"
    
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
    Do
        If Rs.EOF = False Then
            aHit = False
            
            For ii = 1 To 17
                aNinki = Format$(ii, "00")
                aVal = Rs("C" & aNinki)
                If aVal <> "" Then
                    For jj = ii + 1 To 18
                        'コンピ順位 ii, jjの組合せ
                        aNinki2 = Format$(jj, "00")
                        aVal2 = Rs("C" & aNinki2)
                        
                        gstrSql = ""
                        gstrSql = gstrSql + "SELECT "
                        gstrSql = gstrSql + "* "
                        gstrSql = gstrSql + "FROM "
                        gstrSql = gstrSql + "analCmpi "
                        gstrSql = gstrSql + "where "
                        gstrSql = gstrSql + "CmpiPtn ='" & "_" & aNinki & "-" & aVal & "_" & aNinki2 & "-" & aVal2 & "' "            '_01-64_02-62
                        
                        Set Rs2 = db.OpenRecordset(gstrSql, dbOpenDynaset)
                        If Rs2.EOF = False Then
                            txtRes.Text = txtRes.Text & " jyoc:" & Rs("JyoCD") & vbCrLf
                            txtRes.Text = txtRes.Text & " race:" & Rs("racenum") & vbCrLf
                            txtRes.Text = txtRes.Text & " rcnt:" & Rs2("rcnt") & vbCrLf
                            txtRes.Text = txtRes.Text & " hit :" & Rs2("htrate") & vbCrLf
                            txtRes.Text = txtRes.Text & " ret :" & Rs2("rtrate") & vbCrLf
                            txtRes.Text = txtRes.Text & vbCrLf
                            
                            aHit = True
                            Exit For
                        End If
                        
                        Rs2.Close
                        
                    Next jj
                End If
                If aHit = True Then
                    Exit For
                End If
            Next ii
        Else
            Exit Do
        End If
        
        Rs.MoveNext
    Loop
    
    Rs.Close
    
    Command24.Enabled = True
    MsgBox "job finish!"

End Sub

Private Sub Command25_Click()
    Dim aCmpiNinki(1) As String
    Dim aCmpiValue(1) As String
    Dim aYear As String
    Dim aMonthday As String
    
'レイトチェックアウトは1時間1,620円
'￥39,366
'
'Ptn:_01-71_07-52
'Nin:07
'Rac:000154
'Ret:000121.753
'Hit:024.026
'
'Ptn:_06-57_07-51
'Nin:07
'Rac:000102
'Ret:000125.980
'Hit:021.569
    '01-66_05-59
    aCmpiNinki(0) = Left$(txtSelRecipe.Text, 2)
    aCmpiNinki(1) = Mid$(txtSelRecipe.Text, 7, 2)
    aCmpiValue(0) = Mid$(txtSelRecipe.Text, 4, 2)
    aCmpiValue(1) = Right$(txtSelRecipe.Text, 2)
    
    aYear = areaY.Text
    aMonthday = areaMD.Text
'    aYear = Format$(Now, "yyyy")
'    aMonthday = Format$(Now, "mmdd")
    
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "uma_cmpi "
    gstrSql = gstrSql + "where "
    gstrSql = gstrSql + "Year ='" & aYear & "' and "
    gstrSql = gstrSql + "Monthday ='" & aMonthday & "' and "
    gstrSql = gstrSql + "C" & aCmpiNinki(0) & "='" & aCmpiValue(0) & "' and "
    gstrSql = gstrSql + "C" & aCmpiNinki(1) & "='" & aCmpiValue(1) & "' "
    gstrSql = gstrSql + "ORDER BY "
    gstrSql = gstrSql + "JyoCD, RaceNum"
    
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
    Do
        If Rs.EOF = False Then
            txtRes.Text = txtRes.Text & txtSelRecipe.Text & ", " & "Jyo: " & Rs("jyocd") & ", Race: " & Rs("RaceNum") & ", Umaban: " & Rs("U" & aCmpiNinki(1)) & vbCrLf
            Print #fn, txtSelRecipe.Text & ", " & "" & Rs("jyocd") & ", " & Rs("RaceNum") & ", " & Rs("U" & aCmpiNinki(1))
        Else
            
            Exit Do
        End If
        
        Rs.MoveNext
    Loop
        
End Sub

Private Sub Command26_Click()
    Dim aaa() As String
    Dim iii As Integer
    
    aaa = Split(Text1.Text, vbCrLf)
    
    fn = FreeFile
    Open App.Path & "\result_" & Format$(Now, "yyyymmddhhnnss") & ".csv" For Output As #fn
    
    Print #fn, "recepe"; t & ", " & "Jyo" & ", Race" & ", Umaban"
    
    For iii = 0 To UBound(aaa)
        If aaa(iii) <> "" Then
            txtSelRecipe.Text = aaa(iii)
            Call Command25_Click
        End If
    Next iii
    
    Close #fn
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
    
    Dim aDrawDate As String
    Dim aStartDate As Date
    Dim aDays As Long
    Dim win As Long
    Dim total As Long
    Dim mny As Long
    Dim betTotal As Long
    Dim betWin As Long
    Dim resUma(41) As String
    Dim resPay(41) As String
    Dim umauma(5) As String
    Dim aSelUma As String
    Dim aChoiceFlg As Integer
    Dim aChoiceMax As Integer
    Dim aNinki As Integer
    Dim aKariWin As Long
    Dim aKariMny As Long
    Dim aKariBetTotal As Long
    Dim aMsg As String
    Dim aWebMsg As String
    Dim aWk As String
    Dim aRaceNum As String
    Dim aLog As String
    Dim aWork As String
    Dim aMaxRenpai(36500) As Long
    Dim aRenpaiCnt As Long
    Dim aNin As Integer
    Dim aVal As Integer
    Dim aHit As Boolean
    Dim ii As Long
    
    Dim aNin2 As Integer
    Dim aVal2 As Integer
    
    aNin2 = 16
    aVal2 = 46

For aNin = aNin2 To aNin2
    For aVal = aVal2 To aVal2 Step -1
        
        fn = FreeFile
        Open App.Path & "\dic\" & Format$(aNin, "00") & "_" & Format$(aVal, "00") & "_harai.csv" For Output As #fn
        
        aRenpaiCnt = 0
        total = 0
        win = 0
        mny = 0
        
        For ii = 1 To 36500
            aMaxRenpai(aRenpaiCnt) = 0
        Next ii
        
        gstrSql = ""
        gstrSql = gstrSql + "SELECT "
        gstrSql = gstrSql + "RACE.SyussoTosu,race.HassoTime, UMA_RACE.Year, UMA_RACE.MonthDay, UMA_RACE.JyoCD,UMA_RACE.racenum,  "
        gstrSql = gstrSql + "UMA_RACE.KakuteiJyuni,  UMA_RACE.Ninki, UMA_RACE.TanOdds5, UMA_RACE.TanNinki5, UMA_RACE.TanNinki1, UMA_RACE.fukuninki1, UMA_RACE.FOdds, UMA_RACE.Umaban, UMA_RACE.DMJyuni, UMA_RACE.CmpiNinki,UMA_RACE.hensaCmpi,UMA_RACE.CmpiValue "
        gstrSql = gstrSql + "FROM "
        gstrSql = gstrSql + "RACE INNER JOIN "
        gstrSql = gstrSql + "UMA_RACE ON "
        gstrSql = gstrSql + "(RACE.RaceNum = UMA_RACE.RaceNum) AND "
        gstrSql = gstrSql + "(RACE.JyoCD = UMA_RACE.JyoCD) AND "
        gstrSql = gstrSql + "(RACE.MonthDay = UMA_RACE.MonthDay) AND "
        gstrSql = gstrSql + "(RACE.Year = UMA_RACE.Year) "
        gstrSql = gstrSql + "WHERE "
        gstrSql = gstrSql + "(UMA_RACE.year<='2016') and "
        gstrSql = gstrSql + "(UMA_RACE.JyoCD<='10')  and "
        gstrSql = gstrSql + "(UMA_RACE.CmpiNinki='" & Format$(aNin, "00") & "') and "
        gstrSql = gstrSql + "(UMA_RACE.CmpiValue='" & Format$(aVal, "00") & "') "
    
        gstrSql = gstrSql + "ORDER BY "
        gstrSql = gstrSql + "UMA_RACE.Year, UMA_RACE.MonthDay, UMA_RACE.JyoCD, UMA_RACE.RaceNum, UMA_RACE.CmpiNinki"
        ' テーブル名を指定してレコードセットを作成する
        Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
        Do
            If Rs.EOF = False Then
                
                aYear = Rs("year")
                aMonthday = Rs("MonthDay")
                aJyoCd = Rs("JyoCD")
                aRaceNum = Rs("RaceNum")
                gYear = aYear
                gMonthDay = aMonthday
                gJyoCD = aJyoCd
                gRaceNum = aRaceNum
                umauma(0) = Rs("umaban")
                
                total = total + 1
                
                aWork = "0"
                aHit = False
                
                ret = getHarai(resUma, resPay)
                If ret = 0 Then
                    For idx2 = 3 To 7
                        If resUma(idx2) = umauma(0) Then
                            '次のレースのデータの時と最後に清算
                            If aRenpaiCnt > 0 Then
                                aMaxRenpai(aRenpaiCnt) = aMaxRenpai(aRenpaiCnt) + 1
                                aRenpaiCnt = 0
                            End If
                            
                            aHit = True
                            aKariWin = 1
                            win = win + 1
                            mny = mny + CLng(resPay(idx2))
                            aWork = CStr(CLng(resPay(idx2)))
                            Exit For
                        End If
                    Next idx2
                End If
                
                If aHit = False Then
                    aRenpaiCnt = aRenpaiCnt + 1
                End If
                
                Print #fn, aWork
    '            aLog = aLog & "," & aWork
                
            Else
                Exit Do
            End If
            
            Rs.MoveNext
        
        Loop
        
        Rs.Close
        
        Close #fn
        
        fn = FreeFile
        Open App.Path & "\dic\" & Format$(aNin, "00") & "_" & Format$(aVal, "00") & "_renpai.csv" For Output As #fn
        For ii = 1 To 36500
            Print #fn, str(aMaxRenpai(ii))
        Next ii
        
        Close #fn
        
        
        Me.Caption = Format$(aNin, "00") & "_" & Format$(aVal, "00")
        Me.Refresh
        DoEvents
        
    Next aVal
Next aNin
    
    MsgBox "finish!"
    
    Command27.Enabled = True

End Sub

Private Sub Command28_Click()
    Command28.Enabled = False
    
    Dim aDrawDate As String
    Dim aStartDate As Date
    Dim aDays As Long
    Dim win As Long
    Dim total As Long
    Dim mny As Long
    Dim betTotal As Long
    Dim betWin As Long
    Dim resUma(41) As String
    Dim resPay(41) As String
    Dim umauma(5) As String
    Dim aSelUma As String
    Dim aChoiceFlg As Integer
    Dim aChoiceMax As Integer
    Dim aNinki As Integer
    Dim aKariWin As Long
    Dim aKariMny As Long
    Dim aKariBetTotal As Long
    Dim aMsg As String
    Dim aWebMsg As String
    Dim aWk As String
    Dim aRaceNum As String
    Dim aLog As String
    Dim aWork As String
    Dim aMaxRenpai(36500) As Long
    Dim aRenpaiCnt As Long
    Dim aNin As Integer
    Dim aVal As Integer
    Dim aHit As Boolean
    Dim aRenHit As Boolean
    Dim ii As Long
    Dim aRenRen As Long
    Dim aNin2 As Integer
    Dim aVal2 As Integer
    
    aRenRen = 289
    aNin2 = 18
    aVal2 = 41
    
For aNin = aNin2 To aNin2
    For aVal = aVal2 To aVal2 Step -1
        
        fn = FreeFile
        Open App.Path & "\dic\" & Format$(aNin, "00") & "_" & Format$(aVal, "00") & "_harai_" & CStr(aRenRen) & ".csv" For Output As #fn
        
        aRenpaiCnt = 0
        total = 0
        win = 0
        mny = 0
        
        For ii = 1 To 36500
            aMaxRenpai(aRenpaiCnt) = 0
        Next ii
        
        gstrSql = ""
        gstrSql = gstrSql + "SELECT "
        gstrSql = gstrSql + "RACE.SyussoTosu,race.HassoTime, UMA_RACE.Year, UMA_RACE.MonthDay, UMA_RACE.JyoCD,UMA_RACE.racenum,  "
        gstrSql = gstrSql + "UMA_RACE.KakuteiJyuni,  UMA_RACE.Ninki, UMA_RACE.TanOdds5, UMA_RACE.TanNinki5, UMA_RACE.TanNinki1, UMA_RACE.fukuninki1, UMA_RACE.FOdds, UMA_RACE.Umaban, UMA_RACE.DMJyuni, UMA_RACE.CmpiNinki,UMA_RACE.hensaCmpi,UMA_RACE.CmpiValue "
        gstrSql = gstrSql + "FROM "
        gstrSql = gstrSql + "RACE INNER JOIN "
        gstrSql = gstrSql + "UMA_RACE ON "
        gstrSql = gstrSql + "(RACE.RaceNum = UMA_RACE.RaceNum) AND "
        gstrSql = gstrSql + "(RACE.JyoCD = UMA_RACE.JyoCD) AND "
        gstrSql = gstrSql + "(RACE.MonthDay = UMA_RACE.MonthDay) AND "
        gstrSql = gstrSql + "(RACE.Year = UMA_RACE.Year) "
        gstrSql = gstrSql + "WHERE "
        gstrSql = gstrSql + "(UMA_RACE.year<='2016') and "
        gstrSql = gstrSql + "(UMA_RACE.JyoCD<='10')  and "
        gstrSql = gstrSql + "(UMA_RACE.CmpiNinki='" & Format$(aNin, "00") & "') and "
        gstrSql = gstrSql + "(UMA_RACE.CmpiValue='" & Format$(aVal, "00") & "') "
    
        gstrSql = gstrSql + "ORDER BY "
        gstrSql = gstrSql + "UMA_RACE.Year, UMA_RACE.MonthDay, UMA_RACE.JyoCD, UMA_RACE.RaceNum, UMA_RACE.CmpiNinki"
        ' テーブル名を指定してレコードセットを作成する
        Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
        Do
            If Rs.EOF = False Then
                
                aYear = Rs("year")
                aMonthday = Rs("MonthDay")
                aJyoCd = Rs("JyoCD")
                aRaceNum = Rs("RaceNum")
                gYear = aYear
                gMonthDay = aMonthday
                gJyoCD = aJyoCd
                gRaceNum = aRaceNum
                umauma(0) = Rs("umaban")
                
                aRenHit = False
                If aRenpaiCnt = aRenRen Then
                    total = total + 1
                    aRenHit = True
                End If
                
                aWork = "0"
                aHit = False
                
                ret = getHarai(resUma, resPay)
                If ret = 0 Then
                    For idx2 = 3 To 7
                        If resUma(idx2) = umauma(0) Then
                            If aRenHit = True Then
                                win = win + 1
                                mny = mny + CLng(resPay(idx2))
                                aWork = CStr(CLng(resPay(idx2)))
                                Debug.Print aYear & ", " & aMonthday & ", " & aJyoCd & ", " & aRaceNum
                            End If
                            
                            '次のレースのデータの時と最後に清算
                            aRenpaiCnt = 0
                            
                            aHit = True
                            Exit For
                        End If
                    Next idx2
                End If
                
                If aHit = False Then
                    aRenpaiCnt = aRenpaiCnt + 1
                End If
                
                If aRenHit = True Then
                    Print #fn, aWork
                    aRenpaiCnt = 0
                End If
    '            aLog = aLog & "," & aWork
                
            Else
                Exit Do
            End If
            
            Rs.MoveNext
        
        Loop
        
        Rs.Close
        
        Close #fn
        
        Me.Caption = Format$(aNin, "00") & "_" & Format$(aVal, "00")
        Me.Refresh
        DoEvents
        
    Next aVal
Next aNin
    
    MsgBox "finish!"
    
    Command28.Enabled = True

End Sub

Private Sub Command29_Click()
    Dim dd() As String
  Dim Stream As Object
    Dim str As String
    Dim str2 As String
    Dim kbn As Integer
    
    Command3.Enabled = False
    
    kbn = 0
    '指定のURLを表示
    myURL = "http://ikumen.buhi-buhi.com/"
    
    ie.Navigate2 myURL
'    If chkD.Value = 1 Then
        ie.Visible = True    'IE を表示
'    End If

    Me.Caption = "Year start"
    Me.Refresh

    Do While ie.Busy = True Or ie.readyState <> 4
        DoEvents
    Loop


Dim objLink As Object
 
    For Each objLink In ie.Document.getElementsByTagName("A")
        If objLink.innerText = "にほんブログ村" Then
            ie.navigate objLink.href
            Exit For
        End If
    Next
End Sub

Private Sub Command3_Click()
    Dim dd() As String
  Dim Stream As Object
    Dim str As String
    Dim str2 As String
    Dim kbn As Integer
    
    Dim ii As Integer
    Dim fcnt As Integer
    
    For ii = 0 To 5
        gFilename(ii) = ""
    Next ii
    List1.Clear
    
    Command3.Enabled = False
    
    kbn = 0
    '指定のURLを表示
    myURL = "http://p.nikkansports.com/goku-uma/member/compi/compi_list.zpl?year=2016&mode=kako"
    
    ie.Navigate2 myURL
'    If chkD.Value = 1 Then
        ie.Visible = True    'IE を表示
'    End If

    Me.Caption = "Year start"
    Me.Refresh

    Do While ie.Busy = True Or ie.readyState <> 4
        DoEvents
    Loop

    Me.Caption = "Year comp"
    Me.Refresh
    '
    '年のURLを取得する
    str = getHTMLString(ie)
    If str = "" Then
'        GoTo exit_here
    End If
    gStr = str
    Call getYear(kbn)

    '年ループ   gYear gUrlYear
    For aYear = 0 To UBound(gYears)
'    For aYear = 0 To UBound(gYear)
        If gYears(aYear) = areaY.Text Or areaY.Text = "" Then
            '指定の年サイトに移動
            myURL = gUrlYear(aYear)
            ie.Navigate2 myURL

'            If chkD.Value = 1 Then
                ie.Visible = True    'IE を表示
'            End If

            Me.Caption = "Year Get start"
            Me.Refresh

            Do While ie.Busy = True Or ie.readyState <> 4
                DoEvents
            Loop

            Me.Caption = "Year Get comp"
            Me.Refresh

            'すべての日付のURLを取得する
            str = getHTMLString(ie)
            If str = "" Then
'                GoTo exit_here
            End If
            gStr = str
            Call getDay(kbn)
            '日付ループ gDay gPosDay
            For aDay = 0 To UBound(gUrlDay)
                If Right$(gUrlDay(aDay), 4) = areaMD.Text Or areaMD.Text = "" Then
                
    '            For aDay = 0 To UBound(gDay)
                    
                    
                    myURL = gUrlDay(aDay)
                    ie.Navigate2 myURL
    
    '                If chkD.Value = 1 Then
                        ie.Visible = True    'IE を表示
    '                End If
    
                    Me.Caption = "Day start"
                    Me.Refresh
    
                    Do While ie.Busy = True Or ie.readyState <> 4
                        DoEvents
                    Loop
    
                    Me.Caption = "Day comp"
                    Me.Refresh
    
                    '全レースのURLを取得する +結果
                    str = getHTMLString(ie)
                    If str = "" Then
                    End If
                    gStr = str
                    
                    str2 = Replace(str, vbLf, vbCr & vbLf)
                    ff = App.Path & "\" & gYears(aYear) & gDayFmt(aDay) & gPosDay(aDay)
                    Call FilePutContents(ff & ".txt", str2, "utf-8")
                    gFilename(fcnt) = ff & ".txt"
                    List1.AddItem gFilename(fcnt)
                    fcnt = fcnt + 1
                End If
            Next aDay

        End If
    Next aYear

    Command3.Enabled = True
    MsgBox "job finish!"
End Sub

' 指定されたファイルに指定された文字列を出力する
Public Sub FilePutContents(ByVal sFileName As String, sBuffer As String, Optional sEncoding As String, Optional bSaveToWorkbookPath As Boolean)
    Dim oFso As Object
    Dim oFile As Object
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
    
    ' フラグが指定された場合はワークブックのパスに保存する
    If bSaveToWorkbookPath Then
        sFileName = oFso.GetParentFolderName(ActiveWorkbook.FullName) + "\" + sFileName
    End If

    If sEncoding <> "" Then
        ' エンコーディングが指定された場合は ADODB.Stream を利用して文字コードを変換する
        Dim oAdo As Object
        Set oAdo = CreateObject("ADODB.Stream")
        oAdo.Type = 2 'adTypeText
        oAdo.Charset = sEncoding
        
        oAdo.Open
        oAdo.WriteText sBuffer
        
        ' UTF-8 であれば BOM つきで出力されているはずなので削る
        If LCase(sEncoding) = "utf-8" Then
            ' 出力された BOM をスキップして読み込み直す
            oAdo.position = 0   ' Type の変更には Position が 0 である必要あり
            oAdo.Type = 1 'adTypeBinary
            oAdo.position = 3   ' 先頭の 3 bytes（BOM）をスキップ
            Dim sEncodedBuffer As Variant
            sEncodedBuffer = oAdo.Read()
            
            ' ストリームの先頭に戻って内容を再度書きだす
            oAdo.position = 0
            oAdo.write sEncodedBuffer
            oAdo.SetEos     ' ストリームの最後にゴミが残っているので削る
        End If
        oAdo.SaveToFile (sFileName), 2 'adSaveCreateOverWrite
        oAdo.Close
    Else
        ' エンコーディングが指定されていない場合は FileSystemObject で出力する
        Set oFile = oFso.CreateTextFile(sFileName, True)
        oFile.write sBuffer
        oFile.Close
    End If
End Sub

Private Sub getDay(kbn As Integer)
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbLf & "^"
'    kaigyo = vbCr & "$" & vbLf & "^"
'    kaigyo = vbCr & vbLf
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True

    If kbn = 0 Then
'        'コンピ指数
        .Pattern = "<dt>[0-9]+月[0-9]+日.+" & vbLf & ".+" & vbLf & ".+" & vbLf & ".+" & vbLf & ".+" & vbLf & ".+" & vbLf
        '結果
'        .Pattern = "a href.+[0-9]+月[0-9]+日\("
    Else
        .Pattern = "<dt>[0-9]+月[0-9]+日.+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+nbsp;"
    End If
    
    Dim aWk As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.value
        cnt = cnt + 1
        
        ReDim Preserve gDay(cnt)
        ReDim Preserve gDayFmt(cnt)
        ReDim Preserve gUrlDay(cnt)
        ReDim Preserve gPosDay(cnt)
        ReDim Preserve gPosDayCd(cnt)
        ReDim Preserve gPosDayDbCd(cnt)
        
        gWk = Mid$(retstr, 5)
        gWk = Left$(gWk, InStr(gWk, "<") - 1)

        gDay(cnt) = gWk

        aWk = Format$(Left$(gDay(cnt), InStr(gDay(cnt), "月") - 1), "00")
        gDayFmt(cnt) = aWk
        aWk = Format$(Mid$(gDay(cnt), InStr(gDay(cnt), "月") + 1, InStr(gDay(cnt), "日") - (InStr(gDay(cnt), "月") + 1)), "00")
        gDayFmt(cnt) = gDayFmt(cnt) & aWk
        
        If kbn = 0 Then
'            'chuuou
''            '結果
''            gWk = Mid$(retstr, 5)
''            gWk = Left$(gWk, InStr(gWk, "<") - 1)
''
''            gDay(cnt) = gWk
''
''            aWk = Format$(Left$(gDay(cnt), InStr(gDay(cnt), "月") - 1), "00")
''            gDayFmt(cnt) = aWk
''            aWk = Format$(Mid$(gDay(cnt), InStr(gDay(cnt), "月") + 1, InStr(gDay(cnt), "日") - (InStr(gDay(cnt), "月") + 1)), "00")
''            gDayFmt(cnt) = gDayFmt(cnt) & aWk
'
'            gWk = Mid$(retstr, 60)
'            gWk = Left$(gWk, Len(gWk) - 1)
'
'            gDay(cnt) = gWk
'
'            aWk = Format$(Left$(gDay(cnt), InStr(gDay(cnt), "月") - 1), "00")
'            gDayFmt(cnt) = aWk
'            aWk = Format$(Mid$(gDay(cnt), InStr(gDay(cnt), "月") + 1, InStr(gDay(cnt), "日") - (InStr(gDay(cnt), "月") + 1)), "00")
'            gDayFmt(cnt) = gDayFmt(cnt) & aWk
'
'
'            aWk = Mid$(retstr, InStr(retstr, "course_id=") + 33)
'            aWk = Mid$(aWk, 1, InStr(aWk, "&nbsp") - 1)
'            gPosDay(cnt) = aWk

            'http://p.nikkansports.com/goku-uma/member/result/result_day-list.zpl?date=20111211&mode=kako
'            aWk = Mid$(retstr, InStr(retstr, "a href=") + 9, 48)
'            gUrlDay(cnt) = Replace("http://p.nikkansports.com/goku-uma/member/result" & aWk, "amp;", "")

            'コンピ指数
            aWk = retstr 'Mid$(retstr, InStr(retstr, gPosDay(cnt)) + Len(gPosDay(cnt)))
            'http://p.nikkansports.com/goku-uma/member/compi/compi.zpl?course_id=005&date=20161106
            gUrlDay(cnt) = Mid$(aWk, InStr(aWk, "a href=") + 9, 42)
            gUrlDay(cnt) = Replace("http://p.nikkansports.com/goku-uma/member/compi" & gUrlDay(cnt), "amp;", "")
            aWk = Mid$(aWk, InStr(aWk, "course_id=") + 33)
            aWk = Mid$(aWk, 1, InStr(aWk, "&nbsp") - 1)
            gPosDay(cnt) = aWk
            aWk = Mid$(retstr, InStr(retstr, gPosDay(cnt)) + Len(gPosDay(cnt)))
            
            Do
                If InStr(aWk, "course_id=") > 0 Then
                    cnt = cnt + 1

                    ReDim Preserve gDay(cnt)
                    gDay(cnt) = gDay(cnt - 1)
                    ReDim Preserve gDayFmt(cnt)
                    gDayFmt(cnt) = gDayFmt(cnt - 1)

                    ReDim Preserve gUrlDay(cnt)
                    ReDim Preserve gPosDay(cnt)
                    ReDim Preserve gPosDayCd(cnt)
                    ReDim Preserve gPosDayDbCd(cnt)

                    'コンピ指数
                    'http://p.nikkansports.com/goku-uma/member/compi/compi.zpl?course_id=005&date=20161106
                    gUrlDay(cnt) = Mid$(aWk, InStr(aWk, "a href=") + 9, 42)
                    gUrlDay(cnt) = Replace("http://p.nikkansports.com/goku-uma/member/compi" & gUrlDay(cnt), "amp;", "")
                    aWk = Mid$(aWk, InStr(aWk, "course_id=") + 33)
                    aWk = Mid$(aWk, 1, InStr(aWk, "&nbsp") - 1)
                    gPosDay(cnt) = aWk
                    aWk = Mid$(retstr, InStr(retstr, gPosDay(cnt)) + Len(gPosDay(cnt)))

                Else
                   Exit Do
                End If
            Loop
        Else
            'nankan
            gWk = Mid$(retstr, 5)
            gWk = Left$(gWk, InStr(gWk, "<") - 1)
            
            gDay(cnt) = gWk
            
            aWk = Format$(Left$(gDay(cnt), InStr(gDay(cnt), "月") - 1), "00")
            gDayFmt(cnt) = aWk
            aWk = Format$(Mid$(gDay(cnt), InStr(gDay(cnt), "月") + 1, InStr(gDay(cnt), "日") - (InStr(gDay(cnt), "月") + 1)), "00")
            gDayFmt(cnt) = gDayFmt(cnt) & aWk
            
            gWk = Mid$(retstr, InStr(retstr, "/goku-uma"))
            aWk = Mid$(gWk, InStr(gWk, "kako") + 6)
            aWk = Left$(aWk, Len(aWk) - 6)
            gPosDay(cnt) = aWk
            
            Select Case gPosDay(cnt)
            Case "浦和"
                gPosDayCd(cnt) = "18"
                gPosDayDbCd(cnt) = "42"
            Case "船橋"
                gPosDayCd(cnt) = "19"
                gPosDayDbCd(cnt) = "43"
            Case "大井"
                gPosDayCd(cnt) = "20"
                gPosDayDbCd(cnt) = "44"
            Case "川崎"
                gPosDayCd(cnt) = "21"
                gPosDayDbCd(cnt) = "45"
            End Select
            
            gWk = Left$(gWk, InStr(gWk, "kako") + 3)
            'http://p.nikkansports.com/goku-uma/member/races/pdf_list_top_nankan.zpl?date_place_id=244472&mode=kako
            gUrlDay(cnt) = "http://p.nikkansports.com" & gWk
        End If
'        Debug.Print gUrlDay(cnt)
    Next
    
End With
'    Debug.Print "day"

End Sub



Private Sub getYear(kbn As Integer)
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbCr & "$" & vbLf & "^"
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    
    If kbn = 0 Then
'         .Pattern = "result_list\.zpl\?year=....&amp;mode=kako"">....年</a>"
         .Pattern = "compi_list\.zpl\?year=....&amp;mode=kako"">....年</a>"
    Else
        'a href="past_list_nankan.zpl?year=2016&mode=kako">2016年</a>
         .Pattern = "past_list_nankan\.zpl\?year=....&amp;mode=kako"">....年</a>"
    End If
    
    Dim firstY As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.value
        cnt = cnt + 1
        If firstY = "" Then
            firstY = retstr
        Else
            If firstY = retstr Then
                Exit For
            End If
        End If
        ReDim Preserve gYears(cnt)
        ReDim Preserve gUrlYear(cnt)
        
        gYears(cnt) = Left$(Right$(retstr, 9), 4)
        '南関競馬の結果
        'http://p.nikkansports.com/goku-uma/member/races/past_list_nankan.zpl?year=2015&mode=kako
        '中央競馬のコンピ指数
        'http://p.nikkansports.com/goku-uma/member/compi/compi_list.zpl?year=2015&mode=kako
        '中央競馬の結果
        'http://p.nikkansports.com/goku-uma/member/result/result_list.zpl?year=2015&mode=kako
        If kbn = 0 Then
'            'コンピ指数
            gUrlYear(cnt) = "http://p.nikkansports.com/goku-uma/member/compi/" & Left$(retstr, Len(retstr) - 11)
            '結果
'            gUrlYear(cnt) = "http://p.nikkansports.com/goku-uma/member/result/" & Left$(retstr, Len(retstr) - 11)
            gUrlYear(cnt) = Replace(gUrlYear(cnt), "amp;", "")
        Else
            '結果
            gUrlYear(cnt) = "http://p.nikkansports.com/goku-uma/member/races/" & Left$(retstr, Len(retstr) - 11)
        End If
        
    Next
    
End With

End Sub

Private Function getHTMLString(ie As Object) As String
On Error GoTo err_handler
    Dim htdoc As HTMLDocument
    Set htdoc = ie.Document
    
    Dim ret As String
    ret = htdoc.getElementsByTagName("HTML")(0).outerHTML & vbCrLf
    getHTMLString = ret
    
    Exit Function
err_handler:
    
    Debug.Print err.Description
    
    getHTMLString = ""
    
    Exit Function
End Function

Private Sub Command30_Click()
    Dim aRace(3, 13, 20, 20) As String    '開催場所、レース、対象１番、対象２番
    Dim aHit(3, 13, 20, 20) As String
    Dim aRet(3, 13, 20, 20) As String
    Dim aJyoCdBk As String
    Dim aJyoCd As String
    Dim aRaceNum As String
    Dim aJyoCdIdx As Integer
    Dim aRaceIdx As Integer
    Dim aOneIdx As Integer
    Dim aTwoIdx As Integer
    Dim aCmpiRnk(20) As String
    Dim aCmpiVal(20) As String
    Dim aFirstFlg As Boolean
    Dim aUmaOne As Integer
    Dim aUmaTwo As Integer
    Dim aCmpiPtn As String
    Dim aLookNinki As String
    
    aFirstFlg = True
    
    gYear = Format$(Now, "yyyy")
    gMonthDay = Format$(Now, "mmdd")
    
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "uma_RACE "
    gstrSql = gstrSql + "where "
    gstrSql = gstrSql + "Year ='" & gYear & "' and "
    gstrSql = gstrSql + "Monthday ='" & gMonthDay & "' and "
    gstrSql = gstrSql + "CmpiNinki <>'' "
    gstrSql = gstrSql + "order by JyoCD, racenum, umaban"
    
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
    Do
        If Rs.EOF = False Then
            umaban = Rs("umaban")
            
            If aFirstFlg <> True Then
                If umaban = "01" Then
                    'データ収集実行 _01-64_02-62
                    For aUmaOne = 1 To 20
                        If aCmpiVal(aUmaOne) <> "" Then
                            For aUmaTwo = 1 To 20
                                If aUmaOne <> aUmaTwo Then
                                    If aCmpiVal(aUmaTwo) <> "" Then
                                        If aCmpiRnk(aUmaOne) < aCmpiRnk(aUmaTwo) Then
                                            aCmpiPtn = "_" & aCmpiRnk(aUmaOne) & "-" & aCmpiVal(aUmaOne) & "_" & aCmpiRnk(aUmaTwo) & "-" & aCmpiVal(aUmaTwo)
                                            aLookNinki = aCmpiRnk(aUmaTwo)
                                        
                                            gstrSql = ""
                                            gstrSql = gstrSql + "SELECT "
                                            gstrSql = gstrSql + "* "
                                            gstrSql = gstrSql + "FROM "
                                            gstrSql = gstrSql + "analCmpi "
                                            gstrSql = gstrSql + "where "
                                            gstrSql = gstrSql + "CmpiPtn ='" & aCmpiPtn & "' and "
                                            gstrSql = gstrSql + "LookNinki ='" & aLookNinki & "' "
                                            
                                            Set Rs2 = db.OpenRecordset(gstrSql, dbOpenDynaset)
                                            
                                            If Rs2.EOF = False Then
                                                aRace(aJyoCdIdx, CInt(aRaceNum), aUmaOne, aUmaTwo) = Rs2("Rcnt")
                                                aHit(aJyoCdIdx, CInt(aRaceNum), aUmaOne, aUmaTwo) = Rs2("HtRate")
                                                aRet(aJyoCdIdx, CInt(aRaceNum), aUmaOne, aUmaTwo) = Rs2("RtRate")
                                            End If
                                            
                                            Rs2.Close
                                        Else
                                            aRace(aJyoCdIdx, CInt(aRaceNum), aUmaOne, aUmaTwo) = 0
                                            aHit(aJyoCdIdx, CInt(aRaceNum), aUmaOne, aUmaTwo) = 0
                                            aRet(aJyoCdIdx, CInt(aRaceNum), aUmaOne, aUmaTwo) = 0
                                        End If
                                        
                                    End If
                                Else
                                    aRace(aJyoCdIdx, CInt(aRaceNum), aUmaOne, aUmaTwo) = 0
                                    aHit(aJyoCdIdx, CInt(aRaceNum), aUmaOne, aUmaTwo) = 0
                                    aRet(aJyoCdIdx, CInt(aRaceNum), aUmaOne, aUmaTwo) = 0
                                End If
                            Next aUmaTwo
                        End If
                        
                    Next aUmaOne
                    
                    'データクリア
                    For aUmaTwo = 1 To 20
                        aCmpiRnk(aUmaTwo) = ""
                        aCmpiVal(aUmaTwo) = ""
                    Next aUmaTwo
                End If
            Else
                '初回のみ、判定スルー
                aFirstFlg = False
            End If
            
            aJyoCd = Rs("JyoCd")
            If aJyoCdBk <> "" And aJyoCdBk <> aJyoCd Then
                aJyoCdIdx = aJyoCdIdx + 1
            End If
            aJyoCdBk = aJyoCd
            
            aRaceNum = Rs("RaceNum")
            
            aCmpiRnk(CInt(umaban)) = Rs("cmpininki")
            aRace(aJyoCdIdx, CInt(aRaceNum), CInt(umaban), 0) = Rs("cmpininki")
            aCmpiVal(CInt(umaban)) = Rs("CmpiValue")
            aHit(aJyoCdIdx, CInt(aRaceNum), CInt(umaban), 0) = Rs("CmpiValue")
            
        Else
            Exit Do
        End If
        
        Rs.MoveNext
    
    Loop
    
    Rs.Close
    
    Dim fileNum1 As Long
    Dim fileNum2 As Long
    Dim fileNum3 As Long
    Dim fName As String
    Dim aDat1 As String
    
    
    For aJyoCdIdx = 0 To 3
        For aRaceIdx = 1 To 12
            '出力
            fileNum1 = FreeFile
            fName = Format$(aJyoCdIdx, "00") & "_" & Format$(aRaceIdx, "00") & "_cnt.csv"
            Open fName For Output As #fileNum1
            fileNum2 = FreeFile
            fName = Format$(aJyoCdIdx, "00") & "_" & Format$(aRaceIdx, "00") & "_hit.csv"
            Open fName For Output As #fileNum2
            fileNum3 = FreeFile
            fName = Format$(aJyoCdIdx, "00") & "_" & Format$(aRaceIdx, "00") & "_ret.csv"
            Open fName For Output As #fileNum3
            
            
            aDat1 = ",1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16"
            Print #fileNum1, aDat1
            Print #fileNum2, aDat1
            Print #fileNum3, aDat1
            
            aDat1 = ""
            aDat2 = ""
            For aUmaOne = 1 To 19
                aDat1 = aDat1 & "," & aRace(aJyoCdIdx, aRaceIdx, aUmaOne, 0)
                aDat2 = aDat2 & "," & aHit(aJyoCdIdx, aRaceIdx, aUmaOne, 0)
            Next aUmaOne
            Print #fileNum1, aDat1
            Print #fileNum1, aDat2
            Print #fileNum2, aDat1
            Print #fileNum2, aDat2
            Print #fileNum3, aDat1
            Print #fileNum3, aDat2
            
            For aUmaOne = 1 To 19
                aDat1 = ""
                aDat2 = ""
                aDat3 = ""
                
                For aUmaTwo = 1 To 19
                    aDat1 = aDat1 & "," & aRace(aJyoCdIdx, aRaceIdx, aUmaOne, aUmaTwo)
                    aDat2 = aDat2 & "," & aHit(aJyoCdIdx, aRaceIdx, aUmaOne, aUmaTwo)
                    aDat3 = aDat3 & "," & aRet(aJyoCdIdx, aRaceIdx, aUmaOne, aUmaTwo)
                Next aUmaTwo
                
                Print #fileNum1, aDat1
                Print #fileNum2, aDat2
                Print #fileNum3, aDat3
            
            Next aUmaOne
            
            Close #fileNum1
            Close #fileNum2
            Close #fileNum3
            
        Next aRaceIdx
    Next aJyoCdIdx
    
    MsgBox "finish"
    
End Sub

Private Sub Command31_Click()
    Dim key() As Byte
    Dim iv() As Byte
    Dim data() As Byte
    Dim objCipher As Cipher
    Dim retDat As String
    Dim str As String
    
    key = StringUtility.stringToByte("27842784midoniko")
    iv = StringUtility.stringToByte("midoniko27842784")
    
    Dim inD As String
    

    On Error GoTo ErrorHandler
    Set objCipher = New Cipher

'    Call objCipher.encrypt(key, iv, data)
'    retdata = Base64.encode(data)
    
    Dim aRcnt As String
    Dim aRtRate As String
    Dim aHtRate As String

    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "analCmpi "
    
    'DB input
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
    Do
        If Rs.EOF = False Then
            aLookNinki = Rs("LookNinki")
            aCmpiPtn = Rs("CmpiPtn")
            aRcnt = Rs("Rcnt")
            aRtRate = Rs("RtRate")
            aHtRate = Rs("HtRate")
            
            data = StringUtility.stringToByte(aCmpiPtn)
            str = data
            
            'DB output
            
            'データベースから削除
            gstrSql = ""
            gstrSql = gstrSql + "delete from analCmpi "
            gstrSql = gstrSql + "where"
            gstrSql = gstrSql + "'CmpiPtn=" & aCmpiPtn & "'"
            
            db.Execute gstrSql, dbFailOnError
            
            'データベースに追加
            gstrSql = ""
            gstrSql = gstrSql + "insert into analCmpi (Rcnt, RtRate, HtRate, CmpiPtn, LookNinki"
            gstrSql = gstrSql + ") values ("
            gstrSql = gstrSql + "'" + aRcnt & "', "
            gstrSql = gstrSql + "'" + aRtRate & "', "
            gstrSql = gstrSql + "'" + aHtRate & "', "
            gstrSql = gstrSql + "'" + str & "', "
            gstrSql = gstrSql + "'" + aLookNinki & "')"
            
            db.Execute gstrSql, dbFailOnError

            DoEvents
        Else
            Exit Do
        End If
        
        Rs.MoveNext
    Loop
    
    Rs.Close
    
    
    
    Call objCipher.decrypt(key, iv, data)
    
    Debug.Print StringUtility.byteToString(data)

    Exit Sub

ErrorHandler:
    Dim message As String

    message = "エラーコード: &H" & Hex(err.Number) & vbCrLf & _
        "ソース: " & err.Source & vbCrLf & err.Description
        MsgBox message, vbCritical

End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Dim i As Integer
    Dim src As String
    Dim file As String
    Dim wfile As String
    
    txtAll.Text = App.Path
    
'    file = txtAll.Text & "\cmpi" & Format$(Now, "yyyymmddhhnnss") & ".txt" 'txtAll.Text
'    wfile = txtAll.Text & "\cmpW" & Format$(Now, "yyyymmddhhnnss") & ".txt" 'txtAll.Text
    file = txtAll.Text & "\" & Format$(Now, "yyyymmddhhnnss") & ".txt" 'txtAll.Text
    wfile = txtAll.Text & "\" & Format$(Now, "yyyymmddhhnnss") & ".txt" 'txtAll.Text
    
    For i = 0 To List1.ListCount - 1
        src = List1.List(i)
        Call TextCodeChg(src)
'        Call msHTML2Txt(src, file)
        If chkNankan.value = 0 Then
            Call Compi2Txt(src & ".txt", file, wfile)
        Else
            Call Compi2TxtNankan(src & ".txt", file, wfile)
        End If
    Next i
    
    Dim ii As Integer
    Dim fcnt As Integer
    Dim toPath As String
    Dim Fso     As New FileSystemObject
    Dim FsoFile As file
    Dim spFile() As String
    
    toPath = App.Path & "\cmpiSel\"
    spFile = Split(file, "\")
    Set FsoFile = Fso.GetFile(file)
    FsoFile.Move (toPath & spFile(UBound(spFile)))
    
    For ii = 0 To 5
        If gFilename(ii) <> "" Then
            Set FsoFile = Fso.GetFile(gFilename(ii))
            FsoFile.Delete
            Set FsoFile = Fso.GetFile(gFilename(ii) & ".txt")
            FsoFile.Delete
        End If
    Next ii
    
    Set FsoFile = Nothing
    
    Command4.Enabled = True
    MsgBox "job finish!"
End Sub

Private Sub Compi2Txt(src As String, file As String, wfile As String)

Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
Dim strResult As String '置換後の文字列
Dim Matches
Dim Match
Dim fnTfr As Integer
Dim fn As Integer
Dim wfn As Integer
Dim lCnt As Integer
Dim data() As String
Dim wk As String
Dim wk2 As String
Dim wk3 As String
Dim wkPrt As String
Dim pos As Long
Dim phase As Long
Dim raceNo As String
Dim wkRaceNo As String
Dim retstr As String '
Dim nen As String
Dim gatu As String
Dim niti As String
Dim basho As String
Dim cmpininki As String
Dim cmpidata(12, 20) As String      'ninki,value
Dim cmpiTfr(12, 20) As String      'ninki,value
Dim umaban As Integer
Dim value As String
Dim idx As Integer
Dim backup As String
Dim wakCnt As Integer
Dim wakD() As String
Dim plc As Integer
Dim kire As String
Dim smpl As String
Dim wban As String
Dim cmpV As String


Dim wkwk As String
'    If optMode(0).value = True Then
         wkwk = "馬番"
'    Else
'         wkwk = "wakuNum"
'    End If


'HTMLファイル(param.)をメモリーに展開
'<<ファイル 開>>
fn = FreeFile
Open src For Input As #fn

'<<ファイル 読>>
lCnt = 0
Do Until EOF(fn)
    Line Input #fn, wk
    ReDim Preserve data(lCnt)
    data(lCnt) = wk
    lCnt = lCnt + 1
Loop

'<<ファイル 閉>>
Close #fn

'wfn = FreeFile
'Open wfile For Append As #wfn


'<<データ解析>>
'正規表現オブジェクトの宣言
Set objRegExp = New RegExp

With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    
    phase = 0
    For lCnt = 0 To UBound(data)
        
        Select Case phase
        '開催場所、年月日を検索
        '<TH BGCOLOR="#F56403" COLSPAN=31><FONT SIZE=+2>馬番コンピ　　　　</FONT><FONT SIZE=+2>2008年1月19日 1回中山5日目</FONT><FONT SIZE=+2>　　　　枠番コンピ</FONT></TH>
        '<h2 id="contentTit">2012年1月5日　コンピ指数－1回中山1日目</h2>
        Case 0
'             .Pattern = "<FONT SIZE=\+2>20.+日目"
             .Pattern = "<h2 id=""contentTit"">.+日目"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
               pos = Match.FirstIndex       '一致する文字列が見つかった位置
               retstr = Match.value
            Next
            If pos = 0 Then
'                .Pattern = "<font size=""\+2"">20.+日目"
                .Pattern = "contentTit""\>20.+日目"
                Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                '<FONT SIZE=+2>2008年1月20日 1回中山6日目
                '年
                 .Pattern = ">.+年"
                pos = 0
                Set Matches = .Execute(retstr)   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   wk = Match.value
                Next
                nen = Mid$(wk, 2, 4)
                '月
                 .Pattern = "年.+月"
                pos = 0
                Set Matches = .Execute(retstr)   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   wk = Match.value
                Next
                If Len(wk) = 3 Then
                    gatu = Mid$(wk, 2, 1)
                Else
                    gatu = Mid$(wk, 2, 2)
                End If
                '日
                 .Pattern = "月.+日　コンピ指数"
                pos = 0
                Set Matches = .Execute(retstr)   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   wk = Match.value
                Next
                If Len(wk) = 9 Then
                    niti = Mid$(wk, 2, 1)
                Else
                    niti = Mid$(wk, 2, 2)
                End If
                '開催場所
                 .Pattern = "回.+日目"
                pos = 0
                Set Matches = .Execute(retstr)   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   wk = Match.value
                Next
                Select Case Mid$(wk, 2, 2)
                Case "札幌"
                    basho = "01"
                Case "函館"
                    basho = "02"
                Case "福島"
                    basho = "03"
                Case "新潟"
                    basho = "04"
                Case "東京"
                    basho = "05"
                Case "中山"
                    basho = "06"
                Case "中京"
                    basho = "07"
                Case "京都"
                    basho = "08"
                Case "阪神"
                    basho = "09"
                Case "小倉"
                    basho = "10"
                End Select
                
                phase = 1
            End If
        'レース番号を検索
        '<TD NOWRAP> １Ｒ<BR>サラ３歳</TD>
        '<td class="racename"><span class="race">12R</span>
        Case 1
'            .Pattern = ">.+Ｒ<BR>"
            .Pattern = ">.+R\<\/span\>"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
               pos = Match.FirstIndex       '一致する文字列が見つかった位置
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = ">.+Ｒ<br>"
                Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                If raceNo <> Mid$(retstr, 21, 2) Then
                    wkRaceNo = Mid$(retstr, 21, 2)
                    If Right$(wkRaceNo, 1) = "R" Then
                        wkRaceNo = StrConv(Left$(wkRaceNo, 1), vbNarrow)
                    End If
                    If wkRaceNo <> raceNo And IsNumeric(wkRaceNo) Then
                        raceNo = wkRaceNo
                        phase = 2
                    End If
                End If
            End If
        'コンピ指数データ直前を検索
        '<TD NOWRAP>馬番<BR>指数</TD>
        Case 2
'             .Pattern = "馬番<BR>指数"
'            If optMode(0).value = True Then
                .Pattern = "馬番" & "\<br\>指数" '0504"\<br \/\>指数"
'            Else
'                .Pattern = "wakuNum.+\>8"
'            End If
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
               pos = Match.FirstIndex       '一致する文字列が見つかった位置
               retstr = Match.value
            Next
            If pos = 0 Then
'                If optMode(0).value = True Then
                    .Pattern = wkwk & "<br>指数"
'                Else
'                    .Pattern = "wakuNum.+\>8"
'                End If
                Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                phase = 3
                cmpininki = 0
            End If
        
        'コンピ指数データを検索
        '<TD NOWRAP>９<BR>68</TD>
        Case 3
'             .Pattern = "NOWRAP>.+<"
'            If optMode(0).value = True Then         '0524
                .Pattern = "\>.+\<br\>.+\<\/td\>"
'            Else
'
'                ReDim wakD(0)
'
'                Do
'                    'wakuren以降を</td>で区切る
'                    plc = InStr(data(lCnt), "</td>")
'                    If plc > 0 Then
'                        smpl = data(lCnt)
'                        Do
'                            ReDim Preserve wakD(UBound(wakD) + 1)
'                            kire = Left$(smpl, plc + 4)
'                            wakD(UBound(wakD)) = kire
'                            kire = Mid$(smpl, plc + 5)
'                            plc = InStr(kire, "</td>")
'                            If plc > 0 Then
'                                smpl = kire
'                            Else
'                                Exit Do
'                            End If
'                        Loop
'                    Else
'                        '不要データ
'                    End If
'
'                    If lCnt = UBound(data) Then
'                        Exit Do
'                    End If
'
'                    lCnt = lCnt + 1
'                Loop
'
'                lCnt = 0
'                wk = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho
'                wkPrt = wk
'                Do
'                    If UBound(wakD) < lCnt Then
''                        Print #wfn, wkPrt
'                        Exit Do
'                    End If
'                    If wakD(lCnt) <> "" Then
'                        retstr = ""
'                        'レース番号？
'                        .Pattern = "race""\>.+R\<\/span\>"
'                        pos = 0
'                        Set Matches = .Execute(wakD(lCnt))   ' 検索を実行します。
'                        For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
'                           pos = Match.FirstIndex       '一致する文字列が見つかった位置
'                           retstr = Match.value
'                        Next
'
'                        If retstr <> "" Then
'                            'レース番号抽出
'                            kire = Mid$(retstr, 7, 2)
'                            If Right$(kire, 1) = "R" Then
'                                raceNo = "0" & Left$(kire, 1)
'                            Else
'                                raceNo = kire
'                            End If
'
'                            If wkPrt <> wk Then
''                                Print #wfn, wkPrt
'                                wkPrt = wk
'                            End If
'
'                            wkPrt = wkPrt & "," & raceNo
'                        Else
'                            '枠番抽出
'                            .Pattern = "\>.+\<br \/\>"
'                            pos = 0
'                            Set Matches = .Execute(wakD(lCnt))   ' 検索を実行します。
'                            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
'                               pos = Match.FirstIndex       '一致する文字列が見つかった位置
'                               retstr = Match.value
'                            Next
'
'                            wban = Mid$(retstr, 2, 2)
'                            If wban <> "枠番" And wban <> "" Then
'
'                                wkPrt = wkPrt & "," & wban
'
'                                'コンピ指数抽出
'                                .Pattern = "\>.+?\<\/td\>"
'                                pos = 0
'                                Set Matches = .Execute(wakD(lCnt))   ' 検索を実行します。
'                                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
'                                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
'                                   retstr = Match.value
'                                Next
'
'                                cmpV = Left$(Right$(retstr, 7), 2)
'                                If cmpV = "取消" Then
'                                    cmpV = "00"
'                                End If
'                                wkPrt = wkPrt & "," & cmpV
'                            End If
'                        End If
'                    End If
'
'                    lCnt = lCnt + 1
'                Loop
'
''                Close #wfn
'
'                cmdCmpi.Enabled = True
'
'                Exit Sub
'            End If
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
               pos = Match.FirstIndex       '一致する文字列が見つかった位置
               retstr = Match.value
            Next
            If pos = 0 Then
'                If optMode(0).value = True Then
                    .Pattern = "nowrap>.+<"
'                Else
'                    .Pattern = "枠番.+\>.+\<br \/\>.+\<\/td\>"
'                End If
                Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   retstr = Match.value
                Next
            End If
            
            
            If pos = 0 Then
                '終端チェック
                '<TD COLSPAN=2 NOWRAP>　</TD>
'                 .Pattern = "<TD COLSPAN=. NOWRAP>"
                 .Pattern = "　\<\/td\>"
                 
                pos = 0
                Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   retstr = Match.value
                Next
                If pos = 0 Then
                    .Pattern = "tr\>"
                    Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
                    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                       pos = Match.FirstIndex       '一致する文字列が見つかった位置
                       retstr = Match.value
                    Next
                End If
                
                If pos <> 0 Then
                    If raceNo = "12" Then
                        Exit For
                    Else
                        phase = 1
                    End If
                End If
            Else
                'data 取り込み
                'NOWRAP>11<BR>71<
                ' 6<br />86</td>
                 
                 'umaban
                '.Pattern = "NOWRAP>.+<BR>"
                If InStr(retstr, "枠番") = 0 Then
                    .Pattern = "\>.+\<br\>.+\<\/td\>"
                Else
                    .Pattern = "race""\>.+R\<"
                    pos = 0
                    Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
                    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                       pos = Match.FirstIndex       '一致する文字列が見つかった位置
                       retstr = Match.value
                    Next
                    
                    raceNo = Mid$(retstr, 7, 2)
                    If Right$(raceNo, 1) = "R" Then
                        raceNo = Left$(raceNo, 1)
                    End If
                    
                    .Pattern = "枠番.+\<br \/\>.+\<\/td\>"
                    pos = 0
                    Set Matches = .Execute(data(lCnt))   ' 検索を実行します。
                    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                       pos = Match.FirstIndex       '一致する文字列が見つかった位置
                       retstr = Match.value
                    Next
                End If
                
                pos = 0
                Set Matches = .Execute(retstr)   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                   pos = Match.FirstIndex       '一致する文字列が見つかった位置
                   wk = Match.value
                Next
                If pos = 0 Then
                    .Pattern = "nowrap\>.+\<br\>"
                    Set Matches = .Execute(retstr)   ' 検索を実行します。
                    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                       pos = Match.FirstIndex       '一致する文字列が見つかった位置
                       wk = Match.value
                    Next
                End If
                
                cmpininki = cmpininki + 1
                If InStr(retstr, "suishou") = 0 Then
                    If Mid$(wk, 9, 1) = "<" Then
                        wk = Mid$(wk, 2, 2)
                    Else
                        wk = Mid$(wk, 2, 2)
                    End If
                Else
                    wk = Mid$(wk, 36, 2)
                End If
                
'                If optMode(1).value = True Then
'                    wakCnt = wakCnt + 1
'                End If
                
                If IsNumeric(wk) = False Or wakCnt = 8 Then
                    wakCnt = 0
                    
                    If raceNo = "12" Then
                        Exit For
                    Else
'                        If optMode(1).value = True Then
'                            phase = 3
'                        Else
                            phase = 1
'                        End If
                    End If
                Else
                    umaban = wk
                    
'                    backup = cmpidata(CInt(raceNo), umaban)
                    
                     'value
                     .Pattern = "\<br\>.+\<"        '0524
                    pos = 0
                    
'                    If optMode(1).value = True Then
'                        retstr = Right$(retstr, 13)
'                    End If
                    
                    
                    Set Matches = .Execute(retstr)   ' 検索を実行します。
                    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                       pos = Match.FirstIndex       '一致する文字列が見つかった位置
                       wk = Match.value
                    Next
                    If pos = 0 Then
                        .Pattern = "<br>.+<"
                        Set Matches = .Execute(retstr)   ' 検索を実行します。
                        For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                           pos = Match.FirstIndex       '一致する文字列が見つかった位置
                           wk = Match.value
                        Next
                    End If
                    
                    If Mid$(wk, 6, 1) = "<" Then
                        value = Mid$(wk, 5, 2)
                    Else
                        value = Mid$(wk, 5, 2)
                    End If
                    If IsNumeric(value) = False And Left$(value, 1) <> "消" Then
                        If raceNo = "12" Then
                            Exit For
                        Else
                            phase = 1
                        End If
                    Else
                        '消が複数あるケースを対応 20170524
                        If Left$(value, 1) = "消" Then
                            value = "0"
                        End If
                        cmpidata(CInt(raceNo), umaban) = cmpininki
                        cmpidata(CInt(raceNo), umaban) = cmpidata(CInt(raceNo), umaban) & "," & value
                        cmpiTfr(CInt(raceNo), umaban) = value
                        
                        Do
                            If InStr(Mid$(retstr, InStr(retstr, "消") + 6), "消") = 0 Then
                                Exit Do
                            End If
                            
                            cmpininki = cmpininki + 1
                            retstr = Mid$(retstr, InStr(retstr, "消") + 6)
                            
                            .Pattern = "\<td\>.+\<"        '0524
                            
                            Set Matches = .Execute(retstr)   ' 検索を実行します。
                            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                               pos = Match.FirstIndex       '一致する文字列が見つかった位置
                               wk = Match.value
                            Next
                            
                            If Mid$(wk, 6, 1) = "<" Then
                                value = Mid$(wk, 5, 2)
                            Else
                                value = Mid$(wk, 5, 2)
                            End If
                            
                            umaban = CInt(value)
                            cmpidata(CInt(raceNo), umaban) = cmpininki
                            cmpidata(CInt(raceNo), umaban) = cmpidata(CInt(raceNo), umaban) & "," & "0"
                            cmpiTfr(CInt(raceNo), umaban) = "0"
                        Loop
                        
                    End If
                    
                    retstr = "" 'retstr
                End If
            End If
        
        End Select
        
    Next lCnt
    
End With

Set objRegExp = Nothing

'テキストファイル(param.)へ出力
src = file
fn = FreeFile
Open src For Append As #fn

'<<ファイル 書>>

wk = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho

Dim tfr As String


    For idx = 1 To 12
        
        tfr = txtAll.Text & "\" & wk & Format$(idx, "00") & ".csv"
        fnTfr = FreeFile
If chkTF.value = 1 Then
        Open tfr For Output As #fnTfr
End If
        
        wk2 = ""
        For lCnt = 1 To 20
            wk2 = wk2 & "," & cmpidata(idx, lCnt)
            If cmpiTfr(idx, lCnt) <> "" Then
    '            If wk3 = "" Then
    '                wk3 = cmpiTfr(idx, lCnt)
    '            Else
                    wk3 = wk3 & "," & cmpiTfr(idx, lCnt)
    '            End If
            End If
        Next lCnt
        
        wk2 = wk & Format$(idx, "00") & wk2
        Print #fn, wk2
If chkTF.value = 1 Then
        Print #fnTfr, wk & Format$(idx, "00") & wk3
        
        Close #fnTfr
End If
    
    Next idx

'<<ファイル 閉>>
Close #fn
'Close #wfn


End Sub
Private Sub TextCodeChg(pSrc As String)
    ' テキストをバイト配列で読込
    Dim ipath As String: ipath = pSrc   'App.Path & "\TestUtf8.txt"
    Dim idat() As Byte
    ReDim idat(FileLen(ipath) - 1) As Byte
    Dim intFileNo As Integer
    intFileNo = FreeFile
    Open ipath For Binary As intFileNo
    Get intFileNo, , idat
    Close intFileNo
            
    ' 文字コード判定(blnBin=バイナリ判定無し)
    Dim cod As String: cod = objNonCode.GetCodeName(idat, blnBin:=False)

    ' 判定した文字コードをString(UNICODE)に変換
    Dim uni As String
    Select Case cod
        Case "SJIS"
            ' SJISからUNICODEへの変換
            uni = objNonCode.SJIS_To_VbUnicode(idat)
        Case "JIS"
            ' JISからUNICODEへの変換
            uni = objNonCode.JIS_To_VbUnicode(idat)
        Case "EUC"
            ' EUCからUNICODEへの変換
            uni = objNonCode.EUC_To_VbUnicode(idat)
        Case "UNICODE"
            ' UNICODEからUNICODEへの変換
            uni = objNonCode.UNICODE_To_VbUnicode(idat)
        Case "UTF7"
            ' UTF-7からUNICODEへの変換
            uni = objNonCode.UTF7_To_VbUnicode(idat)
        Case "UTF8"
            ' UTF-8からUNICODEへの変換
            uni = objNonCode.UTF8_To_VbUnicode(idat)
        Case "BIN"
            ' SJISからUNICODEへの変換
            uni = objNonCode.SJIS_To_VbUnicode(idat)
        Case Else
            ' SJISからUNICODEへの変換
            uni = objNonCode.SJIS_To_VbUnicode(idat)
    End Select

    ' 読込ファイルの改行コードをCRLFへ変換
    uni = objNonCode.ChangeReturnToCrLf(uni)
    
    ' String(UNICODE)を出力したい文字コードのByte配列に変換
    Dim odat() As Byte
    cod = strOutCode
    Select Case cod
        Case "SJIS"
            ' UNICODEからSJISへの変換
            odat = objNonCode.VbUnicode_To_SJIS(uni)
        Case "JIS"
            ' UNICODEからJISへの変換
            odat = objNonCode.VbUnicode_To_JIS(uni)
        Case "EUC"
            ' UNICODEからEUCへの変換
            odat = objNonCode.VbUnicode_To_EUC(uni)
        Case "UNICODE"
            ' UNICODEからUNICODEへの変換
            odat = objNonCode.VbUnicode_To_UNICODE(uni)
        Case "UTF7"
            ' UNICODEからUTF7への変換
            odat = objNonCode.VbUnicode_To_UTF7(uni)
        Case "UTF8"
            ' UNICODEからUTF8への変換
            odat = objNonCode.VbUnicode_To_UTF8(uni)
        Case Else
            ' UNICODEからSJISへの変換
            odat = objNonCode.VbUnicode_To_SJIS(uni)
    End Select

    ' 出力ファイルをバイナリ形式で出力
    Dim opath As String: opath = pSrc & ".txt" 'App.Path & "\TestOut.txt"
    If Len(Dir(opath)) <> 0 Then
        Kill opath
    End If
    intFileNo = FreeFile
    Open opath For Binary As intFileNo
    Put intFileNo, , odat
    Close intFileNo
End Sub


Private Sub Command5_Click()
    Command5.Enabled = False
    'ファイルリスト作成
    ' FileSystemObject (FSO) の新しいインスタンスを生成する
    Dim cFso As FileSystemObject
    Set cFso = New FileSystemObject

    ' Folder オブジェクトを取得する
    Dim cFolder As Folder
    Set cFolder = cFso.GetFolder(App.Path & "\cmpiSel\")

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFso = Nothing

    Dim stPrompt As String
    Dim cFile    As file

    ' すべてのファイルを列挙する
    For Each cFile In cFolder.Files
        stPrompt = stPrompt & cFile.Path & ","
    Next cFile

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFolder = Nothing
    Set cFile = Nothing
    
    Files = Split(stPrompt, ",")
    
    Dim jj As Long
    Dim kk As Long
    Dim umaban As String
    Dim aCmpiNinki As String
    Dim aCmpiValue As String
    Dim aR() As String
    
    raceCnt = 0
    HaraiCnt = 0
    
    For jj = 0 To UBound(Files) - 1
        
        
        fDat = aYear & aMonthday & aJyoCd & aRaceNum
        
        'コンピファイルを読み込み
        fnum = FreeFile()
        
        Open Files(jj) For Input As #fnum
        
        Do Until EOF(fnum)
            Line Input #fnum, wk
            aR = Split(wk, ",")
            '200701060801,3,66,12,44,16,40,10,47,8,51,5,58,13,43,6,53,4,60,1,78,7,52,2,70,14,42,15,41,11,45,9,49,,,,
            aYear = Left$(aR(0), 4)
            aMonthday = Mid$(aR(0), 5, 4)
            aJyoCd = Mid$(aR(0), 9, 2)
            aRaceNum = Mid$(aR(0), 11, 2)
            
            '人気、コンピ指数のセットで並ぶ。馬番順
            For kk = 1 To (UBound(aR) / 2) - 1
                If aR(kk) = "" Then
                    Exit For
                End If
                
                umaban = Format$(kk, "00")
                
                gstrSql = ""
                gstrSql = gstrSql + "SELECT "
                gstrSql = gstrSql + "* "
                gstrSql = gstrSql + "FROM "
                gstrSql = gstrSql + "uma_RACE "
                gstrSql = gstrSql + "where "
                gstrSql = gstrSql + "Year ='" & aYear & "' and "
                gstrSql = gstrSql + "Monthday ='" & aMonthday & "' and "
                gstrSql = gstrSql + "JyoCD ='" & aJyoCd & "' and "
                gstrSql = gstrSql + "racenum ='" & aRaceNum & "' and "
                gstrSql = gstrSql + "umaban ='" & umaban & "'"
                
                Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
                
                If Rs.EOF = True Then
                    'データベースに追加
                    aCmpiNinki = Format$(aR(kk * 2 - 1), "00")
                    aCmpiValue = Format$(aR(kk * 2), "00")
                    gstrSql = ""
                    gstrSql = gstrSql + "insert into uma_race (Year, monthday, jyocd, racenum, umaban, CmpiNinki, CmpiValue"
                    gstrSql = gstrSql + ") values ("
                    
                    gstrSql = gstrSql + "'" & aYear & "', "
                    gstrSql = gstrSql + "'" & aMonthday & "', "
                    gstrSql = gstrSql + "'" & aJyoCd & "', "
                    gstrSql = gstrSql + "'" & aRaceNum & "', "
                    gstrSql = gstrSql + "'" & umaban & "', "
                    gstrSql = gstrSql + "'" & aCmpiNinki & "', "
                    gstrSql = gstrSql + "'" & aCmpiValue & "')"
                    
                    db.Execute gstrSql, dbFailOnError
                Else
                End If
                
                Rs.Close
                
            Next kk
            
        Loop
        
        Close #fnum
    
    Next jj
    
    Command5.Enabled = True
    
    MsgBox "job finish!"

End Sub

Private Sub Command6_Click()
    '20070106京都.txt
    
    'ファイルがあれば、データベース更新ミス？
    
End Sub

Private Sub Command7_Click()
    Dim ii As Long
    
    picpic.Cls
    
    For ii = 0 To 365 * 10
        Call sDot(ii * 4, 1000)
'        Call sDot(ii * 4, 1100)
    Next ii

End Sub

Private Sub Command8_Click()
    Dim aTosu
    Dim aLvl(18) As Integer
    Dim aKumi As String
    Dim aAllKumi() As String
    Dim aCurLvl As Integer
    
    Debug.Print Now
    
    ReDim aAllKumi(0)
    
    For aTosu = 1 To 18
        
        For aLvl01 = 90 To (40 - 17) Step -1
            aKumi = ""
            'aKumi = "01-" & aLvl01
            If aTosu > 1 Then
                For aLvl02 = (aLvl01 + 1) To (40 - 16) Step -1
                    'aKumi = aKumi & "_02-" & aLvl02
                    
                    If aTosu > 2 Then
                        For aLvl03 = (aLvl02 + 1) To (40 - 15) Step -1
                            'aKumi = aKumi & "_03-" & aLvl03
                    
                            If aTosu > 3 Then
                                For aLvl04 = (aLvl03 + 1) To (40 - 14) Step -1
                                    'aKumi = aKumi & "_04-" & aLvl04
                    
                                    If aTosu > 4 Then
                                        For aLvl05 = (aLvl04 + 1) To (40 - 13) Step -1
                                            'aKumi = aKumi & "_05-" & aLvl05
                    
                                            If aTosu > 5 Then
                                                For aLvl06 = (aLvl05 + 1) To (40 - 12) Step -1
                                                    'aKumi = aKumi & "_06-" & aLvl06
                                                    
                                                    If aTosu > 6 Then
                                                        For aLvl07 = (aLvl06 + 1) To (40 - 11) Step -1
                                                            'aKumi = aKumi & "_07-" & aLvl07
                                                        
                                                            If aTosu > 7 Then
                                                                For aLvl08 = (aLvl07 + 1) To (40 - 10) Step -1
                                                                    'aKumi = aKumi & "_08-" & aLvl08
                                                                    
                                                                    If aTosu > 8 Then
                                                                        For aLvl09 = (aLvl08 + 1) To (40 - 9) Step -1
                                                                            'aKumi = aKumi & "_09-" & aLvl09
                                                                
                                                                            If aTosu > 9 Then
                                                                                For aLvl10 = (aLvl09 + 1) To (40 - 8) Step -1
                                                                                    'aKumi = aKumi & "_10-" & aLvl10
                                                                                    
                                                                                    If aTosu > 10 Then
                                                                                        For aLvl11 = (aLvl10 + 1) To (40 - 7) Step -1
                                                                                            'aKumi = aKumi & "_11-" & aLvl11
                                                                                            
                                                                                            If aTosu > 11 Then
                                                                                                For aLvl12 = (aLvl11 + 1) To (40 - 6) Step -1
                                                                                                    'aKumi = aKumi & "_12-" & aLvl12
                                                                                            
                                                                                                    If aTosu > 12 Then
                                                                                                        For aLvl13 = (aLvl12 + 1) To (40 - 5) Step -1
                                                                                                            'aKumi = aKumi & "_13-" & aLvl13
                                                                                                            
                                                                                                            If aTosu > 13 Then
                                                                                                                For aLvl14 = (aLvl13 + 1) To (40 - 4) Step -1
                                                                                                                    'aKumi = aKumi & "_14-" & aLvl14
                                                                                                                    If aTosu > 14 Then
                                                                                                                    
                                                                                                                        For aLvl15 = (aLvl14 + 1) To (40 - 3) Step -1
                                                                                                                            'aKumi = aKumi & "_15-" & aLvl15
                                                                                                                            
                                                                                                                            If aTosu > 15 Then
                                                                                                                                For aLvl16 = (aLvl15 + 1) To (40 - 2) Step -1
                                                                                                                                    'aKumi = aKumi & "_16-" & aLvl16
                                                                                                                                    
                                                                                                                                    If aTosu > 16 Then
                                                                                                                                        For aLvl17 = (aLvl16 + 1) To (40 - 1) Step -1
                                                                                                                                            'aKumi = aKumi & "_17-" & aLvl17
                                                                                                                                            
                                                                                                                                            If aTosu > 17 Then
                                                                                                                                                For aLvl18 = (aLvl17 + 1) To (40) Step -1
                                                                                                                                                    'aKumi = aKumi & "_18-" & aLvl18
                                                                                                                                                    
                                                                                                                                                Next aLvl18
                                                                                                                                            End If
                                                                                                                                        Next aLvl17
                                                                                                                                    End If
                                                                                                                                Next aLvl16
                                                                                                                            End If
                                                                                                                        Next aLvl15
                                                                                                                    End If
                                                                                                                Next aLvl14
                                                                                                            End If
                                                                                                        Next aLvl13
                                                                                                    End If
                                                                                                Next aLvl12
                                                                                            End If
                                                                                        Next aLvl11
                                                                                    End If
                                                                                Next aLvl10
                                                                            End If
                                                                        Next aLvl09
                                                                    End If
                                                                Next aLvl08
                                                            End If
                                                        Next aLvl07
                                                    End If
                                                Next aLvl06
                                            End If
                                        Next aLvl05
                                    End If
                                Next aLvl04
                            End If
                        Next aLvl03
                    End If
                Next aLvl02
            End If
        Next aLvl01
    Next aTosu
    
    Debug.Print Now
    
    MsgBox "finish"
End Sub

Private Sub Command9_Click()
    wk = (Combination(18, 16))
    fnTfr = FreeFile
    Open App.Path & "\test.txt" For Output As #fnTfr
    Print #fnTfr, wk
        
    Close #fnTfr
End Sub
Function Combination(Max As Integer, Count As Integer, Optional Min As Integer = 1, Optional Preceding As String = "") As String
     Dim i As Integer, Buffer As String
     If Max - Min + 1 < Count Then
          ' 「4～5 の中の 7個の組み合わせ」 みたいにアリエナイ場合の処理ね
          Let Combination = "Error"
     ElseIf Count = 0 Then
'          「0 個の組み合わせ」 の場合ね
          Let Combination = Preceding & vbCrLf
     ElseIf Max - Min + 1 = Count Then
          ' 「4～5 の中の 2個の組み合わせ」 みたいな場合の処理だよぉ
          ' 残りの数字の個数が 指定の個数 と同じだった場合ね
          Let Buffer = Preceding
          For i = Min To Max
               Let Buffer = Buffer & " " & i
          Next i
          Let Combination = Buffer & vbCrLf
     Else
          Let Buffer = ""
          For i = Min To Max - (Count - 1)
               Let Buffer = Buffer & Combination(Max, Count - 1, i + 1, Preceding & " " & i)
          Next i
          Let Combination = Buffer
     End If
End Function

Private Sub Form_Load()
    Dim mail As String
    Dim aTitle As String
    Dim aBody As String
    Dim aMac As String
    
    List1.OLEDropMode = vbOLEDropManual
    
    areaMD.Text = Format$(Now, "mmdd")
    areaY.Text = Format$(Now, "yyyy")
    
'    aTitle = GC_APLI_NAME & GC_THANKS
'    aBody = GC_AMAZON
'
''    mail = sendMail(aTitle, aBody, GC_BLOG_MAIL)
'
'    aTitle = "コンピ取得ツール"
'    aMac = getMacAddress
''    mail = sendMail(aTitle, aMac, GC_MAC_MAIL)
'
'    If mail = "" Then
'        'OK
'    Else
'        MsgBox GC_FAIL_MAIL
'        End
'    End If a
    
    If CreateObject("NonCodeVb6.NonCodeClass") Is Nothing Then
        If Len(Dir("NonCodeVb6.dll")) <> 0 Then
            ' NonCodeVb6.dllのレジストリ登録
            Shell "regsvr32 /s NonCodeVb6.dll", vbHide
        Else
            ' NonCodeVb6.dllをCode2Code.exeと同じフォルダに置いてください。
            MsgBox _
            "NonCodeVb6.dllが見つかりませんでした。" & vbCrLf & vbCrLf & _
            "NonCodeVb6.dllを" & vbCrLf & "[" & App.Path & "]" & vbCrLf & _
            "に置いてください。"
            End
        End If
    End If
    On Error GoTo 0
    
    Set objNonCode = CreateObject("NonCodeVb6.NonCodeClass")
    
    Dim nowD As String
    nowD = Format$(Now, "yyyymmdd")
    If nowD > "20190101" Then                   '20170403
        MsgBox "最新バージョンをホームページからダウンロードしてください。"
        End
    End If
    
'    'パスワード
'    ret = InputBox("password", "okba")
'    If ret <> "okumieta" Then
'        MsgBox "wrong password"
'        End
'    End If
    
    
    'DB接続
    gRet = cnctDB
    
    
    mSel = 1
    
    mPos(0) = 1
    mPos(1) = 2
    mPos(2) = 3
    mPos(3) = 4
    mPos(4) = 5
    mPos(5) = 6
    mPos(6) = 7
    mPos(7) = 8
    mPos(8) = 9
    mPos(9) = 10
    mPos(10) = 11
    mPos(11) = 12
    mPos(12) = 13
    mPos(13) = 14
    mPos(14) = 15
    mPos(15) = 16
    mPos(16) = 17
    mPos(17) = 18
    
    Call sDraw

    For ii = 0 To 17
        opt(ii).BackColor = vbRed
    Next ii
'        opt(17).Left = 315 * 50
    
    For ii = 0 To 99
        lbl(ii).Top = 1710
        lbl(ii).Width = M_HABA
        lbl(ii).Left = ii * M_HABA
        lbl(ii).Caption = 90 - ii
        If (90 - ii) < 40 Then
            lbl(ii).Visible = False
        End If
    Next ii
    
    picpic.Line (0, 7000 - 500)-(picpic.Width, 7000 - 500), vbWhite
    For ii = 1 To 10
        picpic.Line (700 * ii, 7000)-(700 * ii, 7000 - 200), vbWhite
    Next ii
End Sub
Private Sub sDraw()
    Dim ii As Integer
    
    For ii = 0 To 99
        If mPos(ii) > 0 Then
            opt(mPos(ii) - 1).Width = M_HABA
            opt(mPos(ii) - 1).Left = ii * M_HABA
            opt(mPos(ii) - 1).Caption = mPos(ii)
        End If
    Next ii
    
End Sub



Private Sub List1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lstrTmp             As String
    Dim i As Integer
    
On Error GoTo ErrHandler
    
    'ﾄﾞﾛｯﾌﾟされたものが、ﾌｧｲﾙであるか判断
    If data.GetFormat(vbCFFiles) Then
        For i = 1 To data.Files.Count
            List1.AddItem (data.Files(i))
        Next i
        
    Else
        MsgBox "ドロップされたものがﾌｧｲﾙではありません。"
        Exit Sub
    End If
    
    Exit Sub
ErrHandler:
    MsgBox "error:" & err.Description
    Exit Sub


End Sub

Private Sub opt_Click(Index As Integer)
    opt(Index).value = True
    mSel = Index + 1
    
    If opt(Index).BackColor <> vbBlue Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub

Private Function cnctDB() As Long
    On Error GoTo err_handler
    
    Dim lstrDb              As String
    Dim llngRet             As Long
    
    gDB = PATH_DB
    llngRet = gfConnectDB(gDB)
    If llngRet <> 0 Then
        MsgBox "cnctDB エラー:" & llngRet
        Exit Function
    End If
    
    cnctDB = llngRet
    
    Exit Function

err_handler:
        MsgBox "cnctDB エラー:" & err.Description & vbCr & vbLf & "エラー番号:" & err.Number
End Function

Public Function gfConnectDB(pstrDb As String) As Long

' DAOのオブジェクト変数を宣言する
    
    ' デフォルトのワークスペースを定義する
    Set ws = DBEngine.Workspaces(0)
    ' データベースを開く
    Set db = ws.OpenDatabase(pstrDb, False, False, ";pwd=okutotta")

End Function

Private Function sendMail(msg_subject As String, msg_body As String, aite As String) As String
On Error GoTo err_handler

Set objMail = CreateObject("CDO.Message")

objMail.From = "o.k.keiba@gmail.com"
objMail.To = aite
objMail.Subject = msg_subject
objMail.HTMLBody = msg_body

Dim strConfigurationField  As String

strConfigurationField = "http://schemas.microsoft.com/cdo/configuration/"
With objMail.Configuration.Fields
   .Item(strConfigurationField & "sendusing") = 2
   .Item(strConfigurationField & "smtpserver") = "smtp.googlemail.com"
   .Item(strConfigurationField & "smtpserverport") = 465
   .Item(strConfigurationField & "smtpusessl") = True
   .Item(strConfigurationField & "smtpauthenticate") = 1
   .Item(strConfigurationField & "sendusername") = "o.k.keiba@gmail.com"
   .Item(strConfigurationField & "sendpassword") = "lets.keiba7"
   .Item(strConfigurationField & "smtpconnectiontimeout") = 60
   .Update
End With

objMail.send

Set objMail = Nothing

sendMail = ""

Exit Function

err_handler:
    
sendMail = err.Description

End Function

'中央
'4, 3, 66
'2, 3, 67
'10, 19, 68
'2, 2, 69
'9, 17, 69
'10, 17, 69
'8, 16, 70
'9, 20, 70
'10, 20, 70
'11, 23, 71
'8, 19, 72
'9, 21, 72
'13, 25, 73
'8, 21, 75
'9, 22, 75
'11, 26, 75
'12, 26, 75
'3, 3, 76
'6, 19, 76
'7, 21, 76
'9, 26, 76
'6, 20, 78
'10, 27, 78
'5, 18, 79
'9, 27, 79
'12, 31, 79
'6, 23, 80
'11, 30, 80
'5, 21, 81
'9, 31, 81
'10, 31, 81
'2, 3, 82
'5, 22, 82
'7, 26, 82
'5, 24, 83
'6, 26, 83
'2, 7, 84
'5, 25, 84
'9, 34, 85
'6, 29, 86
'8, 34, 86
'9, 35, 86
'11, 37, 86
'2, 14, 88
'6, 31, 88
'9, 37, 88
'10, 37, 88
'3, 23, 90

'南関
'2, 1, 67
'3, 3, 69
'7, 17, 69
'7, 17, 70
'9, 21, 70
'10, 21, 70
'11, 24, 70
'10, 23, 71
'9, 24, 72
'10, 24, 72
'3, 5, 73
'10, 24, 73
'12, 26, 73
'14, 32, 73
'3, 4, 74
'10, 26, 74
'4, 10, 75
'6, 18, 75
'7, 23, 75
'8, 24, 75
'11, 28, 75
'14, 33, 75
'4, 11, 76
'6, 19, 76
'4, 13, 77
'7, 24, 77
'13, 31, 77
'10, 30, 78
'13, 35, 78
'9, 31, 79
'5, 21, 80
'8, 29, 80
'9, 30, 80
'10, 31, 80
'4, 19, 81
'5, 23, 81
'8, 29, 81
'5, 23, 82
'6, 25, 82
'9, 32, 82
'11, 34, 82
'7, 30, 83
'8, 30, 83
'11, 35, 83
'10, 36, 84
'7, 31, 85
'8, 33, 85
'10, 37, 85
'4, 25, 86
'8, 34, 86
'14, 40, 86
'9, 37, 87
'6, 34, 90

