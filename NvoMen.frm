VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form NvoMen 
   BackColor       =   &H00C0C0C0&
   Caption         =   "SACC (Sistema de administración y control del comercio)"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   13305
   Icon            =   "NvoMen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   13305
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   3600
   End
   Begin VB.Frame Frame29 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11040
      TabIndex        =   124
      Top             =   120
      Width           =   975
      Begin VB.Image Image3 
         Height          =   750
         Left            =   120
         MouseIcon       =   "NvoMen.frx":1601A
         MousePointer    =   99  'Custom
         Picture         =   "NvoMen.frx":16324
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Avisos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   125
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image27 
         Height          =   750
         Left            =   120
         MouseIcon       =   "NvoMen.frx":167FE
         MousePointer    =   99  'Custom
         Picture         =   "NvoMen.frx":16B08
         Top             =   120
         Width           =   780
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   3600
   End
   Begin VB.TextBox TxtEmp 
      Height          =   285
      Index           =   12
      Left            =   960
      TabIndex        =   123
      Text            =   "Text2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame22 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   12120
      TabIndex        =   121
      Top             =   120
      Width           =   975
      Begin VB.Image Image2 
         Height          =   195
         Left            =   720
         Picture         =   "NvoMen.frx":17076
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image Image20 
         Height          =   630
         Left            =   120
         MouseIcon       =   "NvoMen.frx":17392
         MousePointer    =   99  'Custom
         Picture         =   "NvoMen.frx":1769C
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tickets"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   122
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox TxtProvider 
      Height          =   285
      Left            =   1920
      TabIndex        =   119
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtEmp 
      Height          =   285
      Index           =   11
      Left            =   840
      TabIndex        =   118
      Text            =   "Text2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TxtBaseDatos 
      Height          =   285
      Left            =   1920
      TabIndex        =   117
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtUsuario 
      Height          =   285
      Left            =   1920
      TabIndex        =   116
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtContrasena 
      Height          =   285
      Left            =   1920
      TabIndex        =   115
      Top             =   6960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   78
      Left            =   240
      TabIndex        =   114
      Top             =   8040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TxtEmp 
      Height          =   285
      Index           =   10
      Left            =   720
      TabIndex        =   111
      Text            =   "Text2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TxtEmp 
      Height          =   285
      Index           =   9
      Left            =   600
      TabIndex        =   110
      Text            =   "Text2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TxtEmp 
      Height          =   285
      Index           =   8
      Left            =   480
      TabIndex        =   109
      Text            =   "Text2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TxtEmp 
      Height          =   285
      Index           =   7
      Left            =   360
      TabIndex        =   108
      Text            =   "Text2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TxtEmp 
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   107
      Text            =   "Text2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TxtEmp 
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   106
      Text            =   "Text2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TxtEmp 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   105
      Text            =   "Text2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TxtEmp 
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   104
      Text            =   "Text2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TxtEmp 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   103
      Text            =   "Text2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TxtEmp 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   102
      Text            =   "Text2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TxtEmp 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   101
      Text            =   "Text2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   77
      Left            =   120
      TabIndex        =   100
      Top             =   8040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   76
      Left            =   1320
      TabIndex        =   98
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   75
      Left            =   1200
      TabIndex        =   97
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   74
      Left            =   1080
      TabIndex        =   99
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   73
      Left            =   960
      TabIndex        =   92
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   72
      Left            =   840
      TabIndex        =   91
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   71
      Left            =   720
      TabIndex        =   90
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   70
      Left            =   600
      TabIndex        =   89
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   69
      Left            =   480
      TabIndex        =   88
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   68
      Left            =   360
      TabIndex        =   87
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   67
      Left            =   240
      TabIndex        =   86
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   66
      Left            =   120
      TabIndex        =   85
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   65
      Left            =   1320
      TabIndex        =   84
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   64
      Left            =   1200
      TabIndex        =   83
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   63
      Left            =   1080
      TabIndex        =   82
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   62
      Left            =   960
      TabIndex        =   81
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   61
      Left            =   840
      TabIndex        =   80
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   60
      Left            =   720
      TabIndex        =   79
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   59
      Left            =   600
      TabIndex        =   78
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   58
      Left            =   480
      TabIndex        =   77
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   57
      Left            =   360
      TabIndex        =   76
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   56
      Left            =   240
      TabIndex        =   75
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   55
      Left            =   120
      TabIndex        =   74
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   54
      Left            =   1320
      TabIndex        =   73
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   53
      Left            =   1200
      TabIndex        =   72
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   52
      Left            =   1080
      TabIndex        =   71
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   51
      Left            =   960
      TabIndex        =   70
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   50
      Left            =   840
      TabIndex        =   69
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   49
      Left            =   720
      TabIndex        =   68
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   48
      Left            =   600
      TabIndex        =   67
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   47
      Left            =   480
      TabIndex        =   66
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   46
      Left            =   360
      TabIndex        =   65
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   45
      Left            =   240
      TabIndex        =   64
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   44
      Left            =   120
      TabIndex        =   63
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   43
      Left            =   1320
      TabIndex        =   62
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   42
      Left            =   1200
      TabIndex        =   61
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   41
      Left            =   1080
      TabIndex        =   60
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   40
      Left            =   960
      TabIndex        =   59
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   39
      Left            =   840
      TabIndex        =   58
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   38
      Left            =   720
      TabIndex        =   57
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   37
      Left            =   600
      TabIndex        =   56
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   36
      Left            =   480
      TabIndex        =   55
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   35
      Left            =   360
      TabIndex        =   54
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   34
      Left            =   240
      TabIndex        =   53
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   33
      Left            =   120
      TabIndex        =   52
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   32
      Left            =   1320
      TabIndex        =   51
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   31
      Left            =   1200
      TabIndex        =   50
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   30
      Left            =   1080
      TabIndex        =   49
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   29
      Left            =   960
      TabIndex        =   48
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   28
      Left            =   840
      TabIndex        =   47
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   27
      Left            =   720
      TabIndex        =   46
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   600
      TabIndex        =   45
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   25
      Left            =   480
      TabIndex        =   44
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   24
      Left            =   360
      TabIndex        =   43
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   240
      TabIndex        =   42
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   120
      TabIndex        =   41
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   21
      Left            =   1320
      TabIndex        =   40
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   1200
      TabIndex        =   39
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   1080
      TabIndex        =   38
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   960
      TabIndex        =   37
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   840
      TabIndex        =   36
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   720
      TabIndex        =   35
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   600
      TabIndex        =   34
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   480
      TabIndex        =   33
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   360
      TabIndex        =   32
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   31
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   120
      TabIndex        =   30
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   1320
      TabIndex        =   29
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   1200
      TabIndex        =   28
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   27
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   960
      TabIndex        =   26
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   840
      TabIndex        =   25
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   -3360
      ScaleHeight     =   1395
      ScaleWidth      =   34935
      TabIndex        =   96
      Top             =   0
      Width           =   35000
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   720
      TabIndex        =   18
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   17
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   16
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   15
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   120
      Top             =   5280
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   27
      CheckBorderColor=   7021576
      SelMenuBorder   =   7021576
      SelMenuBackColor=   14073525
      SelMenuForeColor=   16777215
      SelCheckBackColor=   13740436
      MenuBorderColor =   6956042
      SeparatorColor  =   -2147483632
      MenuBackColor   =   14737632
      MenuForeColor   =   0
      CheckBackColor  =   15326939
      CheckForeColor  =   4194304
      DisabledMenuBorderColor=   -2147483632
      DisabledMenuBackColor=   16777215
      DisabledMenuForeColor=   -2147483631
      MenuBarBackColor=   -2147483633
      MenuPopupBackColor=   16777215
      ShortCutNormalColor=   0
      ShortCutSelectColor=   16777215
      ArrowNormalColor=   4194304
      ArrowSelectColor=   12484864
      ShadowColor     =   0
      Bmp:1           =   "NvoMen.frx":19076
      Mask:1          =   16777215
      Key:1           =   "#SubSistema"
      Bmp:2           =   "NvoMen.frx":19618
      Mask:2          =   16777215
      Key:2           =   "#SubEliminar"
      Bmp:3           =   "NvoMen.frx":1996A
      Mask:3          =   16777215
      Key:3           =   "#SubJuegosdeReparacion"
      Bmp:4           =   "NvoMen.frx":19E6C
      Mask:4          =   16777215
      Key:4           =   "#SubProcesos"
      Bmp:5           =   "NvoMen.frx":1A1BE
      Mask:5          =   16777215
      Key:5           =   "#SubRevision"
      Bmp:6           =   "NvoMen.frx":1A510
      Mask:6          =   16777215
      Key:6           =   "#SubBloquear"
      Bmp:7           =   "NvoMen.frx":1AA4E
      Mask:7          =   16777215
      Key:7           =   "#SubCerrarSesion"
      Bmp:8           =   "NvoMen.frx":1AF8C
      Mask:8          =   16777215
      Key:8           =   "#SubAtención"
      Bmp:9           =   "NvoMen.frx":1B48E
      Mask:9          =   16777215
      Key:9           =   "#SubAdministración"
      Bmp:10          =   "NvoMen.frx":1B990
      Mask:10         =   16777215
      Key:10          =   "#SubSalir"
      Bmp:11          =   "NvoMen.frx":1BE1A
      Mask:11         =   16777215
      Key:11          =   "#SubVentas"
      Bmp:12          =   "NvoMen.frx":1C31C
      Mask:12         =   16777215
      Key:12          =   "#SubNuevo"
      Bmp:13          =   "NvoMen.frx":1C66E
      Mask:13         =   16777215
      Key:13          =   "#SubCotizar"
      Bmp:14          =   "NvoMen.frx":1CB70
      Mask:14         =   16777215
      Key:14          =   "#SubOrdenesdeCompra"
      Bmp:15          =   "NvoMen.frx":1D072
      Mask:15         =   16777215
      Key:15          =   "#SubMateriaPrima"
      Bmp:16          =   "NvoMen.frx":1D4FC
      Mask:16         =   16777215
      Key:16          =   "#SubPedidosAlmacen"
      Bmp:17          =   "NvoMen.frx":1D90E
      Mask:17         =   16777215
      Key:17          =   "#SubMovimientos"
      Bmp:18          =   "NvoMen.frx":1DF7C
      Mask:18         =   16777215
      Key:18          =   "#SubEntradas"
      Bmp:19          =   "NvoMen.frx":1E562
      Mask:19         =   16777215
      Key:19          =   "#SubRevisiones"
      Bmp:20          =   "NvoMen.frx":1EA64
      Mask:20         =   16777215
      Key:20          =   "#SubSoporteTecnico"
      Bmp:21          =   "NvoMen.frx":1EEEE
      Mask:21         =   16777215
      Key:21          =   "#SubMensajeros"
      Bmp:22          =   "NvoMen.frx":1F3C0
      Mask:22         =   16777215
      Key:22          =   "#SubContabilidad"
      Bmp:23          =   "NvoMen.frx":1F912
      Mask:23         =   16777215
      Key:23          =   "#SubConsultas"
      Bmp:24          =   "NvoMen.frx":1FDC4
      Mask:24         =   16777215
      Key:24          =   "#SubReportes"
      Bmp:25          =   "NvoMen.frx":203B6
      Mask:25         =   16777215
      Key:25          =   "#SubPedir"
      Bmp:26          =   "NvoMen.frx":207F8
      Mask:26         =   16777215
      Key:26          =   "#SubOpciones"
      Bmp:27          =   "NvoMen.frx":20F0A
      Key:27          =   "#MenContratos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7980
      Width           =   13305
      _ExtentX        =   23469
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "DESP"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "02:04 p.m."
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   0
            TextSave        =   "12/06/2023"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Text            =   "JLB Systems"
            TextSave        =   "JLB Systems"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Text            =   "Versión 2.8.0"
            TextSave        =   "Versión 2.8.0"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   9
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   7
      Left            =   960
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   6
      Left            =   840
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   5
      Left            =   720
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   6120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   113
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   7
      Left            =   960
      TabIndex        =   112
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   5
      Left            =   840
      TabIndex        =   24
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   4
      Left            =   720
      TabIndex        =   23
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   22
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   21
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   20
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtServidor 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LblEmpresa 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   975
      Left            =   840
      TabIndex        =   120
      Top             =   2400
      Width           =   10935
   End
   Begin VB.Label lblPuestoSucursal 
      BackStyle       =   0  'Transparent
      Caption         =   "Puesto y Sucursal"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   95
      Top             =   2160
      Width           =   8415
   End
   Begin VB.Label lblHola 
      BackStyle       =   0  'Transparent
      Caption         =   "Saludo Nombre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   720
      TabIndex        =   94
      Top             =   1680
      Width           =   8175
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1680
      TabIndex        =   93
      Top             =   4800
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   9240
      Picture         =   "NvoMen.frx":21C72
      Top             =   6000
      Width           =   3630
   End
   Begin VB.Menu MenVentas 
      Caption         =   "&Ventas"
      Begin VB.Menu SubVentas 
         Caption         =   "Ventas"
         Begin VB.Menu SubMenPuntodeVenta 
            Caption         =   "Punto de Venta"
         End
         Begin VB.Menu SubMenVentasProgramadas 
            Caption         =   "Ventas Programadas"
         End
         Begin VB.Menu SubMenVentasEspeciales 
            Caption         =   "Ventas Especiales"
         End
      End
      Begin VB.Menu SubAtención 
         Caption         =   "Atención"
         Begin VB.Menu SubMenTramitarGarantia 
            Caption         =   "Tramitar Garantía"
         End
         Begin VB.Menu SubMenAutorizarGarantia 
            Caption         =   "Autorizar Garantía"
         End
         Begin VB.Menu R3 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenAutorizarRemanufactura 
            Caption         =   "Autorizar Remanufactura"
         End
         Begin VB.Menu R4 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenPresta 
            Caption         =   "Prestamos"
         End
      End
      Begin VB.Menu SubAdministración 
         Caption         =   "Administración"
         Begin VB.Menu SubMenAutAltaCliente 
            Caption         =   "Autorizar Alta de Cliente"
         End
         Begin VB.Menu R40 
            Caption         =   "-"
         End
         Begin VB.Menu SumMenCortedeCaja 
            Caption         =   "Corte de Caja"
         End
         Begin VB.Menu r22 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenCancelaciones 
            Caption         =   "Cancelar Ventas/Facturas"
         End
         Begin VB.Menu SubMenCanComand 
            Caption         =   "Cancelar Comandas"
         End
         Begin VB.Menu SubMenCanRefa 
            Caption         =   "Cancelar/Refactura"
         End
         Begin VB.Menu r5 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenValedeCajaVentas 
            Caption         =   "Vale de Caja"
         End
         Begin VB.Menu SubMenLicitacion 
            Caption         =   "Licitación"
         End
         Begin VB.Menu SubMenPromocion 
            Caption         =   "Promoción"
         End
         Begin VB.Menu SubMenPermisos 
            Caption         =   "Permisos"
         End
         Begin VB.Menu SubMenNotaCredito 
            Caption         =   "Notas de Crédito"
         End
         Begin VB.Menu SubMenSancion 
            Caption         =   "Sanción a venta"
         End
      End
      Begin VB.Menu SubOpciones 
         Caption         =   "Herramientas"
         Begin VB.Menu SubMenReimprimir 
            Caption         =   "Reimprimir"
         End
         Begin VB.Menu r19 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenVerComandasPendientes 
            Caption         =   "Ver Comandas Pendientes"
         End
         Begin VB.Menu SubMenBuscarComanda 
            Caption         =   "Buscar Comanda"
         End
         Begin VB.Menu SubMenConsCobCom 
            Caption         =   "Consultar Cobro de Comanda"
         End
         Begin VB.Menu r23 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenCambiarVentadeSucursal 
            Caption         =   "Cambiar Venta de Sucursal"
         End
         Begin VB.Menu SubMenCambiarVentadeCliente 
            Caption         =   "Cambiar Cliente de Venta/Comanda"
         End
         Begin VB.Menu SubMenCambiarFormadePago 
            Caption         =   "Cambiar Forma de Pago"
         End
      End
   End
   Begin VB.Menu MenCompras 
      Caption         =   "&Compras"
      Begin VB.Menu SubCotizar 
         Caption         =   "Cotizar"
         Begin VB.Menu SubMenRequisicion 
            Caption         =   "Requisición"
         End
         Begin VB.Menu SubMenRevisar 
            Caption         =   "Revisar Cotizaciones"
         End
         Begin VB.Menu SubMenAsignar 
            Caption         =   "Asignar Proveedor"
         End
         Begin VB.Menu SubmenPreordendecompra 
            Caption         =   "Pre-orden de Compra"
         End
      End
      Begin VB.Menu SubOrdenesdeCompra 
         Caption         =   "Ordenes de Compra"
         Begin VB.Menu SubMenAutorizar 
            Caption         =   "Autorizar"
         End
         Begin VB.Menu SubMenImprimirOrdendeCompra 
            Caption         =   "Imprimir Orden de Compra"
         End
         Begin VB.Menu r1 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenOrdendeCompra 
            Caption         =   "Orden de Compra Rápida"
         End
         Begin VB.Menu SubMenModOR 
            Caption         =   "Modificar Orden Rápida"
         End
         Begin VB.Menu SubMenCancelarOrdenRapida 
            Caption         =   "Cancelar Orden Rápida"
         End
      End
      Begin VB.Menu SubMateriaPrima 
         Caption         =   "Proveedores Varios"
         Begin VB.Menu SubMenCompraenAlmacen1 
            Caption         =   "Compra a Proveedores Varios"
         End
         Begin VB.Menu SubMenCancelarCompraAlmacen1 
            Caption         =   "Cancelar Compra a  Proveedores Varios"
         End
      End
   End
   Begin VB.Menu MenAlmacen 
      Caption         =   "&Almacén"
      Begin VB.Menu SubPedidosAlmacen 
         Caption         =   "Pedidos"
         Begin VB.Menu SubMenHacerRequisicion 
            Caption         =   "Hacer Requisición"
         End
         Begin VB.Menu SubMenMaxMinAlma3 
            Caption         =   "Máximos y Mínimos Almacén 3"
         End
      End
      Begin VB.Menu SubMovimientos 
         Caption         =   "Movimientos"
         Begin VB.Menu SubMenInventarios 
            Caption         =   "Inventarios"
         End
         Begin VB.Menu r7 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenSurtirVentaProgramada 
            Caption         =   "Surtir Venta Programada"
         End
         Begin VB.Menu l1 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenSalidas 
            Caption         =   "Salidas por Garantía"
         End
         Begin VB.Menu r8 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenPRoducir 
            Caption         =   "Producir/Reemplazar Existencias"
         End
         Begin VB.Menu SubMenReemplazar 
            Caption         =   "Reemplazar"
            Visible         =   0   'False
         End
         Begin VB.Menu r9 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenPrestamos 
            Caption         =   "Prestamos a Clientes"
         End
         Begin VB.Menu r30 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenVentasCanceladas 
            Caption         =   "Ver Ventas Canceladas"
         End
      End
      Begin VB.Menu SubEntradas 
         Caption         =   "Entradas"
         Begin VB.Menu SubMenLlegadadeOrdendeCompra 
            Caption         =   "Llegada de Orden de Compra"
         End
         Begin VB.Menu SubMenEntrOCR 
            Caption         =   "Entrada Orden Rápida"
         End
         Begin VB.Menu SubMenEntrOrdenProd 
            Caption         =   "Entrada de Orden de Producción"
         End
         Begin VB.Menu SubMenAlmacen1 
            Caption         =   "Entrada Compra a  Proveedores Varios"
         End
         Begin VB.Menu r16 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenTraspasosEntreSucursales 
            Caption         =   "Traspasos Entre Sucursales"
         End
      End
   End
   Begin VB.Menu MenProducción 
      Caption         =   "&Producción"
      Begin VB.Menu SubProcesos 
         Caption         =   "Procesos"
         Begin VB.Menu SubMenRevision 
            Caption         =   "Revisión"
         End
         Begin VB.Menu SubMenProduccion 
            Caption         =   "Producción"
         End
         Begin VB.Menu SubMenCalidad 
            Caption         =   "Calidad"
         End
      End
      Begin VB.Menu SubRevisiones 
         Caption         =   "Revisiónes"
         Begin VB.Menu SubMenOrdenMaxMin 
            Caption         =   "Producciones de Máximos y Mínimos"
         End
         Begin VB.Menu PedMaxMin 
            Caption         =   "Pedido de Máximos y Mínimos"
         End
         Begin VB.Menu r20 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenScrap 
            Caption         =   "Scrap"
         End
         Begin VB.Menu SubMenMaterialExtra 
            Caption         =   "Material Extra"
         End
         Begin VB.Menu L13 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenRevisarCompraenAlmacen1 
            Caption         =   "Revisar Compra a  Proveedores Varios"
         End
      End
      Begin VB.Menu SubJuegosdeReparacion 
         Caption         =   "Juegos de Reparación"
         Begin VB.Menu SubMenNuevo 
            Caption         =   "Nuevo"
         End
         Begin VB.Menu SubMenReemplazarInsumo 
            Caption         =   "Reemplazar Insumo"
         End
         Begin VB.Menu SubMenVer 
            Caption         =   "Ver"
         End
      End
   End
   Begin VB.Menu MenAdministración 
      Caption         =   "Ad&ministración"
      Begin VB.Menu SubNuevo 
         Caption         =   "Agregar"
         Begin VB.Menu SubMenNuevoAgente 
            Caption         =   "Nuevo Usuario"
         End
         Begin VB.Menu SubMenNuevoCliente 
            Caption         =   "Nuevo Cliente"
         End
         Begin VB.Menu SubMenNvoDpto 
            Caption         =   "Nuevo Departamento"
         End
         Begin VB.Menu SubMenNuevaSucursal 
            Caption         =   "Nueva Sucursal"
         End
         Begin VB.Menu SubMenNuevoProveedor 
            Caption         =   "Nuevo Proveedor"
         End
         Begin VB.Menu SumMenProvRapi 
            Caption         =   "Nuevo Proveedor Orden Rápida"
         End
         Begin VB.Menu SubMenNuevoMensajero 
            Caption         =   "Nuevo Mensajero"
         End
         Begin VB.Menu SubMenNuevoPRoducto 
            Caption         =   "Nuevo Producto"
         End
         Begin VB.Menu SubMenNuevaMateriaPrima 
            Caption         =   "Nueva Materia Prima"
         End
         Begin VB.Menu subfrmfrmaltaeua 
            Caption         =   "Nueva Bodega"
         End
      End
      Begin VB.Menu SubEliminar 
         Caption         =   "Eliminar/Modificar"
         Begin VB.Menu SubMenEliminarAgente 
            Caption         =   "Eliminar/Modificar Usuario"
         End
         Begin VB.Menu SubMenEliminarCliente 
            Caption         =   "Eliminar/Modificar Cliente"
         End
         Begin VB.Menu SubMenEliminarSucursal 
            Caption         =   "Eliminar/Modificar Sucursal"
         End
         Begin VB.Menu SubMenEliminarProveedor 
            Caption         =   "Eliminar/Modificar Proveedor"
         End
         Begin VB.Menu SubMenEliminarMensajero 
            Caption         =   "Eliminar/Modificar Mensajero"
         End
         Begin VB.Menu SubMenEliminarProducto 
            Caption         =   "Eliminar/Modificar Producto"
         End
         Begin VB.Menu SubMenEliminarMateriaPrima 
            Caption         =   "Eliminar/Modificar Materia Prima"
         End
      End
      Begin VB.Menu SubSistema 
         Caption         =   "Sistema"
         Begin VB.Menu SubMenDatosdelaEmpresa 
            Caption         =   "Datos de la Empresa"
         End
         Begin VB.Menu ImportarPrecios 
            Caption         =   "Importar Precios / Inventarios"
         End
         Begin VB.Menu r15 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenMonitoreo 
            Caption         =   "Monitoreo"
         End
         Begin VB.Menu SubMenExploBD 
            Caption         =   "Explorador Base de Datos"
         End
         Begin VB.Menu SubMenConfigCorreo 
            Caption         =   "Configurar Correo"
         End
         Begin VB.Menu SubMenBasedeDatos 
            Caption         =   "Base de Datos"
         End
         Begin VB.Menu RA1 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenEmpatarFolios 
            Caption         =   "Empatar Folios"
         End
      End
   End
   Begin VB.Menu MenDepartamentosVarios 
      Caption         =   "&Departamentos Varios"
      Begin VB.Menu SubSoporteTecnico 
         Caption         =   "Soporte Técnico"
      End
      Begin VB.Menu SubMensajeros 
         Caption         =   "Mensajeros"
      End
      Begin VB.Menu SubContabilidad 
         Caption         =   "Contabilidad"
         Begin VB.Menu SubMenPagarOrdendeCompra 
            Caption         =   "Pagar Orden de Compra"
         End
         Begin VB.Menu SubMenPagodeOrdenRapida 
            Caption         =   "Pago de Orden Rápida"
         End
         Begin VB.Menu SubMenPagodeCompraAlmacen1 
            Caption         =   "Pago de Compra a  Proveedores Varios"
         End
         Begin VB.Menu SubMenCanAbono 
            Caption         =   "Cancelar Abonos a Ordenes"
         End
         Begin VB.Menu r10 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenAbonaraCreditos 
            Caption         =   "Abonar a Créditos"
         End
         Begin VB.Menu SubMenIncobrable 
            Caption         =   "Cuentas Incobrables"
         End
         Begin VB.Menu SubMenNotadeCredito 
            Caption         =   "Nota de Crédito"
            Visible         =   0   'False
         End
         Begin VB.Menu SubMenValedeCaja 
            Caption         =   "Vale de Caja"
            Visible         =   0   'False
         End
         Begin VB.Menu SubMenTipodeCambio 
            Caption         =   "Tipo de Cambio"
         End
         Begin VB.Menu SubMenExportaConta 
            Caption         =   "Exportar Contabilidad"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu MenContratos 
         Caption         =   "Contratos"
         Begin VB.Menu SubMenUsoActFijo 
            Caption         =   "Uso de Activo Fijo"
         End
      End
   End
   Begin VB.Menu MenUtilerias 
      Caption         =   "&Utilerías"
      Begin VB.Menu SubConsultas 
         Caption         =   "Consultas"
         Begin VB.Menu SubMenCliente 
            Caption         =   "Cliente"
         End
         Begin VB.Menu SubMenProveedores 
            Caption         =   "Proveedores"
         End
         Begin VB.Menu SubMenProductos 
            Caption         =   "Productos"
         End
         Begin VB.Menu r11 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenConVentas 
            Caption         =   "Información de Ventas"
         End
         Begin VB.Menu SubMenProductosPendientesLicitacion 
            Caption         =   "Productos Pendientes Licitación"
         End
         Begin VB.Menu SubMenVerGanancias 
            Caption         =   "Ver Ganancias"
         End
         Begin VB.Menu r12 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenProdPeriodo 
            Caption         =   "Producciónes por Fechas"
         End
         Begin VB.Menu SubMenProduccionesporFechas 
            Caption         =   "Total Producciónes por Fechas"
         End
         Begin VB.Menu r13 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenCompProv 
            Caption         =   "Compras al Proveedor"
         End
         Begin VB.Menu SubMenComprasdelProducto 
            Caption         =   "Compras Pendientes de Entrada"
         End
         Begin VB.Menu SubMenProductosPedidos 
            Caption         =   "Productos Pedidos"
         End
         Begin VB.Menu r14 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenVerEntradas 
            Caption         =   "Ver Entradas"
         End
         Begin VB.Menu SubMenFaltantes 
            Caption         =   "Faltantes"
         End
         Begin VB.Menu SubMenRastrear 
            Caption         =   "Rastrear"
         End
         Begin VB.Menu SubMenVerExistencias 
            Caption         =   "Ver Existencias"
         End
         Begin VB.Menu LI1 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenGraficasTickets 
            Caption         =   "Gràficas de Tickets"
         End
         Begin VB.Menu SubMenRendimiento 
            Caption         =   "Rendimiento de Departamentos"
         End
      End
      Begin VB.Menu SubPedir 
         Caption         =   "Pedir"
         Begin VB.Menu SubMenHacerPedido 
            Caption         =   "Hacer Pedido a Almacén"
         End
         Begin VB.Menu SubMenOrdenes 
            Caption         =   "Ordenes de Producción"
         End
      End
      Begin VB.Menu SubReportes 
         Caption         =   "Reportes"
         Begin VB.Menu SubMenReporteador 
            Caption         =   "Reporteador"
         End
         Begin VB.Menu SubMenRepInventarios 
            Caption         =   "Reporte de Inventarios"
         End
         Begin VB.Menu g1 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenReportesdeComprasdeClientes 
            Caption         =   "Reportes de Compras de Clientes"
         End
         Begin VB.Menu SubMenReportesdeVentas 
            Caption         =   "Reportes de Ventas"
         End
         Begin VB.Menu SubMenVentCan 
            Caption         =   "Reporte de Ventas Canceladas"
         End
         Begin VB.Menu SubMenVentasporUsuario 
            Caption         =   "Ventas por Sucursal/Usuario"
         End
         Begin VB.Menu SubMenReportedeLicitaciones 
            Caption         =   "Reporte de Licitaciónes"
         End
         Begin VB.Menu SubMenRepVenProg 
            Caption         =   "Reporte de Ventas Programadas Cierre/Facturación"
         End
         Begin VB.Menu g2 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenRepEstOc 
            Caption         =   "Reporte de Estado de OC"
         End
         Begin VB.Menu SubMenRepOCR 
            Caption         =   "Reporte de Ordenes Rapidas"
         End
         Begin VB.Menu scompras 
            Caption         =   "Reporte de Ordenes de Compra"
         End
         Begin VB.Menu SubMenReportedeComprasProveedor 
            Caption         =   "Reporte de Compras a Proveedor"
         End
         Begin VB.Menu SUBFRORPEN 
            Caption         =   "Reporte de Ordenes Pendientes de pago"
         End
         Begin VB.Menu SubMenRepOrdenPa 
            Caption         =   "Reporte de Ordenes Pagadas"
         End
         Begin VB.Menu SubMenRepComProvVar 
            Caption         =   "Reporte de Compras Proveedores Varios"
         End
         Begin VB.Menu SubMenRepOrdenCancel 
            Caption         =   "Reporte de Ordenes Canceladas"
         End
         Begin VB.Menu SubMenRepAbonos 
            Caption         =   "Reporte de Abonos"
         End
         Begin VB.Menu g8 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenRepMovimientos 
            Caption         =   "Concentrado de Movimientos"
         End
         Begin VB.Menu g3 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenRepEntradas 
            Caption         =   "Reporte de Entradas a Almacen"
         End
         Begin VB.Menu SubMenEntradaRapida 
            Caption         =   "Reporte de Entrada de Ordenes Rapidas"
         End
         Begin VB.Menu SubMenReporteExistencias 
            Caption         =   "Reporte de Existencias"
         End
         Begin VB.Menu g5 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenAbonos 
            Caption         =   "Reporte de CxC"
         End
         Begin VB.Menu SubMenRepCxCD 
            Caption         =   "Reporte CXC Detalle"
         End
         Begin VB.Menu SubMenGastos 
            Caption         =   "Reporte de Gastos"
         End
         Begin VB.Menu g4 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenRepPrestAvtFijo 
            Caption         =   "Reporte de Prestamo de Activo Fijo"
         End
         Begin VB.Menu G17 
            Caption         =   "-"
         End
         Begin VB.Menu SubMenCostosVentas 
            Caption         =   "Costos de Ventas"
         End
         Begin VB.Menu SubMenCostos 
            Caption         =   "Costos de Inventario"
         End
         Begin VB.Menu SubMenCostoJR 
            Caption         =   "Costo Juego de Reparación"
         End
         Begin VB.Menu SubMenRepPrestamos 
            Caption         =   "Prestamos a Clientes"
         End
      End
   End
   Begin VB.Menu MenEmpresa 
      Caption         =   "&Empresas"
      Begin VB.Menu SubMenEmp 
         Caption         =   "Empresas Nuevas"
         Index           =   0
      End
      Begin VB.Menu g12 
         Caption         =   "-"
      End
      Begin VB.Menu SubMenNuevaEmpresa 
         Caption         =   "Crear Nueva Empresa"
      End
   End
   Begin VB.Menu MenSalir 
      Caption         =   "&Salir"
      Begin VB.Menu SubBloquear 
         Caption         =   "Bloquear"
      End
      Begin VB.Menu SubCerrarSesion 
         Caption         =   "Cerrar Sesión"
      End
      Begin VB.Menu SubSalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "NvoMen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private cnn1 As ADODB.Connection
Private cnn2 As ADODB.Connection
' Esta clase se usará para seleccionar el fichero
Dim SubMen As Integer
Dim validar As Integer
' Tipos, constantes y funciones para FileExist
Const MAX_PATH = 260
Const INVALID_HANDLE_VALUE = -1
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

' FIN DE CODIGO DE FUNCIONES PARA SABER SI EXISTE O NO UN ARCHIVO
'------------------------------------------------------------------------------
' Clase para manejar ficheros INIs
' Permite leer secciones enteras y todas las secciones de un fichero INI
Private sBuffer As String   ' Para usarla en las funciones GetSection(s)
'--- Declaraciones para leer ficheros INI ---
' Leer todas las secciones de un fichero INI, no funciona en Win95
' Esta función no estaba en las declaraciones del API que se incluye con el VB
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
' Leer una sección completa
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
' Leer una clave de un fichero INI
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
' Funcion para ejecutar aplicacion externa al sistema (Factura Electronica) 20/Oct/2011 Armando H Valdez Arras
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Function IniGet(ByVal sFileName As String, ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "") As String
    ' Devuelve el valor de una clave de un fichero INI
    Dim RET As Long
    Dim sRetVal As String
    sRetVal = String$(255, 0)
    RET = GetPrivateProfileString(sSection, sKeyName, sDefault, sRetVal, Len(sRetVal), sFileName)
    If RET = 0 Then
        IniGet = sDefault
    Else
        IniGet = Left$(sRetVal, RET)
    End If
End Function
Private Function IniGetSection(ByVal sFileName As String, ByVal sSection As String) As String()
    '--------------------------------------------------------------------------
    ' Lee una sección entera de un fichero INI
    ' Adaptada para devolver un array de string
    '
    ' Esta función devolverá un array de índice cero
    ' con las claves y valores de la sección
    '
    ' Parámetros de entrada:
    '   sFileName   Nombre del fichero INI
    '   sSection    Nombre de la sección a leer
    ' Devuelve:
    '   Un array con el nombre de la clave y el valor
    '   Para leer los datos:
    '       For i = 0 To UBound(elArray) -1 Step 2
    '           sClave = elArray(i)
    '           sValor = elArray(i+1)
    '       Next
    Dim i As Long
    Dim j As Long
    Dim sTmp As String
    Dim sClave As String
    Dim sValor As String
    Dim aSeccion() As String
    Dim n As Long
    ReDim aSeccion(0)
    ' El tamaño máximo para Windows 95
    sBuffer = String$(32767, Chr$(0))
    n = GetPrivateProfileSection(sSection, sBuffer, Len(sBuffer), sFileName)
    If n Then
        ' Cortar la cadena al número de caracteres devueltos
        sBuffer = Left$(sBuffer, n)
        ' Quitar los vbNullChar extras del final
        i = InStr(sBuffer, vbNullChar & vbNullChar)
        If i Then
            sBuffer = Left$(sBuffer, i - 1)
        End If
        '
        n = -1
        ' Cada una de las entradas estará separada por un Chr$(0)
        Do
            i = InStr(sBuffer, Chr$(0))
            If i Then
                sTmp = LTrim$(Left$(sBuffer, i - 1))
                If Len(sTmp) Then
                    ' Comprobar si tiene el signo igual
                    j = InStr(sTmp, "=")
                    If j Then
                        sClave = Left$(sTmp, j - 1)
                        sValor = LTrim$(Mid$(sTmp, j + 1))
                        '
                        n = n + 2
                        ReDim Preserve aSeccion(n)
                        aSeccion(n - 1) = sClave
                        aSeccion(n) = sValor
                    End If
                End If
                sBuffer = Mid$(sBuffer, i + 1)
            End If
        Loop While i
        If Len(sBuffer) Then
            j = InStr(sBuffer, "=")
            If j Then
                sClave = Left$(sBuffer, j - 1)
                sValor = LTrim$(Mid$(sBuffer, j + 1))
                n = n + 2
                ReDim Preserve aSeccion(n)
                aSeccion(n - 1) = sClave
                aSeccion(n) = sValor
            End If
        End If
    End If
    ' Devolver el array
    IniGetSection = aSeccion
End Function
Private Function IniGetSections(ByVal sFileName As String) As String()
    ' Esta función devolverá un array con todas las secciones del fichero
    ' Parámetros de entrada:
    '   sFileName   Nombre del fichero INI
    ' Devuelve:
    '   Un array con todos los nombres de las secciones
    '   La primera sección estará en el elemento 1,
    '   por tanto, si el array contiene cero elementos es que no hay secciones
    Dim i As Long
    Dim sTmp As String
    Dim n As Long
    Dim aSections() As String
    ReDim aSections(0)
    ' El tamaño máximo para Windows 95
    sBuffer = String$(32767, Chr$(0))
    ' Esta función del API no está definida en el fichero TXT
    n = GetPrivateProfileSectionNames(sBuffer, Len(sBuffer), sFileName)
    If n Then
        ' Cortar la cadena al número de caracteres devueltos
        sBuffer = Left$(sBuffer, n)
        ' Quitar los vbNullChar extras del final
        i = InStr(sBuffer, vbNullChar & vbNullChar)
        If i Then
            sBuffer = Left$(sBuffer, i - 1)
        End If
        n = 0
        ' Cada una de las entradas estará separada por un Chr$(0)
        Do
            i = InStr(sBuffer, Chr$(0))
            If i Then
                sTmp = LTrim$(Left$(sBuffer, i - 1))
                If Len(sTmp) Then
                    n = n + 1
                    ReDim Preserve aSections(n)
                    aSections(n) = sTmp
                End If
                sBuffer = Mid$(sBuffer, i + 1)
            End If
        Loop While i
        If Len(sBuffer) Then
            n = n + 1
            ReDim Preserve aSections(n)
            aSections(n) = sBuffer
        End If
    End If
    ' Devolver el array
    IniGetSections = aSections
End Function
Private Function AppPath(Optional ByVal ConBackSlash As Boolean = True) As String
    ' Devuelve el path del ejecutable                               (23/Abr/02)
    ' con o sin la barra de directorios
    Dim s As String
    s = App.Path
    If ConBackSlash Then
        If Right$(s, 1) <> "\" Then
            s = s & "\"
        End If
    Else
        If Right$(s, 1) = "\" Then
            s = Left$(s, Len(s) - 1)
        End If
    End If
    AppPath = s
End Function
'------------------------------------------------------------------------------
' Fin del código para acceder a los ficheros INIs
'------------------------------------------------------------------------------
Public Function FileExist(ByVal sFile As String) As Boolean
    'comprobar si existe este fichero
    Dim WFD As WIN32_FIND_DATA
    Dim hFindFile As Long
    hFindFile = FindFirstFile(sFile, WFD)
    'Si no se ha encontrado
    If hFindFile = INVALID_HANDLE_VALUE Then
        FileExist = False
    Else
        FileExist = True
        'Cerrar el handle de FindFirst
        hFindFile = FindClose(hFindFile)
    End If
End Function
Private Sub Form_Activate()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Set tRs = New ADODB.Recordset
    Dim tLi As ListItem
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Me.lblHola.Caption = "Hola " & Trim(Me.Text1(1).Text) & " " & Trim(Me.Text1(2).Text) & "!"
    Me.lblPuestoSucursal.Caption = Trim(Me.Text1(3).Text) & " en " & Trim(Me.Text4(0).Text)
    'Sincronizar
    Me.lblEstado.Caption = "Buscando mensajes"
    Me.lblEstado.ForeColor = vbBlue
    DoEvents
    StatusBar1.Panels(8).Text = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    Me.lblEstado.Caption = ""
    'AVISO DE RECORDATORIOS
    sBuscar = "SELECT MENSAJE FROM RECORDATORIOS WHERE (TIPO = 'A') AND (DAY(FECHA_RECORDAR) = DAY(GETDATE())) AND (MONTH(FECHA_RECORDAR) = MONTH(GETDATE())) AND (DEPARTAMENTO = '" & VarMen.Text1(75).Text & "') UNION SELECT MENSAJE FROM RECORDATORIOS WHERE (TIPO = 'M') AND (DAY(FECHA_RECORDAR) = DAY(GETDATE())) AND (DEPARTAMENTO = '" & VarMen.Text1(75).Text & "') UNION SELECT MENSAJE FROM RECORDATORIOS WHERE (TIPO = 'U') AND (FECHA_RECORDAR = '" & Date & "') AND (DEPARTAMENTO = '" & VarMen.Text1(75).Text & "')"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Timer2.Enabled = True
    Else
        Timer2.Enabled = False
    End If
    'CORRIGE IVA EN VENTAS
    sBuscar = "SELECT VENTAS_DETALLE.ID_VENTA, VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.IMPORTE * ALMACEN3.IVA AS IVABIEN FROM VENTAS_DETALLE INNER JOIN ALMACEN3 ON VENTAS_DETALLE.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO WHERE (VENTAS_DETALLE.IVA = 0) AND (ALMACEN3.IVA > 0) AND (VENTAS_DETALLE.IMPORTE > 0)"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            sBuscar = "UPDATE VENTAS_DETALLE SET IVA = " & Format(tRs.Fields("IVABIEN"), "0.00") & " WHERE ID_VENTA = " & tRs.Fields("ID_VENTA") & " AND ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "'"
            cnn.Execute (sBuscar)
            tRs.MoveNext
        Loop
    End If
    'AVISO DE NUEVOS PENDIENTES EN HELPDESK
    sBuscar = "SELECT TICKETS.ID_TICKET FROM TICKETS INNER JOIN USUARIOS ON TICKETS.ID_USUARIO = USUARIOS.ID_USUARIO WHERE (TICKETS.DEPARTAMENTO_DESTINO = '" & VarMen.Text1(75).Text & "') AND (TICKETS.ESTATUS NOT IN ('F')) OR (TICKETS.ID_USUARIO_ATIENDE = '" & VarMen.Text1(0).Text & "') AND (TICKETS.ESTATUS NOT IN ('F')) OR (TICKETS.ESTATUS NOT IN ('F')) AND (TICKETS.ID_USUARIO = '" & VarMen.Text1(0).Text & "')"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If
    'ELIMINA EXISTENCIAS EN CEROS
    sBuscar = "DELETE FROM EXISTENCIAS WHERE CANTIDAD <= 0"
    cnn.Execute (sBuscar)
    'CERRAR CUENTAS EN CEROS
    sBuscar = "UPDATE CUENTAS SET PAGADA = 'S' WHERE (TOTAL_COMPRA = 0) AND (PAGADA = 'N')"
    cnn.Execute (sBuscar)
    'COREGIR ORDENES DE COMPRA QUE SUBTOTAL NO CUADRA CON SUMA DE PRECIOS
    sBuscar = "SELECT ID_ORDEN_COMPRA, ROUND(SUM(PRECIO * CANTIDAD), 2) AS SUB FROM ORDEN_COMPRA_DETALLE GROUP BY ID_ORDEN_COMPRA HAVING (ROUND(SUM(PRECIO * CANTIDAD), 2) <> (SELECT TOTAL FROM ORDEN_COMPRA WHERE (ID_ORDEN_COMPRA = ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA)))"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            sBuscar = "UPDATE ORDEN_COMPRA SET TOTAL = " & Format(tRs.Fields("SUB"), "0.00") & " WHERE (ID_ORDEN_COMPRA = " & tRs.Fields("ID_ORDEN_COMPRA") & ")"
            cnn.Execute (sBuscar)
            tRs.MoveNext
        Loop
    End If
    'COREGIR VENTAS CON IMPUESTOS NULL
    sBuscar = "UPDATE VENTAS SET IMPUESTO1 = 0, IMPUESTO2 = 0, RETENCION = 0 WHERE (ID_VENTA IN (SELECT ID_VENTA FROM VENTAS AS VENTAS_1 WHERE (IMPUESTO1 IS NULL) OR (IMPUESTO2 IS NULL) OR (RETENCION IS NULL)))"
    cnn.Execute (sBuscar)
    'COREGIR VENTAS QUE SUBTOTAL NO CUADRA CON SUMA DE PRECIOS
    sBuscar = "SELECT ID_VENTA, ROUND(SUM(IMPORTE), 2) AS SUB, ROUND(SUM(IMPORTE + IVA + IMPUESTO1 + IMPUESTO2 - RETENCION), 2) AS TOT, ROUND(SUM(IVA), 2) AS IVA, ROUND(SUM(IMPUESTO1), 2) AS IMP1, ROUND(SUM(IMPUESTO2), 2) AS IMP2, ROUND(SUM(RETENCION), 2) AS RET FROM VENTAS_DETALLE GROUP BY ID_VENTA HAVING (ROUND(SUM(IMPORTE + IVA + IMPUESTO1 + IMPUESTO2 - RETENCION), 2) <> ROUND ((SELECT TOTAL FROM VENTAS WHERE (ID_VENTA = dbo.VENTAS_DETALLE.ID_VENTA)), 2))"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            sBuscar = "UPDATE VENTAS SET SUBTOTAL = " & Format(tRs.Fields("SUB"), "0.00") & ", TOTAL = " & Format(tRs.Fields("TOT"), "0.00") & ", IVA = " & Format(tRs.Fields("IVA"), "0.00") & ", IMPUESTO1 = " & Format(tRs.Fields("IMP1"), "0.00") & ", IMPUESTO2 = " & Format(tRs.Fields("IMP2"), "0.00") & ", RETENCION = " & Format(tRs.Fields("RET"), "0.00") & " WHERE (ID_VENTA= " & tRs.Fields("ID_VENTA") & ")"
            cnn.Execute (sBuscar)
            tRs.MoveNext
        Loop
    End If
    ' ELIMINAR PRODUCTOS CON CANTIDAD CERO EN VENTAS
    sBuscar = "DELETE FROM VENTAS_DETALLE WHERE (CANTIDAD = 0)"
    cnn.Execute (sBuscar)
    ' CORRIGE COTIZACIONES SIN ID_PRODUCTO
    sBuscar = "SELECT ALMACEN3.ID_PRODUCTO, ALMACEN3.DESCRIPCION FROM ALMACEN3 INNER JOIN COTIZA_REQUI ON ALMACEN3.DESCRIPCION = COTIZA_REQUI.DESCRIPCION WHERE COTIZA_REQUI.ID_PRODUCTO = ''"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            sBuscar = "UPDATE COTIZA_REQUI SET ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' WHERE (ID_PRODUCTO = '') AND DESCRIPCION = '" & tRs.Fields("ID_PRODUCTO") & "'"
            cnn.Execute (sBuscar)
            tRs.MoveNext
        Loop
    End If
    'PASA LOS ID_VENTA A LAS CUENTAS (REQUIERE TABLA CUENTA_VENTA)
    sBuscar = "SELECT CUENTA_VENTA.ID_VENTA, CUENTA_VENTA.ID_CUENTA FROM CUENTAS INNER JOIN CUENTA_VENTA ON CUENTAS.ID_CUENTA = CUENTA_VENTA.ID_CUENTA WHERE (CUENTAS.ID_VENTA IS NULL)"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            sBuscar = "UPDATE CUENTAS SET ID_VENTA = '" & tRs.Fields("ID_VENTA") & "' WHERE ID_CUENTA = " & tRs.Fields("ID_CUENTA")
            cnn.Execute (sBuscar)
            tRs.MoveNext
        Loop
    End If
    'PASA LOS FOLIO DE FACTURA A LAS CUENTAS
    sBuscar = "SELECT VENTAS.FOLIO, CUENTAS.ID_CUENTA FROM CUENTAS INNER JOIN VENTAS ON CUENTAS.ID_VENTA = VENTAS.ID_VENTA WHERE (CUENTAS.FOLIO IS NULL) AND (VENTAS.FACTURADO = 1)"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            sBuscar = "UPDATE CUENTAS SET FOLIO = '" & tRs.Fields("FOLIO") & "' WHERE ID_CUENTA = " & tRs.Fields("ID_CUENTA")
            cnn.Execute (sBuscar)
            tRs.MoveNext
        Loop
    End If
    'MARCA CUENTAS SIN ABONOS COMO NO PAGADAS
    'sBuscar = "UPDATE CUENTAS SET PAGADA = 'N' WHERE ID_CUENTA IN(SELECT CUENTAS.ID_CUENTA FROM CUENTAS INNER JOIN CUENTA_VENTA ON CUENTAS.ID_CUENTA = CUENTA_VENTA.ID_CUENTA WHERE (CUENTAS.PAGADA = 'S') AND (CUENTAS.ID_VENTA NOT IN (SELECT ID_VENTA FROM ABONOS_CUENTA)))"
    sBuscar = "UPDATE CUENTAS SET PAGADA = 'N' WHERE ID_CUENTA IN(SELECT ID_CUENTA FROM CUENTAS WHERE (PAGADA = 'S') AND (FOLIO NOT IN (SELECT FOLIO FROM ABONOS_CUENTA)) AND (FOLIO <> ''))"
    cnn.Execute (sBuscar)
    'ELIMINA COTIZACIONES CON CANTIDAD EN CEROS
    sBuscar = "DELETE FROM COTIZA_REQUI WHERE (ESTADO_ACTUAL = 'X') AND (CANTIDAD = 0)"
    cnn.Execute (sBuscar)
    'ABRE CUENTAS CERRADAS QUE EL PAGO TIENE UNA DIFERENCIA DE 1 PESO O MAYOR
    sBuscar = "UPDATE CUENTAS SET PAGADA = 'N' WHERE ID_CUENTA IN (SELECT ID_CUENTA From CUENTAS WHERE (TOTAL_COMPRA > (SELECT SUM(CANT_ABONO + 0.99) AS TOT From ABONOS_CUENTA WHERE (ID_CUENTA = CUENTAS.ID_CUENTA))) AND (PAGADA = 'S'))"
    cnn.Execute (sBuscar)
    'INSERTA A DOMICILIO EN 0 Y GARANTIA EN 0 EN ASISTENCIAS QUE NO CONTIENEN EL VALOR
    sBuscar = "UPDATE ASISTENCIA_TECNICA SET A_DOMICILIO = 0 WHERE (A_DOMICILIO = '')"
    cnn.Execute (sBuscar)
    sBuscar = "UPDATE ASISTENCIA_TECNICA SET GARANTIA = 0 WHERE (GARANTIA = '')"
    cnn.Execute (sBuscar)
    'CIERRA CUENTAS DE DOS O MAS NOTAS DE VENTA QUE FUERON PAGADAS
    sBuscar = "UPDATE CUENTAS SET PAGADA = 'S' WHERE ID_CUENTA IN (SELECT ID_CUENTA FROM CUENTAS WHERE (PAGADA = 'N') AND (FOLIO IN (SELECT FOLIO FROM CUENTAS AS CUENTAS_1 WHERE (PAGADA = 'S'))) AND (FOLIO <> ''))"
    cnn.Execute (sBuscar)
    'PRODUCTOS CON RETENCION EN CERO
    sBuscar = "UPDATE ALMACEN3 SET RETENCION = 0 WHERE (RETENCION IS NULL)"
    cnn.Execute (sBuscar)
    'ELIMINA LAS CANTIDADES EN CERO DE COTIZA _REQUI
    sBuscar = "DELETE FROM COTIZA_REQUI WHERE CANTIDAD = 0"
    cnn.Execute (sBuscar)
    ' PONGO VENTAS PROGRAMADAS NO FINALIZADAS
    sBuscar = "UPDATE PED_CLIEN_DETALLE SET FINALIZADA = 'N' WHERE (FINALIZADA IS NULL)"
    cnn.Execute (sBuscar)
    
    
    'CREA DEUDA EN VENTAS DE CREDITO
    sBuscar = "SELECT ID_VENTA, ID_CLIENTE, ID_USUARIO, FECHA, TOTAL, SUCURSAL, FOLIO FROM VENTAS WHERE ID_VENTA NOT IN (SELECT ID_VENTA FROM CUENTA_VENTA) AND (UNA_EXIBICION = 'N') AND (FACTURADO IN (0, 1))"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            sBuscar = "INSERT INTO CUENTAS (ID_CLIENTE, ID_USUARIO, FECHA, DIAS_CREDITO, DESCUENTO, TOTAL_COMPRA, SUCURSAL, PAGADA, FOLIO, DEUDA, ID_VENTA) VALUES ('" & tRs.Fields("ID_CLIENTE") & "', '" & tRs.Fields("ID_USUARIO") & "', '" & Format(tRs.Fields("FECHA"), "dd/mm/yyyy") & "', '30', '0', '" & tRs.Fields("TOTAL") & "', '" & tRs.Fields("SUCURSAL") & "', 'N', '" & tRs.Fields("FOLIO") & "', '" & tRs.Fields("TOTAL") & "', '" & tRs.Fields("ID_VENTA") & "')"
            cnn.Execute (sBuscar)
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT ID_VENTA, ID_CUENTA FROM CUENTAS WHERE ID_CUENTA NOT IN (SELECT ID_CUENTA FROM CUENTA_VENTA)"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            sBuscar = "INSERT INTO CUENTA_VENTA (ID_CUENTA, ID_VENTA) VALUES ('" & tRs.Fields("ID_CUENTA") & "', '" & tRs.Fields("ID_VENTA") & "')"
            cnn.Execute (sBuscar)
            tRs.MoveNext
        Loop
    End If
    
    'CORRIGE PAGOS CON ID_PROVEEDOR EN CERO
    sBuscar = "SELECT ABONOS_PAGO_OC.ID_ABONO, ORDEN_RAPIDA.ID_PROVEEDOR FROM ABONOS_PAGO_OC INNER JOIN ORDEN_RAPIDA ON ABONOS_PAGO_OC.NUM_ORDEN = ORDEN_RAPIDA.ID_ORDEN_RAPIDA WHERE (ABONOS_PAGO_OC.ID_PROVEEDOR = 0)"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            sBuscar = "UPDATE ABONOS_PAGO_OC SET ID_PROVEEDOR = '" & tRs.Fields("ID_PROVEEDOR") & "' WHERE ID_ABONO = '" & tRs.Fields("ID_ABONO") & "'"
            cnn.Execute (sBuscar)
            tRs.MoveNext
        Loop
    End If
    
    'DoEvents
    ValidaMenu
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Exit Sub
ManejaError:
    If Err.Number = 384 Then 'ERROR DE AJUSTE EN PANTALLA CUANDO ESTA MAXIMIZADA
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & ". " & Err.Description & ".", vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub CreaMenu()
On Error GoTo ManejaError
    Dim sBaseDatos As String
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim i As Long
    sBaseDatos = "EMPRESAS_SACC"
    NvoMen.txtServidor.Text = GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "")
    Set cnn1 = New ADODB.Connection
    With cnn1
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & sBaseDatos & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    sBuscar = "SELECT ID, EMPRESA From EMPRESAS WHERE (STATUS = 'A') ORDER BY EMPRESA"
    Set tRs = cnn1.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            i = SubMenEmp.Count
            Load SubMenEmp(CDbl(tRs.Fields("ID")))
            With SubMenEmp(CDbl(tRs.Fields("ID")))
                .Caption = tRs.Fields("EMPRESA")
                .Visible = True
            End With
            tRs.MoveNext
        Loop
        SubMenEmp(0).Visible = False
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & ". " & Err.Description & ".", vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Dim sValor As String
    Dim Guarda As String
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    ' Checar que el sistema no se este  ejecutando e impedir que se ejecute por segunda vez
    'If App.PrevInstance = True Then
    '    MsgBox "El programa ya está siendo ejecutado", vbCritical, "SACC"
    '    Unload Me
        ' O también, puedes poner que el programa gane el foco cuando se abre por segunda vez
    'End If
    TxtContrasena.Text = GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX")
    TxtUsuario.Text = GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX")
    TxtBaseDatos.Text = GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER")
    txtServidor.Text = GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX")
    TxtProvider.Text = GetSetting("APTONER", "ConfigSACC", "PROVIDER", "SQLOLEDB.1")
    Set cnn1 = New ADODB.Connection
    With cnn1
        .ConnectionString = "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=EMPRESAS_SACC;Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDORPPAL", "LINUX") & ";"
        .Open
    End With
    If GetSetting("APTONER", "ConfigSACC", "LIC", "N") = "N" Then 'CONDICION SOLO PARA EL LICENCIADO
        sBuscar = "SELECT SERVIDOR FROM EMPRESAS WHERE BASE_DATOS = '" & TxtBaseDatos.Text & "'"
        Set tRs = cnn1.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            txtServidor.Text = tRs.Fields("SERVIDOR")
            SaveSetting "APTONER", "ConfigSACC", "SERVIDOR", tRs.Fields("SERVIDOR")
        End If
    End If
    FrmLogin.Label1.Caption = GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER")
    CreaMenu
    If FileExist(App.Path & "\Server.Ini") And GetSetting("APTONER", "ConfigSACC", "RegAprovSACC", "0") = "ValAprovReg" Then
        sValor = ""
        'txtServidor.Text = IniGet(App.Path & "\Server.Ini", "Servidor", "Nombre", sValor)
        Set cnn = New ADODB.Connection
        With cnn
            .ConnectionString = "Provider=" & NvoMen.TxtProvider.Text & ";;Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
            .Open
        End With
        sBuscar = "SELECT ID_USUARIO FROM USUARIOS"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.BOF And tRs.EOF) Then
            FrmLogin.Show vbModal, Me
        End If
    Else
        RegSACC.Show vbModal
        Unload Me
    End If
    If FileExist(App.Path & "\Server.Ini") Then
        Set cnn = New ADODB.Connection
        With cnn
            .ConnectionString = _
                "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
            .Open
        End With
    Else
        Exit Sub
    End If
    sBuscar = "SELECT FECHA FROM RESPALDOS_BD WHERE FECHA = '" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    If tRs.EOF And tRs.BOF Then
        ' Empareja las CXC en precios mal
        sBuscar = "SELECT ID_CUENTA, DEUDA, TOTAL_COMPRA From vsCxC WHERE(TOTAL_COMPRA <> DEUDA)"
        Set tRs = cnn.Execute(sBuscar)
        Do While Not tRs.EOF
            sBuscar = "UPDATE CUENTAS SET DEUDA = " & tRs.Fields("TOTAL_COMPRA") & " WHERE ID_CUENTA = " & tRs.Fields("ID_CUENTA")
            Set tRs1 = cnn.Execute(sBuscar)
            tRs.MoveNext
        Loop
        'Poner fecha de facturacion
        sBuscar = "SELECT TOP (1) FECHA From RESPALDOS_BD ORDER BY ID DESC"
        Set tRs = cnn.Execute(sBuscar)
        Do While Not tRs.EOF
            sBuscar = "UPDATE VENTAS SET FECHA_FACTURA = '" & Format(tRs.Fields("FECHA"), "dd/MM/yyyy") & "' WHERE (FECHA_FACTURA IS NULL) AND (FOLIO <> 'CANCELADO') AND (FACTURADO = '1')"
            Set tRs1 = cnn.Execute(sBuscar)
            tRs.MoveNext
        Loop
        'Respalda la BD
        If GetSetting("APTONER", "ConfigSACC", "CreaRespaldo", "N") = "S" Then
            Guarda = "C:\BackUpSACC" & Mid(TxtEmp(8).Text, 1, 4) & Format(Date, "dd/mm/yyyy") & ".Bak"
            Guarda = Replace(Guarda, "/", "-")
            sBuscar = "BACKUP DATABASE " & GetSetting("APTONER", "ConfigSACC", "DATABASE", "LINUX") & " TO DISK = '" & Guarda & "' WITH FORMAT,NAME = 'res'"
            cnn.Execute (sBuscar)
            sBuscar = "INSERT INTO RESPALDOS_BD (FECHA, NOMBRE_RESPALDO) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', 'BackUpSACC" & Mid(TxtEmp(8).Text, 1, 4) & Format(Date, "dd/mm/yyyy") & ".Bak')"
            cnn.Execute (sBuscar)
        End If
        'PONER COSTOS A JUEGOS DE REPARACION EN VENTAS
        sBuscar = "SELECT ID_PRODUCTO FROM VENTAS_DETALLE WHERE (PRECIO_COSTO Is Null) GROUP BY ID_PRODUCTO"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                sBuscar = "SELECT SUM(PRECIO_COSTO * CANTIDAD) AS TOT FROM VsCostoJR WHERE ID_REPARACION = '" & tRs.Fields("ID_PRODUCTO") & "'"
                Set tRs1 = cnn.Execute(sBuscar)
                If Not (tRs1.EOF And tRs1.BOF) Then
                    If tRs1.Fields("TOT") > 0 Then
                        sBuscar = "UPDATE VENTAS_DETALLE SET PRECIO_COSTO = " & tRs1.Fields("TOT") & ", GANANCIA = ((PRECIO_VENTA / " & tRs1.Fields("TOT") & ") -1) WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND (PRECIO_COSTO Is Null)"
                        cnn.Execute (sBuscar)
                    End If
                End If
                tRs.MoveNext
            Loop
        End If
        'PONER COSTO A PRODUCTOS DE ALMACEN3 EN VENTAS SEGUN SU PRECIO DE COMPRA
        sBuscar = "SELECT ID_PRODUCTO FROM VENTAS_DETALLE WHERE (PRECIO_COSTO Is Null) GROUP BY ID_PRODUCTO"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                sBuscar = "SELECT PRECIO FROM ORDEN_COMPRA_DETALLE WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "'"
                Set tRs1 = cnn.Execute(sBuscar)
                If Not (tRs1.EOF And tRs1.BOF) Then
                    If tRs1.Fields("PRECIO") > 0 Then
                        sBuscar = "UPDATE VENTAS_DETALLE SET PRECIO_COSTO = " & tRs1.Fields("PRECIO") & ", GANANCIA = ((PRECIO_VENTA / " & tRs1.Fields("PRECIO") & ") -1) WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND (PRECIO_COSTO Is Null)"
                        cnn.Execute (sBuscar)
                    End If
                End If
                tRs.MoveNext
            Loop
        End If
        'PONER COSTO A PRODUCTOS DE ALMACEN3 EN VENTAS SEGUN SU PRECIO ESTIMADO
        sBuscar = "SELECT ID_PRODUCTO FROM VENTAS_DETALLE WHERE (PRECIO_COSTO Is Null) GROUP BY ID_PRODUCTO"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                sBuscar = "SELECT PRECIO_COSTO FROM ALMACEN3 WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "'"
                Set tRs1 = cnn.Execute(sBuscar)
                If Not (tRs1.EOF And tRs1.BOF) Then
                    If tRs1.Fields("PRECIO_COSTO") > 0 Then
                        sBuscar = "UPDATE VENTAS_DETALLE SET PRECIO_COSTO = " & tRs1.Fields("PRECIO_COSTO") & ", GANANCIA = ((PRECIO_VENTA / " & tRs1.Fields("PRECIO_COSTO") & ") -1) WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND (PRECIO_COSTO Is Null)"
                        cnn.Execute (sBuscar)
                    End If
                End If
                tRs.MoveNext
            Loop
        End If
        'HACER VISIBLES REQUISICIONES ATORADAS O DESAPARECIDAS
        sBuscar = "UPDATE REQUISICION SET ALMACEN = 'A3', MARCA = '&Marca&' WHERE (ALMACEN = '') AND (ACTIVO = 0) AND (CONTADOR = 0) AND (COTIZADA = 0)"
        cnn.Execute (sBuscar)
        'IGUALAR CUENTAS Y VENTAS
        sBuscar = "SELECT VENTAS.TOTAL, CUENTAS.TOTAL_COMPRA, CUENTA_VENTA.ID_VENTA, CUENTA_VENTA.ID_CUENTA FROM VENTAS INNER JOIN CUENTA_VENTA ON VENTAS.ID_VENTA = CUENTA_VENTA.ID_VENTA INNER JOIN CUENTAS ON dbo.CUENTA_VENTA.ID_CUENTA = CUENTAS.ID_CUENTA AND VENTAS.TOTAL <> CUENTAS.TOTAL_COMPRA"
        Set tRs = cnn.Execute(sBuscar)
        Do While Not tRs.EOF
            sBuscar = "UPDATE CUENTAS SET TOTAL_COMPRA = " & tRs.Fields("TOTAL") & ", DEUDA = " & tRs.Fields("TOTAL") & " WHERE ID_CUENTA = " & tRs.Fields("ID_CUENTA")
            Set tRs1 = cnn.Execute(sBuscar)
            sBuscar = "UPDATE VENTAS SET TOTAL = " & tRs.Fields("TOTAL") & " WHERE ID_VENTA = " & tRs.Fields("ID_VENTA")
            Set tRs1 = cnn.Execute(sBuscar)
            tRs.MoveNext
        Loop
    End If
    ' PONER TIPO A COMANDAS FALTANTES
    sBuscar = "UPDATE COMANDAS_DETALLES_2 SET TIPO = 'I' WHERE TIPO NOT IN ('T', 'I') AND ID_PRODUCTO LIKE '__I%'"
    cnn.Execute (sBuscar)
    sBuscar = "UPDATE COMANDAS_DETALLES_2 SET TIPO = 'T' WHERE TIPO NOT IN ('T', 'I') AND ID_PRODUCTO LIKE '__T%'"
    cnn.Execute (sBuscar)
    'CERRAR CUENTAS CON DEUDA MENOS A UN PESO
    sBuscar = "UPDATE CUENTAS SET PAGADA = 'S' WHERE (ID_CUENTA IN (SELECT CUENTAS_1.ID_CUENTA FROM ABONOS_CUENTA INNER JOIN CUENTAS AS CUENTAS_1 ON ABONOS_CUENTA.ID_CUENTA = CUENTAS_1.ID_CUENTA WHERE (CUENTAS_1.TOTAL_COMPRA - ABONOS_CUENTA.CANT_ABONO <= 1) AND (CUENTAS_1.PAGADA = 'N')))"
    cnn.Execute (sBuscar)
    sBuscar = "SELECT * FROM EMPRESA"
    Set tRs = cnn.Execute(sBuscar)
    If tRs.EOF And tRs.BOF Then
        frmEmpresa.Show vbModal
    Else
        Text5(0).Text = tRs.Fields("NOMBRE")
        Text5(1).Text = tRs.Fields("DIRECCION")
        Text5(2).Text = tRs.Fields("TELEFONO")
        Text5(3).Text = tRs.Fields("FAX")
        Text5(4).Text = tRs.Fields("COLONIA")
        Text5(5).Text = tRs.Fields("CD")
        Text5(6).Text = tRs.Fields("ESTADO")
        Text5(7).Text = tRs.Fields("PAIS")
        Text5(8).Text = tRs.Fields("RFC")
        Text5(9).Text = tRs.Fields("CP")
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & ". " & Err.Description & ".", vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub frmfrmaltaeua_Click()
    frmaltaeua.Show vbModal
End Sub
Private Sub Form_Resize()
    Image1.Left = (NvoMen.Width) - (Image1.Width + 400)
    Frame22.Left = (NvoMen.Width) - (Frame22.Width + 400)
    Image1.Top = (NvoMen.Height) - (Image1.Height + 1200)
End Sub
Private Sub Image20_Click()
    FrmTickets.Show vbModal
End Sub
Private Sub Image27_Click()
    FrmRecordatorios.Show vbModal
End Sub
Private Sub Image3_Click()
    FrmRecordatorios.Show vbModal
End Sub
Private Sub ImportarPrecios_Click()
    FrmImportarPrecios.Show vbModal
End Sub
Private Sub PedMaxMin_Click()
    FrmMaxMin.Show vbModal
End Sub
Private Sub scompras_Click()
    Frmrepcomprass.Show vbModal
End Sub
Private Sub SubBloquear_Click()
    FrmLogin.Show vbModal, Me
End Sub
Private Sub SubCerrarSesion_Click()
    Unload Me
End Sub
Private Sub subfrmfrmaltaeua_Click()
    frmaltaeua.Show vbModal
End Sub
Private Sub SUBFRORPEN_Click()
    FRORPEN.Show vbModal
End Sub
Private Sub SubMenAbonaraCreditos_Click()
    Creditos.Show vbModal
End Sub
Private Sub SubMenAbonos_Click()
    FrmReporte.Show vbModal
End Sub
Private Sub SubMenAlmacen1_Click()
    FrmArpvCompAlm1.Show vbModal
End Sub
Private Sub SubMenAsignar_Click()
    frmAutorizarCotizaciones.Show vbModal
End Sub
Private Sub SubMenAutAltaCliente_Click()
    FrmAutAltaCliente.Show vbModal
End Sub
Private Sub SubMenAutorizar_Click()
    frmAutOC.Show vbModal
End Sub
Private Sub SubMenAutorizarGarantia_Click()
    FrmAutGarantia.Show vbModal
End Sub
Private Sub SubMenAutorizarRemanufactura_Click()
    frmAutRema.Show vbModal
End Sub
Private Sub SubMenBasedeDatos_Click()
    Unload Me
    frmConfig.Show vbModal
End Sub
Private Sub SubMenBuscarComanda_Click()
    FrmBuscaComanda.Show vbModal
End Sub
Private Sub SubMenCalidad_Click()
    frmCalidad.Show vbModal
End Sub
Private Sub SubMenCambiarFormadePago_Click()
    FrmModiVenta.Show vbModal
End Sub
Private Sub SubMenCambiarVentadeCliente_Click()
    FrmCamClienVent.Show vbModal
End Sub
Private Sub SubMenCambiarVentadeSucursal_Click()
    FrmCambioVenta.Show vbModal
End Sub
Private Sub SubMenCanAbono_Click()
    FrmCanChque.Show vbModal
End Sub
Private Sub SubMenCancelaciones_Click()
    frmCancelaFactura.Show vbModal
End Sub
Private Sub SubMenCancelarCompraAlmacen1_Click()
    FrmCancelCompCartVac.Show vbModal
End Sub
Private Sub SubMenCancelarOrdenRapida_Click()
    FrmCancelOrdenRapida.Show vbModal
End Sub
Private Sub SubMenCanComand_Click()
    FrmCancelaComanda.Show vbModal
End Sub
Private Sub SubMenCanRefa_Click()
    FrmRefactura.Show vbModal
End Sub
Private Sub SubMenCliente_Click()
     FrmVerClien.Show vbModal
End Sub
Private Sub SubMenCompProv_Click()
    FrmEstadoProveedor.Show vbModal
End Sub
Private Sub SubMenCompraenAlmacen1_Click()
    FrmCompAlm1.Show vbModal
End Sub
Private Sub SubMenComprasdelProducto_Click()
    FrmReporteCompras.Show vbModal
End Sub
Private Sub SubMenConfigCorreo_Click()
    FrmConfigCorreo.Show vbModal
End Sub
Private Sub SubMenConsCobCom_Click()
    FrmConsultaComanda.Show vbModal
End Sub
Private Sub SubMenConVentas_Click()
    FrmConsultaVentas.Show vbModal
End Sub
Private Sub SubMenCostoJR_Click()
    FrmCostoJR.Show vbModal
End Sub
Private Sub SubMenCostos_Click()
    FrmCostoInventario.Show vbModal
End Sub
Private Sub SubMenCostosVentas_Click()
    FrmRepCostosVentas.Show vbModal
End Sub
Private Sub SubMenDatosdelaEmpresa_Click()
    frmEmpresa.Show vbModal
End Sub
Private Sub SubMenEliminarAgente_Click()
    EliAgente.Show vbModal
End Sub
Private Sub SubMenEliminarCliente_Click()
    EliCliente.Show vbModal
End Sub
Private Sub SubMenEliminarMateriaPrima_Click()
    FrmAltaProdAlm1y2.Frame8.Visible = True
    FrmAltaProdAlm1y2.Frame16.Visible = True
    FrmAltaProdAlm1y2.Caption = "Eliminar Productos de Almacn 1 / 2"
    FrmAltaProdAlm1y2.Show vbModal
End Sub
Private Sub SubMenEliminarMensajero_Click()
    EliMensajero.Show vbModal
End Sub
Private Sub SubMenEliminarProducto_Click()
    FrmAltaProdAlm3.Frame16.Visible = True
    FrmAltaProdAlm3.Frame8.Visible = True
    FrmAltaProdAlm3.Caption = "Eliminar Productos de Almacen 3"
    FrmAltaProdAlm3.Show vbModal
End Sub
Private Sub SubMenEliminarProveedor_Click()
    EliProveedor.Show vbModal
End Sub
Private Sub SubMenEliminarSucursal_Click()
    EliSuc.Show vbModal
End Sub
Private Sub SubMenEmp_Click(Index As Integer)
On Error GoTo ManejaError
    Dim sBaseDatos As String
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    sBaseDatos = "EMPRESAS_SACC"
    Set cnn1 = New ADODB.Connection
    With cnn1
        .ConnectionString = "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & sBaseDatos & ";Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDORPPAL", "") & ";"  '& NvoMen.txtServidor.Text & ";"
        'MsgBox "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & sBaseDatos & ";Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDORPPAL", "") & ";"  '& NvoMen.txtServidor.Text & ";"
        .Open
    End With
    sBuscar = "SELECT BASE_DATOS, SERVIDOR, COMPRAS_NAC_INT From EMPRESAS WHERE ID = " & Index
    Set tRs = cnn1.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        sBaseDatos = tRs.Fields("BASE_DATOS")
        NvoMen.txtServidor.Text = tRs.Fields("SERVIDOR")
        Set cnn2 = New ADODB.Connection
        With cnn2
            .ConnectionString = "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & sBaseDatos & ";Data Source=" & NvoMen.txtServidor.Text & ";"
            .Open
        End With
        sBuscar = "SELECT NOMBRE FROM USUARIOS WHERE NOMBRE = '" & VarMen.Text1(1).Text & "' AND ESTADO = 'A'"
        Set tRs1 = cnn2.Execute(sBuscar)
        If Not (tRs1.EOF And tRs1.BOF) Then
            SaveSetting "APTONER", "ConfigSACC", "DATABASE", sBaseDatos
            SaveSetting "APTONER", "ConfigSACC", "SERVIDOR", tRs.Fields("SERVIDOR")
            SaveSetting "APTONER", "ConfigSACC", "COMPRAS_NAC_INT", tRs.Fields("COMPRAS_NAC_INT")
            MsgBox "Cambio a la empresa " & sBaseDatos & ", se reiniciara la aplicación", vbInformation, "SACC"
        Else
            MsgBox "No cuenta con permisos de acceso a esta empresa!", vbExclamation, "SACC"
        End If
        Unload Me
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & ". " & Err.Description & ".", vbCritical, "SACC"
    Err.Clear
    Unload Me
End Sub
Private Sub SubMenEmpatarFolios_Click()
    FrmEmpataFolios.Show vbModal
End Sub
Private Sub SubMenEntradaRapida_Click()
    FrmRepEntradasRapidas.Show vbModal
End Sub
Private Sub SubMenEntrOCR_Click()
    FrmEntradaOrdenRapida.Show vbModal
End Sub
Private Sub SubMenEntrOrdenProd_Click()
    FrmEntradaOrdernProd.Show vbModal
End Sub
Private Sub SubMenExploBD_Click()
    FrmVisorTablas.Show vbModal
End Sub
Private Sub SubMenExportaConta_Click()
    FrmExportaConta.Show vbModal
End Sub
Private Sub SubMenFaltantes_Click()
    Faltantes.Show vbModal
End Sub
Private Sub SubMenGastos_Click()
    FrmRepGastos.Show vbModal
End Sub
Private Sub SubMenGraficasTickets_Click()
    FrmGraficasTickets.Show vbModal
End Sub
Private Sub SubMenHacerPedido_Click()
    Pedidos.Show vbModal
End Sub
Private Sub SubMenHacerRequisicion_Click()
    frmRequisicion.Show vbModal
End Sub
Private Sub SubMenImprimirOrdendeCompra_Click()
    frmOrdenCompra.Show vbModal
End Sub
Private Sub SubMenIncobrable_Click()
    FrmIncobrables.Show vbModal
End Sub
Private Sub SubMenInventarios_Click()
    frmInventarios2.Show vbModal
End Sub
Private Sub SubMenLicitacion_Click()
    FrmLicitacion.Show vbModal
End Sub
Private Sub SubMenLlegadadeOrdendeCompra_Click()
    Ordenes.Show vbModal
End Sub
Private Sub SubMenMaterialExtra_Click()
    frmSalidaInvProd.Show vbModal
End Sub
Private Sub SubMenMaxMinAlma3_Click()
    FrmMaxMinAlma3.Show vbModal
End Sub
Private Sub SubMenModOR_Click()
    FrmModificaOrdenRapida.Show vbModal
End Sub
Private Sub SubMenMonitoreo_Click()
    FrmMonitoreo.Show vbModal
End Sub
Private Sub SubMenNotaCredito_Click()
    NotaCredito.Show vbModal
End Sub
Private Sub SubMenNotadeCredito_Click()
    NotaCredito.Show vbModal
End Sub
Private Sub SubMenNuevaEmpresa_Click()
    FrmNuevaEmpresa.Show vbModal
End Sub
Private Sub SubMenNuevaMateriaPrima_Click()
    FrmAltaProdAlm1y2.Frame8.Visible = True
    FrmAltaProdAlm1y2.Frame16.Visible = False
    FrmAltaProdAlm1y2.Caption = "Alta de Productos de Almacn 1 / 2"
    FrmAltaProdAlm1y2.Show vbModal
End Sub
Private Sub SubMenNuevaSucursal_Click()
    AltaSucu.Show vbModal
End Sub
Private Sub SubMenNuevo_Click()
    FrmNuevoJR.Show vbModal
End Sub
Private Sub SubMenNuevoAgente_Click()
    frmPermisos.Show vbModal
End Sub
Private Sub SubMenNuevoCliente_Click()
    AltaClien.Show vbModal
End Sub
Private Sub SubMenNuevoMensajero_Click()
    FrmNueRep.Show vbModal
End Sub
Private Sub SubMenNuevoPRoducto_Click()
    FrmAltaProdAlm3.Frame16.Visible = False
    FrmAltaProdAlm3.Frame8.Visible = True
    FrmAltaProdAlm3.Caption = "Alta de Producto de Almacen 3"
    FrmAltaProdAlm3.Show vbModal
End Sub
Private Sub SubMenNuevoProveedor_Click()
    Proveedor.Show vbModal
End Sub
Private Sub SubMenNvoDpto_Click()
    FrmDepartamentos.Show vbModal
End Sub
Private Sub SubMenOrdendeCompra_Click()
    FrmOrdenRapida.Show vbModal
End Sub
Private Sub SubMenOrdenes_Click()
    frmOrdenesProduccion.Show vbModal
End Sub
Private Sub SubMenOrdenMaxMin_Click()
    FrmProduccionMaxMin.Show vbModal
End Sub
Private Sub SubMenPagarOrdendeCompra_Click()
    FrmPagosOrdenes.Show vbModal
End Sub
Private Sub SubMenPagodeCompraAlmacen1_Click()
    FrmPagoCompAlm1.Show vbModal
End Sub
Private Sub SubMenPagodeOrdenRapida_Click()
    FrmPagoOrdenRapida.Show vbModal
End Sub
Private Sub SubMenPermisos_Click()
    DarPerVenta.Show vbModal
End Sub
Private Sub SubmenPreordendecompra_Click()
    frmPreOrden.Show vbModal
End Sub
Private Sub SubMenPresta_Click()
    FrmPrestamos.Show vbModal
End Sub
Private Sub SubMenPrestamos_Click()
    FrmPrestamos.Show vbModal
End Sub
Private Sub SubMenProdPeriodo_Click()
    FrmRepConsumibles.Show vbModal
End Sub
Private Sub SubMenProduccion_Click()
    frmProduccion.Show vbModal
End Sub
Private Sub SubMenProduccionesporFechas_Click()
    FrmRepComProd.Show vbModal
End Sub
Private Sub SubMenPRoducir_Click()
    FrmCreaExis.Show vbModal
End Sub
Private Sub SubMenProductos_Click()
    BuscaProd.Show vbModal
End Sub
Private Sub SubMenProductosPedidos_Click()
    FrmBusProdPed.Show vbModal
End Sub
Private Sub SubMenProductosPendientesLicitacion_Click()
    FrmVerComprasMaxMinLic.Show vbModal
End Sub
Private Sub SubMenPromocion_Click()
    frmPromos.Show vbModal
End Sub
Private Sub SubMenProveedores_Click()
    frmProveedores.Show vbModal
End Sub
Private Sub SubMenPuntodeVenta_Click()
    Ventas.Show vbModal
End Sub
Private Sub SubMenRastrear_Click()
    FrmRastrearPed.Show vbModal
End Sub
Private Sub SubMenReemplazar_Click()
    FrmSustiInv.Show vbModal
End Sub
Private Sub SubMenReemplazarInsumo_Click()
    EditarJRVarios.Show vbModal
End Sub
Private Sub SubMenReimprimir_Click()
    FrmReImprime.Show vbModal
End Sub
Private Sub SubMenRendimiento_Click()
    FrmRendimiento.Show vbModal
End Sub
Private Sub SubMenRepAbonos_Click()
    FrmRepAbonos.Show vbModal
End Sub
Private Sub SubMenRepComProvVar_Click()
    FrmRepCartVac.Show vbModal
End Sub
Private Sub SubMenRepCxCD_Click()
    FrmRepCXC.Show vbModal
End Sub
Private Sub SubMenRepEntradas_Click()
    FrmRepEntradas.Show vbModal
End Sub
Private Sub SubMenRepEstOc_Click()
    FrmRepOrdenCompra.Show vbModal
End Sub
Private Sub SubMenRepInventarios_Click()
    FrmInventarios.Show vbModal
End Sub
Private Sub SubMenRepMovimientos_Click()
    FrmRepMovimientos.Show vbModal
End Sub
Private Sub SubMenRepOCR_Click()
    FrmRepOrdenRapida.Show vbModal
End Sub
Private Sub SubMenRepOrdenCancel_Click()
    FrmRepOrdenCancel.Show vbModal
End Sub
Private Sub SubMenRepOrdenPa_Click()
    FrmRepPagos.Show vbModal
End Sub
Private Sub SubMenReporteador_Click()
    Reportes1.Show vbModal
End Sub
Private Sub SubMenReportedeComprasProveedor_Click()
    FrmRepCompras.Show vbModal
End Sub
Private Sub SubMenReportedeLicitaciones_Click()
    'FrmProdLic.Show vbModal
    FrmRepLicitaciones.Show vbModal
End Sub
Private Sub SubMenReporteExistencias_Click()
    FrmExisReales.Show vbModal
End Sub
Private Sub SubMenReportesdeComprasdeClientes_Click()
    FrmProdMasVend.Show vbModal
End Sub
Private Sub SubMenReportesdeVentas_Click()
    FrmRepVentas.Show vbModal
End Sub
Private Sub SubMenRepPrestamos_Click()
    FrmRepPrestamosClientes.Show vbModal
End Sub
Private Sub SubMenRepPrestAvtFijo_Click()
    FrmRepPrestamoActivo.Show vbModal
End Sub
Private Sub SubMenRepVenProg_Click()
    FrmRepVentasProgCerradas.Show vbModal
End Sub
Private Sub SubMenRequisicion_Click()
    frmRequisiciones.Show vbModal
End Sub
Private Sub SubMenRevisar_Click()
    frmVerCotizaciones.Show vbModal
End Sub
Private Sub SubMenRevisarCompraenAlmacen1_Click()
    FrmAcepAlmacen1.Show vbModal
End Sub
Private Sub SubMenRevision_Click()
    frmReviComa.Show vbModal
End Sub
Private Sub SubMensajeros_Click()
    FrmRevDomi.Show vbModal
End Sub
Private Sub SubMenSalidas_Click()
    Salidas.Show vbModal
End Sub
Private Sub SubMenSancion_Click()
    FrmSanciones.Show vbModal
End Sub
Private Sub SubMenScrap_Click()
    frmScrap.Show vbModal
End Sub
Private Sub SubMenSESENKA_Click()
    Dim sBaseDatos As String
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBaseDatos = "SESENKA"
    SaveSetting "APTONER", "ConfigSACC", "SERVIDOR", GetSetting("APTONER", "ConfigSACC", "SERVIDOR1", "")
    NvoMen.txtServidor.Text = GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "")
    SaveSetting "APTONER", "ConfigSACC", "SERVIDOR", "SERVER2"
    Set cnn1 = New ADODB.Connection
    With cnn1
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & sBaseDatos & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    sBuscar = "SELECT NOMBRE FROM USUARIOS WHERE NOMBRE = '" & VarMen.Text1(1).Text & "' AND ESTADO = 'A'"
    Set tRs = cnn1.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        SaveSetting "APTONER", "ConfigSACC", "DATABASE", sBaseDatos
        MsgBox "Cambio a la empresa " & sBaseDatos & ", se reiniciara la aplicación", vbInformation, "SACC"
        Unload Me
    Else
        MsgBox "No cuenta con permisos de acceso a esta empresa!", vbExclamation, "SACC"
    End If
End Sub
Private Sub SubMenSurtirVentaProgramada_Click()
    frmShowPediC.Show vbModal
End Sub
Private Sub SubMenTipodeCambio_Click()
    Dolar.Show vbModal
End Sub
Private Sub SubMenTramitarGarantia_Click()
    frmGarantias.Show vbModal
End Sub
Private Sub SubMenTraspasosEntreSucursales_Click()
    Transfe.Show vbModal
End Sub
Private Sub SubMenUsoActFijo_Click()
    FrmUsoActivo.Show vbModal
End Sub
Private Sub SubMenValedeCaja_Click()
    FrmValeCajaCerrar.Show vbModal
End Sub
Private Sub SubMenValedeCajaVentas_Click()
    FrmValeCaja.Show vbModal
End Sub
Private Sub SubMenVentasCanceladas_Click()
    FrmVerVentasCanceladas.Show vbModal
End Sub
Private Sub SubMenVentasEspeciales_Click()
    PermisoVenta.Show vbModal
End Sub
Private Sub SubMenVentasporUsuario_Click()
    FrmComiciones.Show vbModal
End Sub
Private Sub SubMenVentasProgramadas_Click()
    Programadas.Show vbModal
End Sub
Private Sub SubMenVentCan_Click()
    FrmRepNotasCanceladas.Show vbModal
End Sub
Private Sub SubMenVer_Click()
    VerJuegoRep.Show vbModal
End Sub
Private Sub SubMenVerComandasPendientes_Click()
    FrmComPend.Show vbModal
End Sub
Private Sub SubMenVerEntradas_Click()
    BuscaEntrada.Show vbModal
End Sub
Private Sub SubMenVerExistencias_Click()
    BuscaExist.Show vbModal
End Sub
Private Sub SubMenVerGanancias_Click()
    VentasCostos.Show vbModal
End Sub
Private Sub SubSalir_Click()
    Unload Me
End Sub
Private Sub SubSoporteTecnico_Click()
    frmAStec.Show vbModal
End Sub
Private Sub SumMenCortedeCaja_Click()
    FrmCorteCredito.Show vbModal
End Sub
Sub Sincronizar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Me.lblEstado.Caption = "Sincronizando con los servidores, espere..."
    Me.lblEstado.ForeColor = vbRed
    'DoEvents
    sBuscar = "SELECT DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())) AS FECHAHORA"
    Set tRs = cnn.Execute(sBuscar)
    Time = TimeValue(tRs.Fields("FECHAHORA"))
    Date = DateValue(tRs.Fields("FECHAHORA"))
    Me.lblEstado.Caption = ""
    'DoEvents
    Exit Sub
ManejaError:
    Err.Clear
End Sub
Private Sub ValidaMenu()
    If Text1(12).Text = "N" And Text1(56).Text = "N" And Text1(6).Text = "N" And Text1(7).Text = "N" Then 'OK
        SubMenPuntodeVenta.Enabled = False
        SubMenConsCobCom.Enabled = False
    End If
    If Text1(13).Text = "N" Then
        SumMenCortedeCaja.Enabled = False
    End If
    If Text1(49).Text = "N" Then
        SubMenVentasProgramadas.Enabled = False
    End If
    If Text1(54).Text = "N" Then
        SubMenMaxMinAlma3.Enabled = False
    End If
    If Text1(55).Text = "N" Then
        SubMenFaltantes.Enabled = False
    End If
    If Text1(57).Text = "N" Then
        SubMenAbonaraCreditos.Enabled = False
        SubMenIncobrable.Enabled = False
    End If
    If Text1(48).Text = "N" Then
        SubMenVentasEspeciales.Enabled = False
    End If
    If Text1(10).Text = "N" Then
        SubMenTramitarGarantia.Enabled = False
    End If
    If Text1(11).Text = "N" Then
        SubMenCancelaciones.Enabled = False
        SubMenCanComand.Enabled = False
        SubMenAutAltaCliente.Enabled = False
        SubMenPermisos.Enabled = False
        SubMenLicitacion.Enabled = False
        SubMenCambiarVentadeSucursal.Enabled = False
        ImportarPrecios.Enabled = False
    End If
    If Text1(18).Text = "N" Then
        SubMenNotaCredito.Enabled = False
    End If
    If Text1(9).Text = "N" Then
        SubMenNotadeCredito.Enabled = False
        SubMenValedeCaja.Enabled = False
        SubMenValedeCajaVentas.Enabled = False
    End If
    If Text1(64).Text = "N" Then
        SubMenAutorizarGarantia.Enabled = False
        SubMenAutorizarRemanufactura.Enabled = False
    End If
    If Text1(27).Text = "N" Then
        SubMenCompraenAlmacen1.Enabled = False
    End If
    If Text1(17).Text = "N" Then
        SubMenHacerRequisicion.Enabled = False
        SubMenOrdenes.Enabled = False
    End If
    If Text1(51).Text = "N" Then
        SubMenRastrear.Enabled = False
    End If
    'If Text1(52).Text = "N" Then
    '    SubMenProductosPedidos.Enabled = False
    'End If
    If Text1(15).Text = "N" Then
        SubMenVerExistencias.Enabled = False
    End If
    If Text1(14).Text = "N" Then
        SubMenProductos.Enabled = False
    End If
    If Text1(56).Text = "N" Then
        SubMensajeros.Enabled = False
    End If
    If Text1(24).Text = "N" Then 'OK
        SubMenVerEntradas.Enabled = False
    End If
    If Text1(26).Text = "N" Then
        SubMenInventarios.Enabled = False
    End If
    If Text1(74).Text = "N" Then
        SubMenVentasCanceladas.Enabled = False
    End If
    If Text1(53).Text = "N" Then
        SubMenSurtirVentaProgramada.Enabled = False
    End If
    If Text1(19).Text = "N" Then
        SubMenLlegadadeOrdendeCompra.Enabled = False
        SubMenEntrOCR.Enabled = False
        SubMenEntrOrdenProd.Enabled = False
    End If
    If Text1(20).Text = "N" Then
        SubMenCanAbono.Enabled = False
    End If
    If Text1(23).Text = "N" Then
        SubMenTraspasosEntreSucursales.Enabled = False
    End If
    If Text1(63).Text = "N" Then
        SubMenRequisicion.Enabled = False
    End If
    If Text1(50).Text = "N" Then
        SubMenModOR.Enabled = False
    End If
    If Text1(59).Text = "N" Then
        SubMenRevisar.Enabled = False
    End If
    If Text1(60).Text = "N" Then
        SubMenAsignar.Enabled = False
    End If
    If Text1(61).Text = "N" Then
        SubmenPreordendecompra.Enabled = False
    End If
    If Text1(62).Text = "N" Then
        SubMenImprimirOrdendeCompra.Enabled = False
    End If
    If Text1(28).Text = "N" Then
        SubSoporteTecnico.Enabled = False
    End If
    If Text1(39).Text = "N" Then
        SubMenRevision.Enabled = False
        SubMenOrdenMaxMin.Enabled = False
        PedMaxMin.Enabled = False
    End If
    If Text1(40).Text = "N" Then
        SubMenProduccion.Enabled = False
    End If
    If Text1(41).Text = "N" Then
        SubMenCalidad.Enabled = False
    End If
    If Text1(42).Text = "N" Then
        SubMenVer.Enabled = False
    End If
    If Text1(36).Text = "N" Then
        SubMenSancion.Enabled = False
    End If
    If Text1(47).Text = "N" Then
        SubMenDatosdelaEmpresa.Enabled = False
        SubMenBasedeDatos.Enabled = False
        SubMenExploBD.Enabled = False
        'SubMenExpo.Enabled = False
        subfrmfrmaltaeua.Enabled = False
        SubMenMonitoreo.Enabled = False
        SubMenNuevaEmpresa.Enabled = False
    End If
    If Text1(8).Text = "N" Then
        SubMenCancelarOrdenRapida.Enabled = False
        SubMenCancelarCompraAlmacen1.Enabled = False
    End If
    If Text1(29).Text = "N" Then
        SubMenNvoDpto.Enabled = False
        SubMenNuevoAgente.Enabled = False
        SubMenNuevoMensajero.Enabled = False
    End If
    If Text1(30).Text = "N" Then
        SubMenNuevoCliente.Enabled = False
    End If
    If Text1(31).Text = "N" Then
        SubMenNuevaSucursal.Enabled = False
    End If
    If Text1(32).Text = "N" Then
        SubMenNuevoProveedor.Enabled = False
        SumMenProvRapi.Enabled = False
    End If
    If Text1(22).Text = "N" Then
        SubMenNuevoPRoducto.Enabled = False
    End If
    If Text1(37).Text = "N" Then
        SubMenEliminarProducto.Enabled = False
    End If
    If Text1(21).Text = "N" Then
        SubMenNuevaMateriaPrima.Enabled = False
    End If
    If Text1(44).Text = "N" Then
        SubMenEliminarMateriaPrima.Enabled = False
    End If
    If Text1(33).Text = "N" Then
        SubMenEliminarAgente.Enabled = False
        SubMenEliminarMensajero.Enabled = False
    End If
    If Text1(34).Text = "N" Then
        SubMenEliminarCliente.Enabled = False
    End If
    If Text1(58).Text = "N" Then
        SubMenEliminarSucursal.Enabled = False
    End If
    If Text1(43).Text = "N" Then
        SubMenPromocion.Enabled = False
    End If
    If Text1(45).Text = "N" Then
        SubMenTipodeCambio.Enabled = False
    End If
    If Text1(16).Text = "N" Then
        SubMenHacerPedido.Enabled = False
    End If
    If Text1(38).Text = "N" Then
        SubMenNuevo.Enabled = False
    End If
    If Text1(65).Text = "N" Then
        SubMenReemplazarInsumo.Enabled = False
    End If
    If Text1(66).Text = "N" Then
        SubMenMaterialExtra.Enabled = False
        SubMenScrap.Enabled = False
    End If
    If Text1(67).Text = "N" Then
        SubMenCanRefa.Enabled = False
    End If
    If Text1(68).Text = "N" Then
        'SubMenCompraenAlmacen1.Enabled = False
    End If
    If Text1(69).Text = "N" Then
        SubMenRevisarCompraenAlmacen1.Enabled = False
    End If
    If Text1(70).Text = "N" Then
        SubMenAlmacen1.Enabled = False
    End If
    If Text1(71).Text = "N" Then
        SubMenPagodeCompraAlmacen1.Enabled = False
        SubMenPagarOrdendeCompra.Enabled = False
        SubMenPagodeOrdenRapida.Enabled = False
    End If
    If VarMen.Text1(78).Text = "N" Then
        'REGRESAR ORDEN DE COMPRA
    End If
    If Text1(72).Text = "N" Then
        SubMenPRoducir.Enabled = False
        SubMenReemplazar.Enabled = False
    End If
    If Text1(73).Text = "N" Then
        SubMenPrestamos.Enabled = False
    End If
    If Text1(76).Text = "N" Then
        SubMenEliminarProveedor.Enabled = False
        SubMenEliminarSucursal.Enabled = False
    End If
    If Text1(77).Text = "N" Then
        SubMenAutorizar.Enabled = False
    End If
    If Text1(25).Text = "N" Then
        SubMenOrdendeCompra.Enabled = False
    End If
    If Text1(46).Text = "N" Then
        SubMenSalidas.Enabled = False
    End If
End Sub
Private Sub SumMenProvRapi_Click()
    FrmProvConsumibles.Show vbModal
End Sub
Private Sub Timer1_Timer()
    If Image2.Visible Then
        Image2.Visible = False
    Else
        Image2.Visible = True
    End If
End Sub
Private Sub Timer2_Timer()
    If Image3.Visible Then
        Image3.Visible = False
        Image27.Visible = True
    Else
        Image3.Visible = True
        Image27.Visible = False
    End If
End Sub
