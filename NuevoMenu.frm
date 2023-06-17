VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form NuevoMenu 
   BackColor       =   &H00C0C0C0&
   Caption         =   "SACC ( Sistema de Administración y Control del Comercio )"
   ClientHeight    =   8715
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12195
   Icon            =   "NuevoMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   15195
      TabIndex        =   96
      Top             =   0
      Width           =   15255
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   11040
         TabIndex        =   97
         Top             =   0
         Width           =   975
         Begin VB.Image imgLeer 
            Height          =   630
            Left            =   120
            MouseIcon       =   "NuevoMenu.frx":1601A
            MousePointer    =   99  'Custom
            Picture         =   "NuevoMenu.frx":16324
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Mensajero"
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
            TabIndex        =   98
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   73
      Left            =   1080
      TabIndex        =   92
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   72
      Left            =   960
      TabIndex        =   91
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   71
      Left            =   840
      TabIndex        =   90
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   70
      Left            =   720
      TabIndex        =   89
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   69
      Left            =   600
      TabIndex        =   88
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   68
      Left            =   480
      TabIndex        =   87
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   67
      Left            =   360
      TabIndex        =   86
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   66
      Left            =   1440
      TabIndex        =   85
      Top             =   8040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   65
      Left            =   2520
      TabIndex        =   84
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   64
      Left            =   2400
      TabIndex        =   83
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   63
      Left            =   2280
      TabIndex        =   82
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   62
      Left            =   2160
      TabIndex        =   81
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   61
      Left            =   2040
      TabIndex        =   80
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   60
      Left            =   1920
      TabIndex        =   79
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   59
      Left            =   1800
      TabIndex        =   78
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   58
      Left            =   1680
      TabIndex        =   77
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   57
      Left            =   1560
      TabIndex        =   76
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   56
      Left            =   1440
      TabIndex        =   75
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   55
      Left            =   2520
      TabIndex        =   74
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   54
      Left            =   2400
      TabIndex        =   73
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   53
      Left            =   2280
      TabIndex        =   72
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   52
      Left            =   2160
      TabIndex        =   71
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   51
      Left            =   2040
      TabIndex        =   70
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   50
      Left            =   1920
      TabIndex        =   69
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   49
      Left            =   1800
      TabIndex        =   68
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   48
      Left            =   1680
      TabIndex        =   67
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   47
      Left            =   1560
      TabIndex        =   66
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   46
      Left            =   1440
      TabIndex        =   65
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   45
      Left            =   2520
      TabIndex        =   64
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   44
      Left            =   2400
      TabIndex        =   63
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   43
      Left            =   2280
      TabIndex        =   62
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   42
      Left            =   2160
      TabIndex        =   61
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   41
      Left            =   2040
      TabIndex        =   60
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   40
      Left            =   1920
      TabIndex        =   59
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   39
      Left            =   1800
      TabIndex        =   58
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   38
      Left            =   1680
      TabIndex        =   57
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   37
      Left            =   1560
      TabIndex        =   56
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   36
      Left            =   1440
      TabIndex        =   55
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   35
      Left            =   2520
      TabIndex        =   54
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   34
      Left            =   2400
      TabIndex        =   53
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   33
      Left            =   2280
      TabIndex        =   52
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   32
      Left            =   2160
      TabIndex        =   51
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   31
      Left            =   2040
      TabIndex        =   50
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   30
      Left            =   1920
      TabIndex        =   49
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   29
      Left            =   1800
      TabIndex        =   48
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   28
      Left            =   1680
      TabIndex        =   47
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   27
      Left            =   1560
      TabIndex        =   46
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   1440
      TabIndex        =   45
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   25
      Left            =   2520
      TabIndex        =   44
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   24
      Left            =   2400
      TabIndex        =   43
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   2280
      TabIndex        =   42
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   2160
      TabIndex        =   41
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   21
      Left            =   2040
      TabIndex        =   40
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   1920
      TabIndex        =   39
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   1800
      TabIndex        =   38
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   1680
      TabIndex        =   37
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   1560
      TabIndex        =   36
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   1440
      TabIndex        =   35
      Top             =   6840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   2520
      TabIndex        =   34
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   2400
      TabIndex        =   33
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   2280
      TabIndex        =   32
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   2160
      TabIndex        =   31
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   30
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   1920
      TabIndex        =   29
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   1800
      TabIndex        =   28
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1680
      TabIndex        =   27
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1560
      TabIndex        =   26
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1440
      TabIndex        =   25
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   24
      Text            =   "Text4"
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   23
      Text            =   "Text4"
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   22
      Text            =   "Text4"
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   21
      Text            =   "Text4"
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   20
      Text            =   "Text4"
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   19
      Text            =   "Text4"
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   720
      TabIndex        =   18
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   17
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   16
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   6
      Left            =   960
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   5
      Left            =   720
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   6
      Left            =   840
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   7
      Left            =   960
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   9
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   6600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtServidor 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   150
   End
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   840
      Top             =   6960
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   25
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
      Bmp:1           =   "NuevoMenu.frx":17CFE
      Mask:1          =   16777215
      Key:1           =   "#SubSistema"
      Bmp:2           =   "NuevoMenu.frx":182A0
      Mask:2          =   16777215
      Key:2           =   "#SubEliminar"
      Bmp:3           =   "NuevoMenu.frx":185F2
      Mask:3          =   16777215
      Key:3           =   "#SubJuegosdeReparacion"
      Bmp:4           =   "NuevoMenu.frx":18AF4
      Mask:4          =   16777215
      Key:4           =   "#SubProcesos"
      Bmp:5           =   "NuevoMenu.frx":18E46
      Mask:5          =   16777215
      Key:5           =   "#SubRevision"
      Bmp:6           =   "NuevoMenu.frx":19198
      Mask:6          =   16777215
      Key:6           =   "#SubBloquear"
      Bmp:7           =   "NuevoMenu.frx":196D6
      Mask:7          =   16777215
      Key:7           =   "#SubCerrarSesion"
      Bmp:8           =   "NuevoMenu.frx":19C14
      Mask:8          =   16777215
      Key:8           =   "#SubAtención"
      Bmp:9           =   "NuevoMenu.frx":1A116
      Mask:9          =   16777215
      Key:9           =   "#SubAdministración"
      Bmp:10          =   "NuevoMenu.frx":1A618
      Mask:10         =   16777215
      Key:10          =   "#SubSalir"
      Bmp:11          =   "NuevoMenu.frx":1AAA2
      Mask:11         =   16777215
      Key:11          =   "#SubVentas"
      Bmp:12          =   "NuevoMenu.frx":1AFA4
      Mask:12         =   16777215
      Key:12          =   "#SubNuevo"
      Bmp:13          =   "NuevoMenu.frx":1B2F6
      Mask:13         =   16777215
      Key:13          =   "#SubCotizar"
      Bmp:14          =   "NuevoMenu.frx":1B7F8
      Mask:14         =   16777215
      Key:14          =   "#SubOrdenesdeCompra"
      Bmp:15          =   "NuevoMenu.frx":1BCFA
      Mask:15         =   16777215
      Key:15          =   "#SubMateriaPrima"
      Bmp:16          =   "NuevoMenu.frx":1C184
      Mask:16         =   16777215
      Key:16          =   "#SubPedidosAlmacen"
      Bmp:17          =   "NuevoMenu.frx":1C596
      Mask:17         =   16777215
      Key:17          =   "#SubMovimientos"
      Bmp:18          =   "NuevoMenu.frx":1CC04
      Mask:18         =   16777215
      Key:18          =   "#SubEntradas"
      Bmp:19          =   "NuevoMenu.frx":1D1EA
      Mask:19         =   16777215
      Key:19          =   "#SubRevisiones"
      Bmp:20          =   "NuevoMenu.frx":1D6EC
      Mask:20         =   16777215
      Key:20          =   "#SubSoporteTecnico"
      Bmp:21          =   "NuevoMenu.frx":1DB76
      Mask:21         =   16777215
      Key:21          =   "#SubMensajeros"
      Bmp:22          =   "NuevoMenu.frx":1E048
      Mask:22         =   16777215
      Key:22          =   "#SubContabilidad"
      Bmp:23          =   "NuevoMenu.frx":1E59A
      Mask:23         =   16777215
      Key:23          =   "#SubConsultas"
      Bmp:24          =   "NuevoMenu.frx":1EA4C
      Mask:24         =   16777215
      Key:24          =   "#SubReportes"
      Bmp:25          =   "NuevoMenu.frx":1F03E
      Mask:25         =   16777215
      Key:25          =   "#SubPedir"
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
      Top             =   8340
      Width           =   12195
      _ExtentX        =   21511
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
            TextSave        =   "04:02 p.m."
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   0
            TextSave        =   "23/08/2007"
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
            Text            =   "Versión 2.3.0"
            TextSave        =   "Versión 2.3.0"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPuestoSucursal 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   600
      TabIndex        =   95
      Top             =   2520
      Width           =   9495
   End
   Begin VB.Image Image2 
      Height          =   1530
      Left            =   3480
      Picture         =   "NuevoMenu.frx":1F480
      Top             =   3120
      Width           =   4935
   End
   Begin VB.Label lblHola 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   94
      Top             =   1440
      Width           =   8175
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1560
      TabIndex        =   93
      Top             =   4920
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   8400
      Picture         =   "NuevoMenu.frx":37E6A
      Top             =   6360
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
         Begin VB.Menu SubMenFacturar 
            Caption         =   "Facturar"
         End
         Begin VB.Menu SubMenExistencias 
            Caption         =   "Existencias"
         End
      End
      Begin VB.Menu SubAtención 
         Caption         =   "Atención"
         Begin VB.Menu SubMenTramitarGarantia 
            Caption         =   "Tramitar Garantia"
         End
         Begin VB.Menu SubMenAutorizarGarantia 
            Caption         =   "Autorizar Garantia"
         End
         Begin VB.Menu SubMenAutorizarRemanufactura 
            Caption         =   "Autorizar Remanufactura"
         End
         Begin VB.Menu SubMenPrestamos 
            Caption         =   "Prestamos"
         End
      End
      Begin VB.Menu SubAdministración 
         Caption         =   "Administración"
         Begin VB.Menu SumMenCortedeCaja 
            Caption         =   "Corte de Caja"
         End
         Begin VB.Menu SubMenLicitacion 
            Caption         =   "Licitación"
         End
         Begin VB.Menu SubMenPromocion 
            Caption         =   "Promoción"
         End
         Begin VB.Menu SubMenCancelaciones 
            Caption         =   "Cancelaciones"
         End
         Begin VB.Menu SubMenPermisos 
            Caption         =   "Permisos"
         End
         Begin VB.Menu SubMenCambiarPrecios 
            Caption         =   "Cambiar Precios"
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
            Caption         =   "Revisar"
         End
         Begin VB.Menu SubMenAsignar 
            Caption         =   "Asignar"
         End
         Begin VB.Menu SubmenPreordendecompra 
            Caption         =   "Preorden de Compra"
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
         Begin VB.Menu SubMenOrdendeCompra 
            Caption         =   "Orden de Compra Rapida"
         End
         Begin VB.Menu SubMenPEndientes 
            Caption         =   "Pendientes"
         End
      End
      Begin VB.Menu SubMateriaPrima 
         Caption         =   "Materia Prima"
         Begin VB.Menu SubMenCompraenAlmacen1 
            Caption         =   "Compra en Almacen 1"
         End
         Begin VB.Menu SubMenRevisarCompraenAlmacen1 
            Caption         =   "Revisar Compra en Almacen 1"
         End
         Begin VB.Menu SubMenAlmacen1 
            Caption         =   "Aprovar Compra en Almacen 1"
         End
      End
   End
   Begin VB.Menu MenAlmacen 
      Caption         =   "&Almacen"
      Begin VB.Menu SubPedidosAlmacen 
         Caption         =   "Pedidos"
         Begin VB.Menu SubMenVerPedidos 
            Caption         =   "Ver Peridos"
         End
         Begin VB.Menu SubMenHacerRequisicion 
            Caption         =   "Hacer Requisición"
         End
      End
      Begin VB.Menu SubMovimientos 
         Caption         =   "Movimientos"
         Begin VB.Menu SubMenInventarios 
            Caption         =   "Inventarios"
         End
         Begin VB.Menu SubMenSalidas 
            Caption         =   "Salidas"
         End
         Begin VB.Menu SubMenSurtirSucursal 
            Caption         =   "Surtir Sucursal"
         End
         Begin VB.Menu SubMenSurtirVentaProgramada 
            Caption         =   "Surtir Venta Programada"
         End
         Begin VB.Menu SubMenAjustedeVenta 
            Caption         =   "Ajuste de Venta"
         End
         Begin VB.Menu SubMenPRoducir 
            Caption         =   "Producir"
         End
         Begin VB.Menu SubMenPerdidas 
            Caption         =   "Perdidas"
         End
         Begin VB.Menu SubMenReemplazar 
            Caption         =   "Reemplazar"
         End
         Begin VB.Menu SubMenPrestamoInternodeAlmacen1 
            Caption         =   "Prestamo Interno de Almacen1"
         End
      End
      Begin VB.Menu SubEntradas 
         Caption         =   "Entradas"
         Begin VB.Menu SubMenLlegadadeOrdendeCompra 
            Caption         =   "Llegada de Orden de Compra"
         End
         Begin VB.Menu SubMenTraspasosEntreSucursales 
            Caption         =   "Traspasos Entre Sucursales"
         End
         Begin VB.Menu SubMenEntradasaAlmacenes 
            Caption         =   "Entradas a Almacenes"
         End
         Begin VB.Menu SubMenPEndientesdeEntrega 
            Caption         =   "Pendientes de Entrega"
         End
         Begin VB.Menu SubMenCerrarPendientesdeEntrega 
            Caption         =   "Cerrar Pendientes de Entrega"
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
         Begin VB.Menu SubMenScrap 
            Caption         =   "Scrap"
         End
         Begin VB.Menu SubMenMaterialExtra 
            Caption         =   "Material Extra"
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
      Caption         =   "&Administración"
      Begin VB.Menu SubNuevo 
         Caption         =   "Agregar"
         Begin VB.Menu SubMenNuevoAgente 
            Caption         =   "Nuevo Agente"
         End
         Begin VB.Menu SubMenNuevoCliente 
            Caption         =   "Nuevo Cliente"
         End
         Begin VB.Menu SubMenNuevaSucursal 
            Caption         =   "Nueva Sucursal"
         End
         Begin VB.Menu SubMenNuevoProveedor 
            Caption         =   "Nuevo Proveedor"
         End
         Begin VB.Menu SubMenNuevoMensajero 
            Caption         =   "Nuevo Mensajero"
         End
         Begin VB.Menu SubMenNuevoPRoducto 
            Caption         =   "Nuevo Producto"
         End
         Begin VB.Menu SubMenNuevaMateriaPrima 
            Caption         =   "Nueva Meteria Prima"
         End
         Begin VB.Menu SubMenNuevaMarca 
            Caption         =   "Nueva Marca"
         End
      End
      Begin VB.Menu SubEliminar 
         Caption         =   "Eliminar"
         Begin VB.Menu SubMenEliminarAgente 
            Caption         =   "Elmiminar Agente"
         End
         Begin VB.Menu SubMenEliminarCliente 
            Caption         =   "Eliminar Cliente"
         End
         Begin VB.Menu SubMenEliminarSucursal 
            Caption         =   "Eliminar Sucursal"
         End
         Begin VB.Menu SubMenEliminarProveedor 
            Caption         =   "Eliminar Proveedor"
         End
         Begin VB.Menu SubMenEliminarMensajero 
            Caption         =   "Eliminar Mensajero"
         End
         Begin VB.Menu SubMenEliminarProducto 
            Caption         =   "Eliminar Producto"
         End
         Begin VB.Menu SubMenEliminarMateriaPrima 
            Caption         =   "Eliminar Materia Prima"
         End
      End
      Begin VB.Menu SubSistema 
         Caption         =   "Sistema"
         Begin VB.Menu SubMenDatosdelaEmpresa 
            Caption         =   "Datos de la Empresa"
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
         Begin VB.Menu SubMenAbonaraCreditos 
            Caption         =   "Abonar a Creditos"
         End
         Begin VB.Menu SubMenPagodeCompraAlmacen1 
            Caption         =   "Pago de Compra Almacen1"
         End
         Begin VB.Menu SubMenNotadeCredito 
            Caption         =   "Nota de Credito"
         End
         Begin VB.Menu SubMenValedeCaja 
            Caption         =   "Valde de Caja"
         End
         Begin VB.Menu SubMenTipodeCambio 
            Caption         =   "Tipo de Cambio"
         End
      End
   End
   Begin VB.Menu MenUtilerias 
      Caption         =   "&Utilerias"
      Begin VB.Menu SubConsultas 
         Caption         =   "Consultas"
         Begin VB.Menu SubMenProveedores 
            Caption         =   "Proveedores"
         End
         Begin VB.Menu SubMenProductos 
            Caption         =   "Productos"
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
         Begin VB.Menu SubMenProductosPedidos 
            Caption         =   "Productos Pedidos"
         End
         Begin VB.Menu SubMenVerExistencias 
            Caption         =   "Ver Existencias"
         End
      End
      Begin VB.Menu SubPedir 
         Caption         =   "Pedir"
         Begin VB.Menu SubMenHacerPedido 
            Caption         =   "Hacer Pedido"
         End
         Begin VB.Menu SubMenOrdenes 
            Caption         =   "Ordenes"
         End
         Begin VB.Menu SubMenRastrearPedido 
            Caption         =   "Rastrear Pedido"
         End
      End
      Begin VB.Menu SubReportes 
         Caption         =   "Reportes"
         Begin VB.Menu SubMenReporteador 
            Caption         =   "Reporteador"
         End
         Begin VB.Menu SubMenReportesdeVentas 
            Caption         =   "Reportes de Ventas"
         End
         Begin VB.Menu SubMenBajaraEXCEL 
            Caption         =   "Bajar a EXCEL"
         End
         Begin VB.Menu SubMenVentasporUsuario 
            Caption         =   "Ventas por Usuario"
         End
         Begin VB.Menu SubMenImprimirJuegosdeReparacion 
            Caption         =   "Imprimir Juegos de Reparación"
         End
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
Attribute VB_Name = "NuevoMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
' Esta clase se usará para seleccionar el fichero
Dim SubMen As Integer
Dim Validar As Integer
'Tipos, constantes y funciones para FileExist
Const MAX_PATH = 260
Const INVALID_HANDLE_VALUE = -1

Private Type FILETIME
        dwLowDateTime       As Long
        dwHighDateTime      As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes    As Long
        ftCreationTime      As FILETIME
        ftLastAccessTime    As FILETIME
        ftLastWriteTime     As FILETIME
        nFileSizeHigh       As Long
        nFileSizeLow        As Long
        dwReserved0         As Long
        dwReserved1         As Long
        cFileName           As String * MAX_PATH
        cAlternate          As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" _
    (ByVal hFindFile As Long) As Long

'------------------------------------------------------------------------------
' FIN DE CODIGO DE FUNCIONES PARA SABER SI EXISTE O NO UN ARCHIVO
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' Clase para manejar ficheros INIs
' Permite leer secciones enteras y todas las secciones de un fichero INI
'
' Última revisión:                                                  (04/Abr/01)
'
' ©Guillermo 'guille' Som, 1997-2003
'------------------------------------------------------------------------------

Private sBuffer As String   ' Para usarla en las funciones GetSection(s)

'--- Declaraciones para leer ficheros INI ---
' Leer todas las secciones de un fichero INI, esto seguramente no funciona en Win95
' Esta función no estaba en las declaraciones del API que se incluye con el VB
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" _
    (ByVal lpszReturnBuffer As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

' Leer una sección completa
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" _
    (ByVal lpAppName As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long

' Leer una clave de un fichero INI
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
     ByVal lpDefault As String, ByVal lpReturnedString As String, _
     ByVal nSize As Long, ByVal lpFileName As String) As Long


Private Function IniGet(ByVal sFileName As String, ByVal sSection As String, _
                       ByVal sKeyName As String, _
                       Optional ByVal sDefault As String = "") As String
    '--------------------------------------------------------------------------
    ' Devuelve el valor de una clave de un fichero INI
    ' Los parámetros son:
    '   sFileName   El fichero INI
    '   sSection    La sección de la que se quiere leer
    '   sKeyName    Clave
    '   sDefault    Valor opcional que devolverá si no se encuentra la clave
    '--------------------------------------------------------------------------
    Dim ret As Long
    Dim sRetVal As String
    '
    sRetVal = String$(255, 0)
    '
    ret = GetPrivateProfileString(sSection, sKeyName, sDefault, sRetVal, Len(sRetVal), sFileName)
    If ret = 0 Then
        IniGet = sDefault
    Else
        IniGet = Left$(sRetVal, ret)
    End If
End Function


Private Function IniGetSection(ByVal sFileName As String, _
                              ByVal sSection As String) As String()
    '--------------------------------------------------------------------------
    ' Lee una sección entera de un fichero INI                      (27/Feb/99)
    ' Adaptada para devolver un array de string                     (04/Abr/01)
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
    '
    Dim i As Long
    Dim j As Long
    Dim sTmp As String
    Dim sClave As String
    Dim sValor As String
    '
    Dim aSeccion() As String
    Dim n As Long
    '
    ReDim aSeccion(0)
    '
    ' El tamaño máximo para Windows 95
    sBuffer = String$(32767, Chr$(0))
    '
    n = GetPrivateProfileSection(sSection, sBuffer, Len(sBuffer), sFileName)
    '
    If n Then
        '
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
    '--------------------------------------------------------------------------
    ' Devuelve todas las secciones de un fichero INI                (27/Feb/99)
    ' Adaptada para devolver un array de string                     (04/Abr/01)
    '
    ' Esta función devolverá un array con todas las secciones del fichero
    '
    ' Parámetros de entrada:
    '   sFileName   Nombre del fichero INI
    ' Devuelve:
    '   Un array con todos los nombres de las secciones
    '   La primera sección estará en el elemento 1,
    '   por tanto, si el array contiene cero elementos es que no hay secciones
    '
    Dim i As Long
    Dim sTmp As String
    Dim n As Long
    Dim aSections() As String
    '
    ReDim aSections(0)
    '
    ' El tamaño máximo para Windows 95
    sBuffer = String$(32767, Chr$(0))
    '
    ' Esta función del API no está definida en el fichero TXT
    n = GetPrivateProfileSectionNames(sBuffer, Len(sBuffer), sFileName)
    '
    If n Then
        ' Cortar la cadena al número de caracteres devueltos
        sBuffer = Left$(sBuffer, n)
        ' Quitar los vbNullChar extras del final
        i = InStr(sBuffer, vbNullChar & vbNullChar)
        If i Then
            sBuffer = Left$(sBuffer, i - 1)
        End If
        '
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
    '
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
    Dim tRs As Recordset
    Me.lblHola.Caption = "Hola " & Trim(Me.Text1(1).Text) & " " & Trim(Me.Text1(2).Text) & "!"
    Me.lblPuestoSucursal.Caption = Trim(Me.Text1(3).Text) & " en " & Trim(Me.Text4(0).Text)
    Sincronizar
    Me.lblEstado.Caption = "Buscando mensajes"
    Me.lblEstado.ForeColor = vbBlue
    DoEvents
    If Hay_Mensajes(Menu.Text1(0).Text) Then
        Me.Image2.Visible = True
    Else
        Me.Image2.Visible = False
    End If
    
    Me.lblEstado.Caption = ""
    sBuscar = "DELETE FROM EXISTENCIAS WHERE CANTIDAD <= 0"
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
    DoEvents
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & ". " & Err.Description & ".", vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    On Error GoTo ManejaError
    Dim sValor As String
    If FileExist(App.Path & "\Server.Ini") And GetSetting("APTONER", "ConfigSACC", "RegAprovSACC", "0") = "ValAprovReg" Then
        sValor = ""
        txtServidor.Text = IniGet(App.Path & "\Server.Ini", "Servidor", "Nombre", sValor)
        If Hay_Usuarios Then
            frmLogin.Show vbModal, Me
        End If
    Else
        RegSACC.Show vbModal
        Unload Me
    End If
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    Dim Guarda As String
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tRs1 As Recordset
    sBuscar = "SELECT FECHA FROM RESPALDOS_BD WHERE FECHA = '" & Date & "'"
    On Error Resume Next
    Set tRs = cnn.Execute(sBuscar)
    On Error Resume Next
    If tRs.EOF And tRs.BOF Then
        'Respalda la BD
        Guarda = "C:\RespaldoSACC" & Date & ".Bak"
        Guarda = Replace(Guarda, "/", "-")
        sBuscar = "BACKUP DATABASE APTONER TO DISK = '" & Guarda & "' WITH FORMAT,NAME = 'res'"
        cnn.Execute (sBuscar)
        sBuscar = "INSERT INTO RESPALDOS_BD (FECHA, NOMBRE_RESPALDO) VALUES ('" & Date & "', 'RespaldoSACC" & Date & ".Bak')"
        cnn.Execute (sBuscar)
        ' Empareja las CXC en precios mal
        sBuscar = "SELECT ID_CUENTA, DEUDA, TOTAL_COMPRA From vsCxC WHERE(TOTAL_COMPRA <> DEUDA)"
        Set tRs = cnn.Execute(sBuscar)
        Do While Not tRs.EOF
            sBuscar = "UPDATE CUENTAS SET DEUDA = " & tRs.Fields("TOTAL_COMPRA") & " WHERE ID_CUENTA = " & tRs.Fields("ID_CUENTA")
            Set tRs1 = cnn.Execute(sBuscar)
            tRs.MoveNext
        Loop
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & ". " & Err.Description & ".", vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub imgLeer_Click()
    MsgAPToner.Show
End Sub
Private Sub Form_Resize()
    Image1.Left = (NuevoMenu.Width) - (Image1.Width + 400)
    Image1.Top = (NuevoMenu.Height) - (Image1.Height + 1200)
    Image2.Left = (NuevoMenu.Width / 2) - (Image2.Width / 2)
    Image2.Top = (NuevoMenu.Height / 2) - (Image2.Height / 2)
    Frame3.Left = NuevoMenu.Width - (Frame3.Width + 300)
End Sub
Private Sub Image2_Click()
    MsgAPToner.Show
End Sub
Private Sub SubBloquear_Click()
    frmLogin.Show vbModal, Me
End Sub
Private Sub SubCerrarSesion_Click()
    Unload Me
    Menu.Show
End Sub
Private Sub SubMenAjustedeVenta_Click()
    PermisoAjuste.Show vbModal
End Sub
Private Sub SubMenAlmacen1_Click()
    FrmArpvCompAlm1.Show vbModal
End Sub
Private Sub SubMenAsignar_Click()
    frmAutorizarCotizaciones.Show vbModal
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
Private Sub SubMenCalidad_Click()
    frmCalidad.Show vbModal
End Sub
Private Sub SubMenCambiarPrecios_Click()
    CambioPRe.Show vbModal
End Sub
Private Sub SubMenCancelaciones_Click()
    frmCancelaFactura.Show vbModal
End Sub
Private Sub SubMenCerrarPendientesdeEntrega_Click()
    frmSalidaExistenciasTemporales.Show vbModal
End Sub
Private Sub SubMenCompraenAlmacen1_Click()
    FrmCompAlm1.Show vbModal
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
    FrmEliProdAlm1y2.Show vbModal
End Sub
Private Sub SubMenEliminarMensajero_Click()
    EliMensajero.Show vbModal
End Sub
Private Sub SubMenEliminarProducto_Click()
    FrmEliProdAlm3.Show vbModal
End Sub
Private Sub SubMenEliminarProveedor_Click()
    EliProveedor.Show vbModal
End Sub
Private Sub SubMenEliminarSucursal_Click()
    EliSuc.Show vbModal
End Sub
Private Sub SubMenEntradasaAlmacenes_Click()
    EntradaProd.Show vbModal
End Sub
Private Sub SubMenExistencias_Click()
    FrmVerExisBodega.Show vbModal
End Sub
Private Sub SubMenFacturar_Click()
    frmFactura.Show vbModal
End Sub
Private Sub SubMenHacerRequisicion_Click()
    frmRequisicion.Show vbModal
End Sub
Private Sub SubMenImprimirOrdendeCompra_Click()
    frmOrdenCompra.Show vbModal
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
Private Sub SubMenNuevaMarca_Click()
    Marca.Show vbModal
End Sub
Private Sub SubMenNuevaMateriaPrima_Click()
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
    FrmAltaProdAlm3.Show vbModal
End Sub
Private Sub SubMenNuevoProveedor_Click()
    Proveedor.Show vbModal
End Sub
Private Sub SubMenOrdendeCompra_Click()
    FrmOrdenRapida.Show vbModal
End Sub
Private Sub SubMenPEndientes_Click()
    frmOCPend.Show vbModal
End Sub
Private Sub SubMenPEndientesdeEntrega_Click()
    frmEntradaExistenciasTemporales.Show vbModal
End Sub
Private Sub SubMenPerdidas_Click()
    frmPerdidas.Show vbModal
End Sub
Private Sub SubMenPermisos_Click()
    DarPerVenta.Show vbModal
End Sub
Private Sub SubmenPreordendecompra_Click()
    frmPreOrden.Show vbModal
End Sub
Private Sub SubMenPrestamoInternodeAlmacen1_Click()
    PrestamosCartuchos.Show vbModal
End Sub
Private Sub SubMenPrestamos_Click()
    FrmPrestamos.Show vbModal
End Sub
Private Sub SubMenProduccion_Click()
    frmProduccion.Show vbModal
End Sub
Private Sub SubMenPRoducir_Click()
    FrmCreaExis.Show vbModal
End Sub
Private Sub SubMenPromocion_Click()
    frmPromos.Show vbModal
End Sub
Private Sub SubMenPuntodeVenta_Click()
    Ventas.Show vbModal
End Sub
Private Sub SubMenReemplazar_Click()
    FrmSustiInv.Show vbModal
End Sub
Private Sub SubMenReemplazarInsumo_Click()
    EditarJRVarios.Show vbModal
End Sub
Private Sub SubMenReportesdeVentas_Click()
    FrmProdMasVend.Show vbModal
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
Private Sub SubMenSalidas_Click()
    Salidas.Show vbModal
End Sub
Private Sub SubMenScrap_Click()
    frmScrap.Show vbModal
End Sub
Private Sub SubMenSurtirSucursal_Click()
    frmSurtir.Show vbModal
End Sub
Private Sub SubMenSurtirVentaProgramada_Click()
    frmShowPediC.Show vbModal
End Sub
Private Sub SubMenTramitarGarantia_Click()
    frmGarantias.Show vbModal
End Sub
Private Sub SubMenTraspasosEntreSucursales_Click()
    Transfe.Show vbModal
End Sub
Private Sub SubMenVentasEspeciales_Click()
    PermisoVenta.Show vbModal
End Sub
Private Sub SubMenVentasProgramadas_Click()
    Programadas.Show vbModal
End Sub
Private Sub SubMenVer_Click()
    VerJuegoRep.Show vbModal
End Sub
Private Sub SubMenVerPedidos_Click()
    frmRevPed.Show vbModal
End Sub
Private Sub SubSalir_Click()
    Unload Me
End Sub
Private Sub SumMenCortedeCaja_Click()
    CorteCaja.Show vbModal
End Sub
Sub Sincronizar()
On Error GoTo ManejaError
    Me.lblEstado.Caption = "Sincronizando con los servidores, espere..."
    Me.lblEstado.ForeColor = vbRed
    DoEvents
    deAPTONER.TRAER_HORA_FECHA_SISTEMA
    With deAPTONER.rsTRAER_HORA_FECHA_SISTEMA
        Time = TimeValue(!FECHAHORA)
        Date = DateValue(!FECHAHORA)
        .Close
    End With
    Me.lblEstado.Caption = ""
    DoEvents
    Exit Sub
ManejaError:
    Err.Clear
End Sub
