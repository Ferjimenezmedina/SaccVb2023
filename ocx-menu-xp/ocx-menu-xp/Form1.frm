VERSION 5.00
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Hook Menu Xp"
   ClientHeight    =   2130
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   2640
      Top             =   720
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   4
      CheckBorderColor=   7021576
      SelMenuBorder   =   7021576
      SelMenuBackColor=   14073525
      SelMenuForeColor=   16646297
      SelCheckBackColor=   12294026
      MenuBorderColor =   6956042
      SeparatorColor  =   -2147483632
      MenuBackColor   =   14609903
      MenuForeColor   =   0
      CheckBackColor  =   15326939
      CheckForeColor  =   10027263
      DisabledMenuBorderColor=   -2147483632
      DisabledMenuBackColor=   15660791
      DisabledMenuForeColor=   -2147483631
      MenuBarBackColor=   -2147483644
      MenuPopupBackColor=   16777215
      ShortCutNormalColor=   0
      ShortCutSelectColor=   16646297
      ArrowNormalColor=   10027263
      ArrowSelectColor=   12484864
      ShadowColor     =   0
      Bmp:1           =   "Form1.frx":0000
      Key:1           =   "#mnuImprimir"
      Bmp:2           =   "Form1.frx":0428
      Key:2           =   "#mnuGuardar"
      Bmp:3           =   "Form1.frx":1190
      Key:3           =   "#mnusalir"
      Bmp:4           =   "Form1.frx":15B8
      Key:4           =   "#mnuVista"
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
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuAbrir 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu mnulinea 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnuVista 
         Caption         =   "&Vista preliminar"
      End
      Begin VB.Menu mnulinea2 
         Caption         =   "-"
      End
      Begin VB.Menu mnusalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

End Sub

Private Sub mnulinea_Click()

End Sub
