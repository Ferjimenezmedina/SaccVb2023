VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SACC (Sistema de Administración y Control del Comercio)"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   13050
   ClipControls    =   0   'False
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   13050
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   76
      Left            =   4560
      TabIndex        =   233
      Top             =   5520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   75
      Left            =   4680
      TabIndex        =   232
      Top             =   5520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   74
      Left            =   4800
      TabIndex        =   231
      Top             =   5520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000009&
      Caption         =   "Vale de Caja"
      DisabledPicture =   "frmMenu.frx":1601A
      Height          =   375
      Index           =   8
      Left            =   10080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":189EC
      Style           =   1  'Graphical
      TabIndex        =   230
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000009&
      Caption         =   "Aprov. Compra"
      DisabledPicture =   "frmMenu.frx":1B3BE
      Height          =   375
      Index           =   5
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":1DD90
      Style           =   1  'Graphical
      TabIndex        =   170
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000009&
      Caption         =   "Revisar"
      DisabledPicture =   "frmMenu.frx":20762
      Height          =   375
      Index           =   7
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":23134
      Style           =   1  'Graphical
      TabIndex        =   229
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000009&
      Caption         =   "Prestamo"
      DisabledPicture =   "frmMenu.frx":25B06
      Height          =   375
      Index           =   4
      Left            =   12480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":284D8
      Style           =   1  'Graphical
      TabIndex        =   177
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Material"
      DisabledPicture =   "frmMenu.frx":2AEAA
      Height          =   375
      Index           =   7
      Left            =   10080
      Picture         =   "frmMenu.frx":2D87C
      Style           =   1  'Graphical
      TabIndex        =   228
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Producto"
      DisabledPicture =   "frmMenu.frx":3024E
      Height          =   375
      Index           =   6
      Left            =   8880
      Picture         =   "frmMenu.frx":32C20
      Style           =   1  'Graphical
      TabIndex        =   227
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Vale de Caja"
      DisabledPicture =   "frmMenu.frx":355F2
      Height          =   375
      Index           =   8
      Left            =   8880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":37FC4
      Style           =   1  'Graphical
      TabIndex        =   226
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Caption         =   "J.R."
      DisabledPicture =   "frmMenu.frx":3A996
      Height          =   375
      Index           =   12
      Left            =   7680
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":3D368
      Style           =   1  'Graphical
      TabIndex        =   225
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Caption         =   "Comiciones"
      DisabledPicture =   "frmMenu.frx":3FD3A
      Height          =   375
      Index           =   8
      Left            =   6480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":4270C
      Style           =   1  'Graphical
      TabIndex        =   183
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000009&
      Caption         =   "EXCEL"
      DisabledPicture =   "frmMenu.frx":450DE
      Height          =   375
      Index           =   2
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":47AB0
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Eliminar B.O."
      DisabledPicture =   "frmMenu.frx":4A482
      Height          =   375
      Index           =   2
      Left            =   7680
      Picture         =   "frmMenu.frx":4CE54
      Style           =   1  'Graphical
      TabIndex        =   165
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H80000009&
      Caption         =   "B.O."
      DisabledPicture =   "frmMenu.frx":4F826
      Height          =   375
      Index           =   2
      Left            =   6480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":521F8
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Caption         =   "Pendientes"
      DisabledPicture =   "frmMenu.frx":54BCA
      Height          =   375
      Index           =   11
      Left            =   6480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":5759C
      Style           =   1  'Graphical
      TabIndex        =   224
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000009&
      Caption         =   "Existencia "
      DisabledPicture =   "frmMenu.frx":59F6E
      Height          =   375
      Index           =   8
      Left            =   7680
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":5C940
      Style           =   1  'Graphical
      TabIndex        =   160
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Facturar"
      DisabledPicture =   "frmMenu.frx":5F312
      Height          =   375
      Index           =   2
      Left            =   6480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":61CE4
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Especiales"
      DisabledPicture =   "frmMenu.frx":646B6
      Height          =   375
      Index           =   6
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":67088
      Style           =   1  'Graphical
      TabIndex        =   128
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Programadas"
      DisabledPicture =   "frmMenu.frx":69A5A
      Height          =   375
      Index           =   4
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":6C42C
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "P. de Venta"
      DisabledPicture =   "frmMenu.frx":6EDFE
      Height          =   375
      Index           =   1
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":717D0
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ventas"
      DisabledPicture =   "frmMenu.frx":741A2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":76B74
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H80000009&
      Caption         =   "Prestamos"
      DisabledPicture =   "frmMenu.frx":79546
      Height          =   375
      Index           =   6
      Left            =   6480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":7BF18
      Style           =   1  'Graphical
      TabIndex        =   157
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H80000009&
      Caption         =   "Autorizar Rema"
      DisabledPicture =   "frmMenu.frx":7E8EA
      Height          =   375
      Index           =   5
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":812BC
      Style           =   1  'Graphical
      TabIndex        =   156
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H80000009&
      Caption         =   "Autorizar Gtia"
      DisabledPicture =   "frmMenu.frx":83C8E
      Height          =   375
      Index           =   4
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":86660
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H80000009&
      Caption         =   "Garantias"
      DisabledPicture =   "frmMenu.frx":89032
      Height          =   375
      Index           =   1
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":8BA04
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Atención"
      DisabledPicture =   "frmMenu.frx":8E3D6
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "frmMenu.frx":90DA8
      Style           =   1  'Graphical
      TabIndex        =   135
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000009&
      Caption         =   "Cambio Precio"
      DisabledPicture =   "frmMenu.frx":959EA
      Height          =   375
      Index           =   1
      Left            =   8880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":983BC
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000009&
      Caption         =   "Permisos"
      DisabledPicture =   "frmMenu.frx":9AD8E
      Height          =   375
      Index           =   3
      Left            =   7680
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":9D760
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Cancelar"
      DisabledPicture =   "frmMenu.frx":A0132
      Height          =   375
      Index           =   7
      Left            =   6480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":A2B04
      Style           =   1  'Graphical
      TabIndex        =   159
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000009&
      Caption         =   "Promoción"
      DisabledPicture =   "frmMenu.frx":A54D6
      Height          =   375
      Index           =   1
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":A7EA8
      Style           =   1  'Graphical
      TabIndex        =   117
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000009&
      Caption         =   "Licitación"
      DisabledPicture =   "frmMenu.frx":AA87A
      Height          =   375
      Index           =   2
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":AD24C
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Corte"
      DisabledPicture =   "frmMenu.frx":AFC1E
      Height          =   375
      Index           =   3
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":B25F0
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Administración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "frmMenu.frx":B4FC2
      Style           =   1  'Graphical
      TabIndex        =   201
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000009&
      Caption         =   "Requisición"
      DisabledPicture =   "frmMenu.frx":B9C04
      Height          =   375
      Index           =   7
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":BC5D6
      Style           =   1  'Graphical
      TabIndex        =   155
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000009&
      Caption         =   "Ver Pedidos"
      DisabledPicture =   "frmMenu.frx":BEFA8
      Height          =   375
      Index           =   1
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":C197A
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Pedidos"
      DisabledPicture =   "frmMenu.frx":C434C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":C6D1E
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000009&
      Caption         =   "Reemplazar"
      DisabledPicture =   "frmMenu.frx":C96F0
      Height          =   375
      Index           =   5
      Left            =   11280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":CC0C2
      Style           =   1  'Graphical
      TabIndex        =   182
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000009&
      Caption         =   "Perdidas"
      DisabledPicture =   "frmMenu.frx":CEA94
      Height          =   375
      Index           =   4
      Left            =   10080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":D1466
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000009&
      Caption         =   "Producir"
      DisabledPicture =   "frmMenu.frx":D3E38
      Height          =   375
      Index           =   5
      Left            =   8880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":D680A
      Style           =   1  'Graphical
      TabIndex        =   178
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000009&
      Caption         =   "Ajustar Venta"
      DisabledPicture =   "frmMenu.frx":D91DC
      Height          =   375
      Index           =   6
      Left            =   7680
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":DBBAE
      Style           =   1  'Graphical
      TabIndex        =   133
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000009&
      Caption         =   "Surtir Venta"
      DisabledPicture =   "frmMenu.frx":DE580
      Height          =   375
      Index           =   5
      Left            =   6480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":E0F52
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000009&
      Caption         =   "Surtir Sucursal"
      DisabledPicture =   "frmMenu.frx":E3924
      Height          =   375
      Index           =   4
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":E62F6
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000009&
      Caption         =   "Salida"
      DisabledPicture =   "frmMenu.frx":E8CC8
      Height          =   375
      Index           =   3
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":EB69A
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000009&
      Caption         =   "Inventarios"
      DisabledPicture =   "frmMenu.frx":EE06C
      Height          =   375
      Index           =   2
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":F0A3E
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Movimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "frmMenu.frx":F3410
      Style           =   1  'Graphical
      TabIndex        =   204
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000009&
      Caption         =   "Ventas"
      DisabledPicture =   "frmMenu.frx":F8052
      Height          =   375
      Index           =   6
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":FAA24
      Style           =   1  'Graphical
      TabIndex        =   179
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000009&
      Caption         =   "Almacen"
      DisabledPicture =   "frmMenu.frx":FD3F6
      Height          =   375
      Index           =   2
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":FFDC8
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000009&
      Caption         =   "Traspasos"
      DisabledPicture =   "frmMenu.frx":10279A
      Height          =   375
      Index           =   3
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":10516C
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000009&
      Caption         =   "Orden"
      DisabledPicture =   "frmMenu.frx":107B3E
      Height          =   375
      Index           =   1
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":10A510
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Entradas"
      DisabledPicture =   "frmMenu.frx":10CEE2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "frmMenu.frx":10F8B4
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Caption         =   "Orden Comp."
      DisabledPicture =   "frmMenu.frx":1144F6
      Height          =   375
      Index           =   5
      Left            =   6480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":116EC8
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Caption         =   "Asignar"
      DisabledPicture =   "frmMenu.frx":11989A
      Height          =   375
      Index           =   4
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":11C26C
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Caption         =   "Revisar"
      DisabledPicture =   "frmMenu.frx":11EC3E
      Height          =   375
      Index           =   3
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":121610
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Caption         =   "Requisición"
      DisabledPicture =   "frmMenu.frx":123FE2
      Height          =   375
      Index           =   2
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":1269B4
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cotizar"
      DisabledPicture =   "frmMenu.frx":129386
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":12BD58
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Caption         =   "Orden Rapida"
      DisabledPicture =   "frmMenu.frx":12E72A
      Height          =   375
      Index           =   7
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":1310FC
      Style           =   1  'Graphical
      TabIndex        =   176
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Caption         =   "Imprimir"
      DisabledPicture =   "frmMenu.frx":133ACE
      Height          =   375
      Index           =   6
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":1364A0
      Style           =   1  'Graphical
      TabIndex        =   134
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Caption         =   "Autorizar"
      DisabledPicture =   "frmMenu.frx":138E72
      Height          =   375
      Index           =   9
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":13B844
      Style           =   1  'Graphical
      TabIndex        =   184
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command18 
      Caption         =   "O.C."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "frmMenu.frx":13E216
      Style           =   1  'Graphical
      TabIndex        =   207
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Caption         =   "Almacen 1"
      DisabledPicture =   "frmMenu.frx":142E58
      Height          =   375
      Index           =   1
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":14582A
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Materia Prima"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "frmMenu.frx":1481FC
      Style           =   1  'Graphical
      TabIndex        =   208
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000009&
      Caption         =   "Calidad"
      DisabledPicture =   "frmMenu.frx":14CE3E
      Height          =   375
      Index           =   3
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":14F810
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000009&
      Caption         =   "Producción"
      DisabledPicture =   "frmMenu.frx":1521E2
      Height          =   375
      Index           =   2
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":154BB4
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000009&
      Caption         =   "Revisión"
      DisabledPicture =   "frmMenu.frx":157586
      Height          =   375
      Index           =   1
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":159F58
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Procesos"
      DisabledPicture =   "frmMenu.frx":15C92A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":15F2FC
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000009&
      Caption         =   "Material Extra"
      DisabledPicture =   "frmMenu.frx":161CCE
      Height          =   375
      Index           =   6
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":1646A0
      Style           =   1  'Graphical
      TabIndex        =   167
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000009&
      Caption         =   "Scrap"
      DisabledPicture =   "frmMenu.frx":167072
      Height          =   375
      Index           =   8
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":169A44
      Style           =   1  'Graphical
      TabIndex        =   186
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Revisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "frmMenu.frx":16C416
      Style           =   1  'Graphical
      TabIndex        =   213
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000009&
      Caption         =   "Ver JR"
      DisabledPicture =   "frmMenu.frx":171058
      Height          =   375
      Index           =   4
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":173A2A
      Style           =   1  'Graphical
      TabIndex        =   152
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000009&
      Caption         =   "Editar JR"
      DisabledPicture =   "frmMenu.frx":1763FC
      Height          =   375
      Index           =   5
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":178DCE
      Style           =   1  'Graphical
      TabIndex        =   161
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Nuevo JR"
      DisabledPicture =   "frmMenu.frx":17B7A0
      Height          =   375
      Index           =   8
      Left            =   2880
      Picture         =   "frmMenu.frx":17E172
      Style           =   1  'Graphical
      TabIndex        =   158
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command21 
      Caption         =   "J.R."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "frmMenu.frx":180B44
      Style           =   1  'Graphical
      TabIndex        =   214
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000009&
      Caption         =   "Pendientes"
      DisabledPicture =   "frmMenu.frx":185786
      Height          =   375
      Index           =   1
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":188158
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tecnicos"
      DisabledPicture =   "frmMenu.frx":18AB2A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":18D4FC
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "Domicilios"
      DisabledPicture =   "frmMenu.frx":18FECE
      Height          =   375
      Index           =   7
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":1928A0
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Mensajeros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "frmMenu.frx":195272
      Style           =   1  'Graphical
      TabIndex        =   218
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000009&
      Caption         =   "Tipo Cambio"
      DisabledPicture =   "frmMenu.frx":199EB4
      Height          =   375
      Index           =   4
      Left            =   7680
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":19C886
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000009&
      Caption         =   "Pago Compra"
      DisabledPicture =   "frmMenu.frx":19F258
      Height          =   375
      Index           =   6
      Left            =   6480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":1A1C2A
      Style           =   1  'Graphical
      TabIndex        =   171
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H80000009&
      Caption         =   "Nota Credito"
      DisabledPicture =   "frmMenu.frx":1A45FC
      Height          =   375
      Index           =   3
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":1A6FCE
      Style           =   1  'Graphical
      TabIndex        =   129
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Creditos"
      DisabledPicture =   "frmMenu.frx":1A99A0
      Height          =   375
      Index           =   5
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":1AC372
      Style           =   1  'Graphical
      TabIndex        =   127
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Caption         =   "Pagar OC"
      DisabledPicture =   "frmMenu.frx":1AED44
      Height          =   375
      Index           =   10
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":1B1716
      Style           =   1  'Graphical
      TabIndex        =   187
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Contabilidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "frmMenu.frx":1B40E8
      Style           =   1  'Graphical
      TabIndex        =   219
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000009&
      Caption         =   "Marcas"
      DisabledPicture =   "frmMenu.frx":1B8D2A
      Height          =   375
      Index           =   3
      Left            =   11280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":1BB6FC
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Material"
      DisabledPicture =   "frmMenu.frx":1BE0CE
      Height          =   375
      Index           =   7
      Left            =   10080
      Picture         =   "frmMenu.frx":1C0AA0
      Style           =   1  'Graphical
      TabIndex        =   153
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Producto"
      DisabledPicture =   "frmMenu.frx":1C3472
      Height          =   375
      Index           =   6
      Left            =   8880
      Picture         =   "frmMenu.frx":1C5E44
      Style           =   1  'Graphical
      TabIndex        =   148
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Mensajero"
      DisabledPicture =   "frmMenu.frx":1C8816
      Height          =   375
      Index           =   5
      Left            =   7680
      Picture         =   "frmMenu.frx":1CB1E8
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Proveedor"
      DisabledPicture =   "frmMenu.frx":1CDBBA
      Height          =   375
      Index           =   4
      Left            =   6480
      Picture         =   "frmMenu.frx":1D058C
      Style           =   1  'Graphical
      TabIndex        =   141
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Sucursal"
      DisabledPicture =   "frmMenu.frx":1D2F5E
      Height          =   375
      Index           =   3
      Left            =   5280
      Picture         =   "frmMenu.frx":1D5930
      Style           =   1  'Graphical
      TabIndex        =   140
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Cliente"
      DisabledPicture =   "frmMenu.frx":1D8302
      Height          =   375
      Index           =   2
      Left            =   4080
      Picture         =   "frmMenu.frx":1DACD4
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Agente"
      DisabledPicture =   "frmMenu.frx":1DD6A6
      Height          =   375
      Index           =   1
      Left            =   2880
      Picture         =   "frmMenu.frx":1E0078
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Nuevo"
      DisabledPicture =   "frmMenu.frx":1E2A4A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "frmMenu.frx":1E541C
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Mensajero"
      DisabledPicture =   "frmMenu.frx":1EA05E
      Height          =   375
      Index           =   5
      Left            =   7680
      Picture         =   "frmMenu.frx":1ECA30
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Proveedor"
      DisabledPicture =   "frmMenu.frx":1EF402
      Height          =   375
      Index           =   4
      Left            =   6480
      Picture         =   "frmMenu.frx":1F1DD4
      Style           =   1  'Graphical
      TabIndex        =   147
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Sucursal"
      DisabledPicture =   "frmMenu.frx":1F47A6
      Height          =   375
      Index           =   3
      Left            =   5280
      Picture         =   "frmMenu.frx":1F7178
      Style           =   1  'Graphical
      TabIndex        =   146
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Cliente"
      DisabledPicture =   "frmMenu.frx":1F9B4A
      Height          =   375
      Index           =   2
      Left            =   4080
      Picture         =   "frmMenu.frx":1FC51C
      Style           =   1  'Graphical
      TabIndex        =   145
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Agente"
      DisabledPicture =   "frmMenu.frx":1FEEEE
      Height          =   375
      Index           =   1
      Left            =   2880
      Picture         =   "frmMenu.frx":2018C0
      Style           =   1  'Graphical
      TabIndex        =   144
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Eliminar"
      DisabledPicture =   "frmMenu.frx":204292
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "frmMenu.frx":206C64
      Style           =   1  'Graphical
      TabIndex        =   143
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000009&
      Caption         =   "Empresa"
      DisabledPicture =   "frmMenu.frx":20B8A6
      Height          =   375
      Index           =   7
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":20E278
      Style           =   1  'Graphical
      TabIndex        =   189
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sistema"
      DisabledPicture =   "frmMenu.frx":210C4A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":21361C
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000009&
      Caption         =   "Reportes"
      DisabledPicture =   "frmMenu.frx":215FEE
      Height          =   375
      Index           =   6
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":2189C0
      Style           =   1  'Graphical
      TabIndex        =   188
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Reportes"
      DisabledPicture =   "frmMenu.frx":21B392
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":21DD64
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Rastrear"
      DisabledPicture =   "frmMenu.frx":220736
      Height          =   375
      Index           =   4
      Left            =   5280
      Picture         =   "frmMenu.frx":223108
      Style           =   1  'Graphical
      TabIndex        =   185
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Ordenes"
      DisabledPicture =   "frmMenu.frx":225ADA
      Height          =   375
      Index           =   3
      Left            =   4080
      Picture         =   "frmMenu.frx":2284AC
      Style           =   1  'Graphical
      TabIndex        =   166
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Hacer"
      DisabledPicture =   "frmMenu.frx":22AE7E
      Height          =   375
      Index           =   1
      Left            =   2880
      Picture         =   "frmMenu.frx":22D850
      Style           =   1  'Graphical
      TabIndex        =   164
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Pedidos"
      DisabledPicture =   "frmMenu.frx":230222
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "frmMenu.frx":232BF4
      Style           =   1  'Graphical
      TabIndex        =   154
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "Entradas"
      DisabledPicture =   "frmMenu.frx":237836
      Height          =   375
      Index           =   8
      Left            =   10080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":23A208
      Style           =   1  'Graphical
      TabIndex        =   130
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "Producto"
      DisabledPicture =   "frmMenu.frx":23CBDA
      Height          =   375
      Index           =   6
      Left            =   8880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":23F5AC
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "Existencia"
      DisabledPicture =   "frmMenu.frx":241F7E
      Height          =   375
      Index           =   5
      Left            =   7680
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":244950
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "Prod-Pedido"
      DisabledPicture =   "frmMenu.frx":247322
      Height          =   375
      Index           =   4
      Left            =   6480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":249CF4
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "Rastrear"
      DisabledPicture =   "frmMenu.frx":24C6C6
      Height          =   375
      Index           =   3
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":24F098
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "Faltantes"
      DisabledPicture =   "frmMenu.frx":251A6A
      Height          =   375
      Index           =   2
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":25443C
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "Proveedores"
      DisabledPicture =   "frmMenu.frx":256E0E
      Height          =   375
      Index           =   1
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":2597E0
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Consultas"
      DisabledPicture =   "frmMenu.frx":25C1B2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":25EB84
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H80000009&
      Caption         =   "Salir"
      DisabledPicture =   "frmMenu.frx":261556
      Height          =   375
      Index           =   3
      Left            =   3840
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":263F28
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H80000009&
      Caption         =   "Cerrar Sesión"
      DisabledPicture =   "frmMenu.frx":2668FA
      Height          =   375
      Index           =   2
      Left            =   2640
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":2692CC
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H80000009&
      Caption         =   "Bloquear"
      DisabledPicture =   "frmMenu.frx":26BC9E
      Height          =   375
      Index           =   1
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":26E670
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Salir"
      DisabledPicture =   "frmMenu.frx":271042
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmMenu.frx":273A14
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   7560
      Width           =   1455
   End
   Begin VB.TextBox txtServidor 
      Height          =   285
      Left            =   3480
      TabIndex        =   200
      Top             =   7440
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   9
      Left            =   3120
      TabIndex        =   199
      Text            =   "Text4"
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   8
      Left            =   3000
      TabIndex        =   198
      Text            =   "Text4"
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   7
      Left            =   2880
      TabIndex        =   197
      Text            =   "Text4"
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   6
      Left            =   2760
      TabIndex        =   196
      Text            =   "Text4"
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   5
      Left            =   2640
      TabIndex        =   195
      Text            =   "Text4"
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   194
      Text            =   "Text4"
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   193
      Text            =   "Text4"
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   192
      Text            =   "Text4"
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   191
      Text            =   "Text4"
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   190
      Text            =   "Text4"
      Top             =   6360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   6
      Left            =   2880
      TabIndex        =   126
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   76
      Top             =   8325
      Width           =   13050
      _ExtentX        =   23019
      _ExtentY        =   529
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
            Bevel           =   0
            TextSave        =   "03:55 p.m."
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   0
            TextSave        =   "02/10/2007"
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
   Begin VB.Frame fraMensajes 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   5040
      TabIndex        =   71
      Top             =   5160
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   4560
         TabIndex        =   72
         Top             =   240
         Width           =   975
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Leer"
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
            TabIndex        =   73
            Top             =   960
            Width           =   975
         End
         Begin VB.Image imgLeer 
            Height          =   630
            Left            =   120
            MouseIcon       =   "frmMenu.frx":29E9B6
            MousePointer    =   99  'Custom
            Picture         =   "frmMenu.frx":29ECC0
            Top             =   240
            Width           =   765
         End
      End
      Begin VB.CommandButton cmdLeerMSN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LEER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   840
         Width           =   855
      End
      Begin VB.Image frsfrsfrms 
         Height          =   1410
         Left            =   120
         Picture         =   "frmMenu.frx":2A069A
         Top             =   240
         Width           =   1410
      End
      Begin VB.Image Image3 
         Height          =   810
         Left            =   1920
         Picture         =   "frmMenu.frx":2A6F24
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   69
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   68
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   67
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   66
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   65
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   2640
      TabIndex        =   64
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   63
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   62
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   61
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   3
      Left            =   3360
      TabIndex        =   60
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   4
      Left            =   3480
      TabIndex        =   59
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   5
      Left            =   3600
      TabIndex        =   58
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   4200
      TabIndex        =   57
      Top             =   5760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   4320
      TabIndex        =   56
      Top             =   5760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   4440
      TabIndex        =   55
      Top             =   5760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   4560
      TabIndex        =   54
      Top             =   5760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   4680
      TabIndex        =   53
      Top             =   5760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   4800
      TabIndex        =   52
      Top             =   5760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   4920
      TabIndex        =   51
      Top             =   5760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   5040
      TabIndex        =   50
      Top             =   5760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   5160
      TabIndex        =   49
      Top             =   5760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   5280
      TabIndex        =   48
      Top             =   5760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   4200
      TabIndex        =   47
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   4320
      TabIndex        =   46
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   4440
      TabIndex        =   45
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   4560
      TabIndex        =   44
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   4680
      TabIndex        =   43
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   21
      Left            =   4800
      TabIndex        =   42
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   4920
      TabIndex        =   41
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   5040
      TabIndex        =   40
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   24
      Left            =   5160
      TabIndex        =   39
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   25
      Left            =   5280
      TabIndex        =   38
      Top             =   6000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   4200
      TabIndex        =   37
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   27
      Left            =   4320
      TabIndex        =   36
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   28
      Left            =   4440
      TabIndex        =   35
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   29
      Left            =   4560
      TabIndex        =   34
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   30
      Left            =   4680
      TabIndex        =   33
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   31
      Left            =   4800
      TabIndex        =   32
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   32
      Left            =   4920
      TabIndex        =   31
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   33
      Left            =   5040
      TabIndex        =   30
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   34
      Left            =   5160
      TabIndex        =   29
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   35
      Left            =   5280
      TabIndex        =   28
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   36
      Left            =   4200
      TabIndex        =   27
      Top             =   6480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   37
      Left            =   4320
      TabIndex        =   26
      Top             =   6480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   38
      Left            =   4440
      TabIndex        =   25
      Top             =   6480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   39
      Left            =   4560
      TabIndex        =   24
      Top             =   6480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   40
      Left            =   4680
      TabIndex        =   23
      Top             =   6480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   41
      Left            =   4800
      TabIndex        =   22
      Top             =   6480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   42
      Left            =   4920
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   43
      Left            =   5040
      TabIndex        =   20
      Top             =   6480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   44
      Left            =   5160
      TabIndex        =   19
      Top             =   6480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   45
      Left            =   5280
      TabIndex        =   18
      Top             =   6480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   46
      Left            =   4200
      TabIndex        =   17
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   47
      Left            =   4320
      TabIndex        =   16
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   48
      Left            =   4440
      TabIndex        =   15
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   49
      Left            =   4560
      TabIndex        =   14
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   50
      Left            =   4680
      TabIndex        =   13
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   51
      Left            =   4800
      TabIndex        =   12
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   52
      Left            =   4920
      TabIndex        =   11
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   53
      Left            =   5040
      TabIndex        =   10
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   54
      Left            =   5160
      TabIndex        =   9
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   55
      Left            =   5280
      TabIndex        =   8
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   56
      Left            =   4200
      TabIndex        =   7
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   57
      Left            =   4320
      TabIndex        =   6
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   58
      Left            =   4440
      TabIndex        =   5
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   59
      Left            =   4560
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   60
      Left            =   4680
      TabIndex        =   3
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   61
      Left            =   4800
      TabIndex        =   2
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   62
      Left            =   4920
      TabIndex        =   1
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   63
      Left            =   5040
      TabIndex        =   0
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   64
      Left            =   5160
      TabIndex        =   151
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   65
      Left            =   5280
      TabIndex        =   162
      Top             =   6960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   66
      Left            =   4200
      TabIndex        =   168
      Top             =   7200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   67
      Left            =   5520
      TabIndex        =   169
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   68
      Left            =   5640
      TabIndex        =   172
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   69
      Left            =   5760
      TabIndex        =   173
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   70
      Left            =   5880
      TabIndex        =   174
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   71
      Left            =   6000
      TabIndex        =   175
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   72
      Left            =   6120
      TabIndex        =   180
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   73
      Left            =   6240
      TabIndex        =   181
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   0
      Picture         =   "frmMenu.frx":2AF740
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   202
      Top             =   0
      Width           =   1455
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VENTAS"
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
         TabIndex        =   203
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   1095
      Left            =   0
      Picture         =   "frmMenu.frx":2DA6E2
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   209
      Top             =   1080
      Width           =   1455
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COMPRAS"
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
         TabIndex        =   210
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   0
      Picture         =   "frmMenu.frx":305684
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   205
      Top             =   2160
      Width           =   1455
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ALMACEN"
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
         TabIndex        =   206
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   1095
      Left            =   0
      Picture         =   "frmMenu.frx":330626
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   211
      Top             =   3240
      Width           =   1455
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCCIÓN"
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
         TabIndex        =   212
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture6 
      Height          =   1095
      Left            =   0
      Picture         =   "frmMenu.frx":35B5C8
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   220
      Top             =   4320
      Width           =   1455
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADMON."
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
         TabIndex        =   221
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   1095
      Left            =   0
      Picture         =   "frmMenu.frx":38656A
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   215
      Top             =   5400
      Width           =   1455
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VARIOS"
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
         TabIndex        =   217
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DEPTOS."
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
         TabIndex        =   216
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture7 
      Height          =   1095
      Left            =   0
      Picture         =   "frmMenu.frx":3B150C
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   222
      Top             =   6480
      Width           =   1455
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "UTILERIAS"
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
         TabIndex        =   223
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Image Image8 
      Height          =   1410
      Left            =   9840
      Picture         =   "frmMenu.frx":3DC4AE
      Top             =   3720
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Height          =   11055
      Left            =   0
      TabIndex        =   124
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   77
      Top             =   7560
      Width           =   5775
   End
   Begin VB.Label lblPuestoSucursal 
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
      Left            =   4800
      TabIndex        =   75
      Top             =   4920
      Width           =   8175
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
      Left            =   3960
      TabIndex        =   70
      Top             =   1080
      Width           =   8175
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   5040
      Picture         =   "frmMenu.frx":3E2D38
      Top             =   2640
      Width           =   4950
   End
   Begin VB.Image Image7 
      Height          =   12030
      Left            =   9000
      Picture         =   "frmMenu.frx":4072BA
      Top             =   -120
      Width           =   4080
   End
   Begin VB.Image Image6 
      Height          =   9000
      Left            =   1440
      Picture         =   "frmMenu.frx":4A6F5C
      Top             =   -120
      Width           =   5250
   End
End
Attribute VB_Name = "Menu"
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
' Última revisión:                                                  (04/Abr/01)
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
    sRetVal = String$(255, 0)
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
    ' Esta función devolverá un array de índice cero
    ' con las claves y valores de la sección
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
Private Sub Command1_Click(Index As Integer)
    If Index = 1 Then
        Borra
        Ventas.Show vbModal
    End If
    If Index = 2 Then
        Borra
        frmFactura.Show vbModal
    End If
    If Index = 3 Then
        Borra
        FrmCorteCredito.Show vbModal
    End If
    If Index = 4 Then
        Borra
        Programadas.Show vbModal
    End If
    If Index = 5 Then
        Borra
        Creditos.Show vbModal
    End If
    If Index = 6 Then
        Borra
        PermisoVenta.Show vbModal
    End If
    If Index = 7 Then
        Borra
        frmCancelaFactura.Show vbModal
    End If
    If Index = 8 Then
        Borra
        FrmValeCajaCerrar.Show vbModal
    End If
End Sub
Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 1
        NueMenBot
    End If
End Sub
Private Sub Command10_Click(Index As Integer)
    If Index = 1 Then
        Borra
        frmLogin.Show vbModal, Me
    End If
    If Index = 2 Then
        Borra
        Unload Me
        VarMen.Show
    End If
    If Index = 3 Then
        End
    End If
    Borra
End Sub
Private Sub Command10_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Command1(0).Visible = False
        Command2(0).Visible = False
        Command3(0).Visible = False
        Command4(0).Visible = False
        Command5(0).Visible = False
        Command6(0).Visible = False
        Command7(0).Visible = False
        Command8(0).Visible = False
        Command11(0).Visible = False
        Command12(0).Visible = False
        Command13(0).Visible = False
        Command14(0).Visible = False
        Command15(0).Visible = False
        Command16(0).Visible = False
        Command17(0).Visible = False
        Command18(0).Visible = False
        Command19(0).Visible = False
        Command20(0).Visible = False
        Command21(0).Visible = False
        Command22(0).Visible = False
        Command23(0).Visible = False
        Me.Command1(1).Visible = False
        Me.Command1(2).Visible = False
        Me.Command1(3).Visible = False
        Me.Command1(4).Visible = False
        Me.Command1(5).Visible = False
        Me.Command1(6).Visible = False
        Me.Command1(7).Visible = False
        Me.Command1(8).Visible = False
        Me.Command2(1).Visible = False
        Me.Command2(2).Visible = False
        Me.Command2(3).Visible = False
        Me.Command2(4).Visible = False
        Me.Command2(5).Visible = False
        Me.Command2(6).Visible = False
        Me.Command2(7).Visible = False
        Me.Command2(8).Visible = False
        Me.Command3(1).Visible = False
        Me.Command3(2).Visible = False
        Me.Command3(3).Visible = False
        Me.Command3(4).Visible = False
        Me.Command3(5).Visible = False
        Me.Command3(6).Visible = False
        Me.Command3(7).Visible = False
        Me.Command3(8).Visible = False
        Me.Command4(1).Visible = False
        Me.Command4(2).Visible = False
        Me.Command4(3).Visible = False
        Me.Command4(4).Visible = False
        Me.Command4(5).Visible = False
        Me.Command4(6).Visible = False
        Me.Command4(7).Visible = False
        Me.Command4(8).Visible = False
        Me.Command4(9).Visible = False
        Me.Command4(10).Visible = False
        Me.Command4(11).Visible = False
        Me.Command4(12).Visible = False
        Me.Command5(1).Visible = False
        Me.Command6(1).Visible = False
        Me.Command6(2).Visible = False
        Me.Command6(3).Visible = False
        Me.Command6(4).Visible = False
        Me.Command6(5).Visible = False
        Me.Command6(6).Visible = False
        Me.Command6(7).Visible = False
        Me.Command6(8).Visible = False
        Me.Command7(1).Visible = False
        Me.Command7(2).Visible = False
        Me.Command7(3).Visible = False
        Me.Command7(4).Visible = False
        Me.Command7(5).Visible = False
        Me.Command7(6).Visible = False
        Me.Command7(7).Visible = False
        Me.Command7(8).Visible = False
        Me.Command8(1).Visible = False
        Me.Command8(2).Visible = False
        Me.Command8(3).Visible = False
        Me.Command8(4).Visible = False
        Me.Command8(5).Visible = False
        Me.Command8(6).Visible = False
        Me.Command11(1).Visible = False
        Me.Command11(2).Visible = False
        Me.Command11(3).Visible = False
        Me.Command11(4).Visible = False
        Me.Command11(5).Visible = False
        Me.Command11(6).Visible = False
        Me.Command12(1).Visible = False
        Me.Command12(2).Visible = False
        Me.Command12(3).Visible = False
        Me.Command12(4).Visible = False
        Me.Command12(5).Visible = False
        Me.Command12(6).Visible = False
        Me.Command13(1).Visible = False
        Me.Command13(2).Visible = False
        Me.Command13(3).Visible = False
        Me.Command13(4).Visible = False
        Me.Command13(5).Visible = False
        Me.Command13(6).Visible = False
        Me.Command13(7).Visible = False
        Me.Command13(8).Visible = False
        Me.Command14(1).Visible = False
        Me.Command14(2).Visible = False
        Me.Command14(3).Visible = False
        Me.Command14(4).Visible = False
        Me.Command14(5).Visible = False
        Me.Command14(6).Visible = False
        Me.Command14(7).Visible = False
        Me.Command15(1).Visible = False
        Me.Command15(2).Visible = False
        Me.Command15(3).Visible = False
        Me.Command15(4).Visible = False
        Validar = 10
        NueMenBot
    End If
End Sub
Private Sub Command11_Click(Index As Integer)
    If Index = 1 Then
        Borra
        frmGarantias.Show vbModal
    End If
    If Index = 2 Then
        Borra
        frmEntradaExistenciasTemporales.Show vbModal
    End If
    If Index = 3 Then
        Borra
        NotaCredito.Show vbModal
    End If
    If Index = 4 Then
        Borra
        FrmAutGarantia.Show vbModal
    End If
    If Index = 5 Then
        Borra
        frmAutRema.Show vbModal
    End If
    If Index = 6 Then
        Borra
        FrmPrestamos.Show vbModal
    End If
End Sub
Private Sub Command11_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 11
        NueMenBot
    End If
End Sub
Private Sub Command12_Click(Index As Integer)
    If Index = 1 Then
        Borra
        Ordenes.Show vbModal
    End If
    If Index = 2 Then
        Borra
        EntradaProd.Show vbModal
    End If
    If Index = 3 Then
        Borra
        Transfe.Show vbModal
    End If
    If Index = 4 Then
        Borra
        PrestamosCartuchos.Show vbModal
    End If
    If Index = 5 Then
        Borra
        FrmCreaExis.Show vbModal
    End If
    If Index = 6 Then
        Borra
        FrmProdMasVend.Show vbModal
    End If
End Sub
Private Sub Command12_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 12
        NueMenBot
    End If
End Sub
Private Sub Command13_Click(Index As Integer)
    If Index = 1 Then
        Borra
        frmPermisos.Show vbModal
    End If
    If Index = 2 Then
        Borra
        AltaClien.Show vbModal
    End If
    If Index = 3 Then
        Borra
        AltaSucu.Show vbModal
    End If
    If Index = 4 Then
        Borra
        Proveedor.Show vbModal
    End If
    If Index = 5 Then
        Borra
        FrmNueRep.Show vbModal
    End If
    If Index = 6 Then
        Borra
        FrmAltaProdAlm3.Show vbModal
    End If
    If Index = 7 Then
        Borra
        FrmAltaProdAlm1y2.Show vbModal
    End If
    If Index = 8 Then
        Borra
        FrmNuevoJR.Show vbModal
    End If
End Sub
Private Sub Command13_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 13
        NueMenBot
    End If
End Sub
Private Sub Command14_Click(Index As Integer)
    If Index = 1 Then
        Borra
        EliAgente.Show vbModal
    End If
    If Index = 2 Then
        Borra
        EliCliente.Show vbModal
    End If
    If Index = 3 Then
        Borra
        EliSuc.Show vbModal
    End If
    If Index = 4 Then
        Borra
        EliProveedor.Show vbModal
    End If
    If Index = 5 Then
        Borra
        EliMensajero.Show vbModal
    End If
    If Index = 6 Then
        Borra
        FrmEliProdAlm3.Show vbModal
    End If
    If Index = 7 Then
        Borra
        FrmEliProdAlm1y2.Show vbModal
    End If
End Sub
Private Sub Command14_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 14
        NueMenBot
    End If
End Sub
Private Sub Command15_Click(Index As Integer)
    If Index = 1 Then
        Borra
        Pedidos.Show vbModal
    End If
    If Index = 2 Then
        Borra
        frmSalidaExistenciasTemporales.Show vbModal
    End If
    If Index = 3 Then
        Borra
        frmOrdenesProduccion.Show vbModal
    End If
    If Index = 4 Then
        Borra
        FrmRastPed.Show vbModal
    End If
    Borra
End Sub
Private Sub Command15_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 15
        NueMenBot
    End If
End Sub
Private Sub Command16_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 16
        NueMenBot
    End If
End Sub
Private Sub Command17_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 17
        NueMenBot
    End If
End Sub
Private Sub Command18_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 18
        NueMenBot
    End If
End Sub
Private Sub Command19_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 19
        NueMenBot
    End If
End Sub
Private Sub Command2_Click(Index As Integer)
    If Index = 1 Then
        Borra
        frmProveedores.Show vbModal
    End If
    If Index = 2 Then
        Borra
        Faltantes.Show vbModal
    End If
    If Index = 3 Then
        Borra
        FrmRastrearPed.Show vbModal
    End If
    If Index = 4 Then
        Borra
        FrmBusProdPed.Show vbModal
    End If
    If Index = 5 Then
        Borra
        BuscaExist.Show vbModal
    End If
    If Index = 6 Then
        Borra
        BuscaProd.Show vbModal
    End If
    If Index = 7 Then
        Borra
        FrmRevDomi.Show vbModal
    End If
    If Index = 8 Then
        Borra
        BuscaEntrada.Show vbModal
    End If
End Sub
Private Sub Command2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 2
        NueMenBot
    End If
End Sub
Private Sub Command20_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 20
        NueMenBot
    End If
End Sub
Private Sub Command21_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 21
        NueMenBot
    End If
End Sub
Private Sub Command22_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 22
        NueMenBot
    End If
End Sub
Private Sub Command23_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 23
        NueMenBot
    End If
End Sub
Private Sub Command3_Click(Index As Integer)
    If Index = 1 Then
        Borra
        frmRevPed.Show vbModal
    End If
    If Index = 2 Then
        Borra
        frmInventarios2.Show vbModal
    End If
    If Index = 3 Then
        Borra
        Salidas.Show vbModal
    End If
    If Index = 4 Then
        Borra
        frmSurtir.Show vbModal
    End If
    If Index = 5 Then
        Borra
        frmShowPediC.Show vbModal
    End If
    If Index = 6 Then
        Borra
        PermisoAjuste.Show vbModal
    End If
    If Index = 7 Then
        Borra
        frmRequisicion.Show vbModal
    End If
    If Index = 8 Then
        Borra
        FrmVerExisBodega.Show vbModal
    End If
End Sub
Private Sub Command3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 3
        NueMenBot
    End If
End Sub
Private Sub Command4_Click(Index As Integer)
    If Index = 1 Then
        Borra
        FrmCompAlm1.Show vbModal
    End If
    If Index = 2 Then
        Borra
        frmRequisiciones.Show vbModal
    End If
    If Index = 3 Then
        Borra
        frmVerCotizaciones.Show vbModal
    End If
    If Index = 4 Then
        Borra
        frmAutorizarCotizaciones.Show vbModal
    End If
    If Index = 5 Then
        Borra
        frmPreOrden.Show vbModal
    End If
    If Index = 6 Then
        Borra
        frmOrdenCompra.Show vbModal
    End If
    If Index = 7 Then
        Borra
        FrmOrdenRapida.Show vbModal
    End If
    If Index = 8 Then
        Borra
        FrmComiciones.Show vbModal
    End If
    If Index = 9 Then
        Borra
        frmAutOC.Show vbModal
    End If
    If Index = 10 Then
        Borra
        frmPagoOrden.Show vbModal
    End If
    If Index = 11 Then
        Borra
        frmOCPend.Show vbModal
    End If
    If Index = 12 Then
        Borra
        FrmRepJuegRep.Show vbModal
    End If
End Sub
Private Sub Command4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 4
        NueMenBot
    End If
End Sub
Private Sub Command5_Click(Index As Integer)
    If Index = 1 Then
        Borra
        frmAStec.Show vbModal
    End If
End Sub
Private Sub Command5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 5
        NueMenBot
    End If
End Sub
Private Sub Command6_Click(Index As Integer)
    If Index = 1 Then
        Borra
        frmReviComa.Show vbModal
    End If
    If Index = 2 Then
        Borra
        frmProduccion.Show vbModal
    End If
    If Index = 3 Then
        Borra
        frmCalidad.Show vbModal
    End If
    If Index = 4 Then
        Borra
        VerJuegoRep.Show vbModal
    End If
    If Index = 5 Then
        Borra
        EditarJRVarios.Show vbModal
    End If
    If Index = 6 Then
        Borra
        frmSalidaInvProd.Show vbModal
    End If
    If Index = 7 Then
        Borra
        FrmAcepAlmacen1.Show vbModal
    End If
    If Index = 8 Then
        Borra
        frmScrap.Show vbModal
    End If
End Sub
Private Sub Command6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 6
        NueMenBot
    End If
End Sub
Private Sub Command7_Click(Index As Integer)
    If Index = 1 Then
        Borra
        CambioPRe.Show vbModal
    End If
    If Index = 2 Then
        Borra
        BajaExcel.Show vbModal
    End If
    If Index = 3 Then
        Borra
        DarPerVenta.Show vbModal
    End If
    If Index = 4 Then
        Borra
        frmPerdidas.Show vbModal
    End If
    If Index = 5 Then
        Borra
        FrmArpvCompAlm1.Show vbModal
    End If
    If Index = 6 Then
        Borra
        FrmPagoCompAlm1.Show vbModal
    End If
    If Index = 7 Then
        Borra
        frmEmpresa.Show vbModal
    End If
    If Index = 8 Then
        Borra
        FrmValeCaja.Show vbModal
    End If
End Sub
Private Sub Command7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 7
        NueMenBot
    End If
End Sub
Private Sub Command8_Click(Index As Integer)
    If Index = 1 Then
        Borra
        frmPromos.Show vbModal
    End If
    If Index = 2 Then
        Borra
        FrmLicitacion.Show vbModal
    End If
    If Index = 3 Then
        Borra
        Marca.Show vbModal
    End If
    If Index = 4 Then
        Borra
        Dolar.Show vbModal
    End If
    If Index = 5 Then
        Borra
        FrmSustiInv.Show vbModal
    End If
    If Index = 6 Then
        Borra
        Reportes1.Show vbModal
    End If
End Sub
Private Sub Command8_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Validar = 8
        NueMenBot
    End If
End Sub
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
    If Hay_Mensajes(Me.Text1(0).Text) Then
        Me.fraMensajes.Visible = True
    Else
        Me.fraMensajes.Visible = False
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
    ValidaMenu
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & ". " & Err.Description & ".", vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Click()
    Borra
End Sub
Private Sub Form_Load()
On Error GoTo ManejaError
    Dim sValor As String
    Borra
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
Private Sub Form_Resize()
On Error GoTo ManejaError
    Label1.Height = Menu.Height
    Image7.Top = Menu.Height - (Image7.Height + Me.StatusBar1.Height)
    Image7.Left = Menu.Width - 4740
    '****************************** < CENTRAR LADO > *******************************
    lblHola.Left = (Menu.Width / 2) - (lblHola.Width / 2)
    Image1.Left = (Menu.Width / 2) - (Image1.Width / 2)
    Image8.Left = (Menu.Width / 2) - (Image8.Width / 2)
    lblPuestoSucursal.Left = (Menu.Width / 2) - (lblPuestoSucursal.Width / 2)
    fraMensajes.Left = (Menu.Width / 2) - (fraMensajes.Width / 2)
    lblEstado.Left = (Menu.Width / 2) - (lblEstado.Width / 2)
    '****************************** < CENTRAR ARRIBA > *******************************
    lblHola.Top = (Menu.Height / 2) - (5325 / 2)
    Image1.Top = (Menu.Height / 2) - ((5325 - 915) / 2)
    lblPuestoSucursal.Top = (Menu.Height / 2) - ((5325 - 5360) / 2)
    fraMensajes.Top = (Menu.Height / 2) + ((5325 - 3970) / 2)
    Image8.Top = (Menu.Height / 2) + ((5325 - 3970) / 2)
    lblEstado.Top = (Menu.Height / 2) + ((5325 - 7870) / 2)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & ". " & Err.Description & ".", vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub fraMensajes_Click()
    Borra
End Sub
Private Sub Image1_Click()
    Borra
End Sub
Private Sub Image2_Click()
    Borra
End Sub
Private Sub Image3_Click()
    Borra
End Sub
Private Sub Image6_Click()
    Borra
End Sub
Private Sub Image7_Click()
    Borra
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
Private Sub NueMenBot()
On Error GoTo ManejaError
    If Validar = 1 Then
        Me.Command1(1).Visible = True
        Me.Command1(4).Visible = True
        Me.Command1(6).Visible = True
        Me.Command1(2).Visible = True
        Me.Command3(8).Visible = True
    Else
        Me.Command1(1).Visible = False
        Me.Command1(4).Visible = False
        Me.Command1(6).Visible = False
        Me.Command1(2).Visible = False
        Me.Command3(8).Visible = False
    End If
    If Validar = 2 Then
        Me.Command2(1).Visible = True
        Me.Command2(2).Visible = True
        Me.Command2(3).Visible = True
        Me.Command2(4).Visible = True
        Me.Command2(5).Visible = True
        Me.Command2(6).Visible = True
        Me.Command2(8).Visible = True
    Else
        Me.Command2(1).Visible = False
        Me.Command2(2).Visible = False
        Me.Command2(3).Visible = False
        Me.Command2(4).Visible = False
        Me.Command2(5).Visible = False
        Me.Command2(6).Visible = False
        Me.Command2(8).Visible = False
    End If
    If Validar = 3 Then
        Me.Command3(1).Visible = True
        Me.Command3(7).Visible = True
    Else
        Me.Command3(1).Visible = False
        Me.Command3(7).Visible = False
    End If
    If Validar = 4 Then
        Me.Command4(2).Visible = True
        Me.Command4(3).Visible = True
        Me.Command4(4).Visible = True
        Me.Command4(5).Visible = True
    Else
        Me.Command4(2).Visible = False
        Me.Command4(3).Visible = False
        Me.Command4(4).Visible = False
        Me.Command4(5).Visible = False
    End If
    If Validar = 5 Then
        Me.Command5(1).Visible = True
    Else
        Me.Command5(1).Visible = False
    End If
    If Validar = 6 Then
        Me.Command6(1).Visible = True
        Me.Command6(2).Visible = True
        Me.Command6(3).Visible = True
    Else
        Me.Command6(1).Visible = False
        Me.Command6(2).Visible = False
        Me.Command6(3).Visible = False
    End If
    If Validar = 7 Then
        Me.Command7(7).Visible = True
    Else
        Me.Command7(7).Visible = False
    End If
    If Validar = 8 Then
        Me.Command8(6).Visible = True
        Me.Command7(2).Visible = True
        Me.Command4(8).Visible = True
        Me.Command4(12).Visible = True
        Me.Command12(6).Visible = True
    Else
        Me.Command8(6).Visible = False
        Me.Command7(2).Visible = False
        Me.Command4(8).Visible = False
        Me.Command4(12).Visible = False
        Me.Command12(6).Visible = False
    End If
    If Validar = 10 Then
        Me.Command10(1).Visible = True
        Me.Command10(2).Visible = True
        Me.Command10(3).Visible = True
    Else
        Me.Command10(1).Visible = False
        Me.Command10(2).Visible = False
        Me.Command10(3).Visible = False
    End If
    If Validar = 11 Then
        Me.Command11(1).Visible = True
        Me.Command11(4).Visible = True
        Me.Command11(5).Visible = True
        Me.Command11(6).Visible = True
    Else
        Me.Command11(1).Visible = False
        Me.Command11(4).Visible = False
        Me.Command11(5).Visible = False
        Me.Command11(6).Visible = False
    End If
    If Validar = 12 Then
        Me.Command12(1).Visible = True
        Me.Command12(2).Visible = True
        Me.Command12(3).Visible = True
        Me.Command11(2).Visible = True
        Me.Command15(2).Visible = True
    Else
        Me.Command12(1).Visible = False
        Me.Command12(2).Visible = False
        Me.Command12(3).Visible = False
        Me.Command11(2).Visible = False
        Me.Command15(2).Visible = False
    End If
    If Validar = 13 Then
        Me.Command13(1).Visible = True
        Me.Command13(2).Visible = True
        Me.Command13(3).Visible = True
        Me.Command13(4).Visible = True
        Me.Command13(5).Visible = True
        Me.Command13(6).Visible = True
        Me.Command13(7).Visible = True
        Me.Command8(3).Visible = True
    Else
        Me.Command13(1).Visible = False
        Me.Command13(2).Visible = False
        Me.Command13(3).Visible = False
        Me.Command13(4).Visible = False
        Me.Command13(5).Visible = False
        Me.Command13(6).Visible = False
        Me.Command13(7).Visible = False
        Me.Command8(3).Visible = False
    End If
    If Validar = 14 Then
        Me.Command14(1).Visible = True
        Me.Command14(2).Visible = True
        Me.Command14(3).Visible = True
        Me.Command14(4).Visible = True
        Me.Command14(5).Visible = True
        Me.Command14(6).Visible = True
        Me.Command14(7).Visible = True
    Else
        Me.Command14(1).Visible = False
        Me.Command14(2).Visible = False
        Me.Command14(3).Visible = False
        Me.Command14(4).Visible = False
        Me.Command14(5).Visible = False
        Me.Command14(6).Visible = False
        Me.Command14(7).Visible = False
    End If
    If Validar = 15 Then
        Me.Command15(1).Visible = True
        Me.Command15(3).Visible = True
        Me.Command15(4).Visible = True
    Else
        Me.Command15(1).Visible = False
        Me.Command15(3).Visible = False
        Me.Command15(4).Visible = False
    End If
    If Validar = 16 Then
        Me.Command1(3).Visible = True
        Me.Command8(2).Visible = True
        Me.Command8(1).Visible = True
        Me.Command1(7).Visible = True
        Me.Command7(3).Visible = True
        Me.Command7(1).Visible = True
        Me.Command7(8).Visible = True
    Else
        Me.Command1(3).Visible = False
        Me.Command8(2).Visible = False
        Me.Command8(1).Visible = False
        Me.Command1(7).Visible = False
        Me.Command7(3).Visible = False
        Me.Command7(1).Visible = False
        Me.Command7(8).Visible = False
    End If
    If Validar = 17 Then
        Me.Command3(2).Visible = True
        Me.Command3(3).Visible = True
        Me.Command3(4).Visible = True
        Me.Command3(5).Visible = True
        Me.Command3(6).Visible = True
        Me.Command12(5).Visible = True
        Me.Command7(4).Visible = True
        Me.Command8(5).Visible = True
        Me.Command12(4).Visible = True
    Else
        Me.Command3(2).Visible = False
        Me.Command3(3).Visible = False
        Me.Command3(4).Visible = False
        Me.Command3(5).Visible = False
        Me.Command3(6).Visible = False
        Me.Command12(5).Visible = False
        Me.Command7(4).Visible = False
        Me.Command8(5).Visible = False
        Me.Command12(4).Visible = False
    End If
    If Validar = 18 Then
        Me.Command4(9).Visible = True
        Me.Command4(6).Visible = True
        Me.Command4(7).Visible = True
        Me.Command4(11).Visible = True
    Else
        Me.Command4(9).Visible = False
        Me.Command4(6).Visible = False
        Me.Command4(7).Visible = False
        Me.Command4(11).Visible = False
    End If
    If Validar = 19 Then
        Me.Command4(1).Visible = True
        Me.Command6(7).Visible = True
        Me.Command7(5).Visible = True
    Else
        Me.Command4(1).Visible = False
        Me.Command6(7).Visible = False
        Me.Command7(5).Visible = False
    End If
    If Validar = 20 Then
        Me.Command6(8).Visible = True
        Me.Command6(6).Visible = True
    Else
        Me.Command6(8).Visible = False
        Me.Command6(6).Visible = False
    End If
    If Validar = 21 Then
        Me.Command13(8).Visible = True
        Me.Command6(5).Visible = True
        Me.Command6(4).Visible = True
    Else
        Me.Command13(8).Visible = False
        Me.Command6(5).Visible = False
        Me.Command6(4).Visible = False
    End If
    If Validar = 22 Then
        Me.Command2(7).Visible = True
    Else
        Me.Command2(7).Visible = False
    End If
    If Validar = 23 Then
        Me.Command4(10).Visible = True
        Me.Command1(5).Visible = True
        Me.Command11(3).Visible = True
        Me.Command7(6).Visible = True
        Me.Command8(4).Visible = True
        Me.Command1(8).Visible = True
    Else
        Me.Command4(10).Visible = False
        Me.Command1(5).Visible = False
        Me.Command11(3).Visible = False
        Me.Command7(6).Visible = False
        Me.Command8(4).Visible = False
        Me.Command1(8).Visible = False
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & ". " & Err.Description & ".", vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image8_Click()
    Borra
    MsgAPToner.Show
End Sub
Private Sub imgLeer_Click()
    Borra
    MsgAPToner.Show
End Sub
Private Sub Label1_Click()
    Borra
End Sub
Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 7
    SelMen
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 1
    SelMen
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 2
    SelMen
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 3
    SelMen
End Sub
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 4
    SelMen
End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 5
    SelMen
End Sub
Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 5
    SelMen
End Sub
Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 6
    SelMen
End Sub
Private Sub lblEstado_Click()
    Borra
End Sub
Private Sub lblHola_Click()
    Borra
End Sub
Private Sub lblPuestoSucursal_Click()
    Borra
End Sub
Private Sub ValidaMenu()
    If Text1(12).Text = "N" And Text1(56).Text = "N" And Text1(6).Text = "N" And Text1(7).Text = "N" Then
        Me.Command1(1).Enabled = False
    End If
    If Text1(8).Text = "N" Then
        Me.Command1(2).Enabled = False
    End If
    If Text1(13).Text = "N" Then
        Me.Command1(3).Enabled = False
        Me.Command7(8).Enabled = False
    End If
    If Text1(49).Text = "N" And Text1(50).Text = "N" Then
        Me.Command1(4).Enabled = False
    End If
    If Text1(57).Text = "N" Then
        Me.Command1(5).Enabled = False
    End If
    If Text1(48).Text = "N" Then
        Me.Command1(6).Enabled = False
    End If
    If Text1(10).Text = "N" Then
        Me.Command11(1).Enabled = False
    End If
    If Text1(11).Text = "N" Then
        'Me.Command11(2).Enabled = False
    End If
    If Text1(9).Text = "N" Then
        Me.Command11(3).Enabled = False
        Me.Command1(8).Enabled = False
    End If
    If Text1(64).Text = "N" Then
        Me.Command11(4).Enabled = False
        Me.Command11(5).Enabled = False
    End If
    If Text1(27).Text = "N" Then
        Me.Command2(1).Enabled = False
        Me.Command4(1).Enabled = False
    End If
    If Text1(17).Text = "N" Then
        Me.Command2(2).Enabled = False
        Me.Command3(7).Enabled = False
    End If
    If Text1(51).Text = "N" Then
        Me.Command2(3).Enabled = False
    End If
    If Text1(52).Text = "N" Then
        Me.Command2(4).Enabled = False
    End If
    If Text1(15).Text = "N" Then
        Me.Command2(5).Enabled = False
    End If
    If Text1(14).Text = "N" Then
        Me.Command2(6).Enabled = False
    End If
    If Text1(56).Text = "N" Then
        Me.Command2(7).Enabled = False
    End If
    If Text1(24).Text = "N" Then
        Me.Command2(8).Enabled = False
    End If
    If Text1(26).Text = "N" Then
        Me.Command3(2).Enabled = False
    End If
    If Text1(46).Text = "N" Then
        Me.Command3(3).Enabled = False
    End If
    If Text1(58).Text = "N" Then
        Me.Command3(4).Enabled = False
        Me.Command11(2).Enabled = False
        Me.Command15(2).Enabled = False
    End If
    If Text1(53).Text = "N" Then
        Me.Command3(5).Enabled = False
    End If
    If Text1(55).Text = "N" Then
        Me.Command3(6).Enabled = False
    End If
    If Text1(19).Text = "N" Or Text1(20).Text = "N" Then
        Me.Command12(1).Enabled = False
    End If
    If Text1(18).Text = "N" Then
        Me.Command12(2).Enabled = False
    End If
    If Text1(23).Text = "N" Then
        Me.Command12(3).Enabled = False
    End If
    If Text1(63).Text = "N" Then
        Me.Command2(1).Enabled = False
        Me.Command4(2).Enabled = False
        Me.Command4(8).Enabled = False
    End If
    If Text1(59).Text = "N" Then
        Me.Command4(3).Enabled = False
    End If
    If Text1(60).Text = "N" Then
        Me.Command4(4).Enabled = False
        'Me.Command4(9).Enabled = False
    End If
    If Text1(61).Text = "N" Then
        Me.Command4(5).Enabled = False
    End If
    If Text1(62).Text = "N" Then
        Me.Command4(6).Enabled = False
    End If
    If Text1(28).Text = "N" Then
        Me.Command5(1).Enabled = False
    End If
    If Text1(39).Text = "N" Then
        Me.Command6(1).Enabled = False
    End If
    If Text1(40).Text = "N" Then
        Me.Command6(2).Enabled = False
    End If
    If Text1(41).Text = "N" Then
        Me.Command6(3).Enabled = False
    End If
    If Text1(42).Text = "N" Then
        Me.Command6(4).Enabled = False
        Me.Command4(12).Enabled = False
    End If
    If Text1(36).Text = "N" Then
        Me.Command7(1).Enabled = False
    End If
    If Text1(37).Text = "N" Then
        Me.Command7(2).Enabled = False
    End If
    If Text1(47).Text = "N" Then
        Me.Command7(3).Enabled = False
        Me.Command7(7).Enabled = False
        Me.Command4(9).Enabled = False
    End If
    If Text1(29).Text = "N" Then
        Me.Command13(1).Enabled = False
        Me.Command13(5).Enabled = False
    End If
    If Text1(30).Text = "N" Then
        Me.Command13(2).Enabled = False
    End If
    If Text1(31).Text = "N" Then
        Me.Command13(3).Enabled = False
    End If
    If Text1(32).Text = "N" Then
        Me.Command13(4).Enabled = False
    End If
    If Text1(22).Text = "N" Then
        Me.Command13(6).Enabled = False
        Me.Command14(6).Enabled = False
    End If
    If Text1(21).Text = "N" Then
        Me.Command13(7).Enabled = False
        Me.Command14(7).Enabled = False
    End If
    If Text1(33).Text = "N" Then
        Me.Command14(1).Enabled = False
        Me.Command14(5).Enabled = False
    End If
    If Text1(34).Text = "N" Then
        Me.Command14(2).Enabled = False
        Me.Command14(3).Enabled = False
        Me.Command14(4).Enabled = False
    End If
    If Text1(43).Text = "N" Then
        Me.Command8(1).Enabled = False
        Me.Command8(2).Enabled = False
    End If
    If Text1(44).Text = "N" Then
        Me.Command8(3).Enabled = False
    End If
    If Text1(45).Text = "N" Then
        Me.Command8(4).Enabled = False
    End If
    If Text1(16).Text = "N" Then
        Me.Command15(0).Enabled = False
    End If
    If Text1(38).Text = "N" Then
        Me.Command13(8).Enabled = False
    End If
    If Text1(54).Text = "N" Then
        Me.Command3(8).Enabled = False
    End If
    If Text1(36).Text = "N" Or Text1(37).Text = "N" Or Text1(47).Text = "N" Then
        Me.Command7(4).Enabled = False
    End If
    If Text1(17).Text = "N" Then
        Me.Command15(2).Enabled = False
    End If
    If Text1(65).Text = "N" Then
        Me.Command6(5).Enabled = False
    End If
    If Text1(66).Text = "N" Then
        Me.Command6(6).Enabled = False
        Me.Command6(8).Enabled = False
    End If
    If Text1(67).Text = "N" Then
        Me.Command1(7).Enabled = False
    End If
    If Text1(68).Text = "N" Then
        Me.Command4(1).Enabled = False
        Me.Command4(7).Enabled = False
        Me.Command4(11).Enabled = False
    End If
    If Text1(69).Text = "N" Then
        Me.Command6(7).Enabled = False
    End If
    If Text1(70).Text = "N" Then
        Me.Command7(5).Enabled = False
    End If
    If Text1(71).Text = "N" Then
        Me.Command7(6).Enabled = False
        Me.Command4(10).Enabled = False
    End If
    If Text1(72).Text = "N" Then
        Me.Command12(5).Enabled = False
    End If
    If Text1(73).Text = "N" Then
        Me.Command11(6).Enabled = False
    End If
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 1
    SelMen
End Sub
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 2
    SelMen
End Sub
Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 3
    SelMen
End Sub
Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 4
    SelMen
End Sub
Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 5
    SelMen
End Sub
Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 6
    SelMen
End Sub
Private Sub Picture7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SubMen = 7
    SelMen
End Sub
Private Sub SelMen()
    BorraDos
    If SubMen = 1 Then
        Command1(0).Visible = True
        Command11(0).Visible = True
        Command16(0).Visible = True
    Else
        Command1(0).Visible = False
        Command11(0).Visible = False
        Command16(0).Visible = False
    End If
    If SubMen = 2 Then
        Command3(0).Visible = True
        Command17(0).Visible = True
        Command12(0).Visible = True
    Else
        Command3(0).Visible = False
        Command17(0).Visible = False
        Command12(0).Visible = False
    End If
    If SubMen = 3 Then
        Command4(0).Visible = True
        Command18(0).Visible = True
        Command19(0).Visible = True
    Else
        Command4(0).Visible = False
        Command18(0).Visible = False
        Command19(0).Visible = False
    End If
    If SubMen = 4 Then
        Command6(0).Visible = True
        Command20(0).Visible = True
        Command21(0).Visible = True
    Else
        Command6(0).Visible = False
        Command20(0).Visible = False
        Command21(0).Visible = False
    End If
    If SubMen = 5 Then
        Command5(0).Visible = True
        Command22(0).Visible = True
        Command23(0).Visible = True
    Else
        Command5(0).Visible = False
        Command22(0).Visible = False
        Command23(0).Visible = False
    End If
    If SubMen = 6 Then
        Command13(0).Visible = True
        Command14(0).Visible = True
        Command7(0).Visible = True
    Else
        Command13(0).Visible = False
        Command14(0).Visible = False
        Command7(0).Visible = False
    End If
    If SubMen = 7 Then
        Command2(0).Visible = True
        Command15(0).Visible = True
        Command8(0).Visible = True
    Else
        Command2(0).Visible = False
        Command15(0).Visible = False
        Command8(0).Visible = False
    End If
End Sub
Private Sub Borra()
    Command1(0).Visible = False
    Command2(0).Visible = False
    Command3(0).Visible = False
    Command4(0).Visible = False
    Command5(0).Visible = False
    Command6(0).Visible = False
    Command7(0).Visible = False
    Command8(0).Visible = False
    Command11(0).Visible = False
    Command12(0).Visible = False
    Command13(0).Visible = False
    Command14(0).Visible = False
    Command15(0).Visible = False
    Command16(0).Visible = False
    Command17(0).Visible = False
    Command18(0).Visible = False
    Command19(0).Visible = False
    Command20(0).Visible = False
    Command21(0).Visible = False
    Command22(0).Visible = False
    Command23(0).Visible = False
    
    Me.Command1(1).Visible = False
    Me.Command1(2).Visible = False
    Me.Command1(3).Visible = False
    Me.Command1(4).Visible = False
    Me.Command1(5).Visible = False
    Me.Command1(6).Visible = False
    Me.Command1(7).Visible = False
    Me.Command1(8).Visible = False
    
    Me.Command2(1).Visible = False
    Me.Command2(2).Visible = False
    Me.Command2(3).Visible = False
    Me.Command2(4).Visible = False
    Me.Command2(5).Visible = False
    Me.Command2(6).Visible = False
    Me.Command2(7).Visible = False
    Me.Command2(8).Visible = False
    
    Me.Command3(1).Visible = False
    Me.Command3(2).Visible = False
    Me.Command3(3).Visible = False
    Me.Command3(4).Visible = False
    Me.Command3(5).Visible = False
    Me.Command3(6).Visible = False
    Me.Command3(7).Visible = False
    Me.Command3(8).Visible = False
    
    Me.Command4(1).Visible = False
    Me.Command4(2).Visible = False
    Me.Command4(3).Visible = False
    Me.Command4(4).Visible = False
    Me.Command4(5).Visible = False
    Me.Command4(6).Visible = False
    Me.Command4(7).Visible = False
    Me.Command4(8).Visible = False
    Me.Command4(9).Visible = False
    Me.Command4(10).Visible = False
    Me.Command4(11).Visible = False
    Me.Command4(12).Visible = False
    
    Me.Command5(1).Visible = False
    
    Me.Command6(1).Visible = False
    Me.Command6(2).Visible = False
    Me.Command6(3).Visible = False
    Me.Command6(4).Visible = False
    Me.Command6(5).Visible = False
    Me.Command6(6).Visible = False
    Me.Command6(7).Visible = False
    Me.Command6(8).Visible = False
    
    Me.Command7(1).Visible = False
    Me.Command7(2).Visible = False
    Me.Command7(3).Visible = False
    Me.Command7(4).Visible = False
    Me.Command7(5).Visible = False
    Me.Command7(6).Visible = False
    Me.Command7(7).Visible = False
    Me.Command7(8).Visible = False
    
    Me.Command8(1).Visible = False
    Me.Command8(2).Visible = False
    Me.Command8(3).Visible = False
    Me.Command8(4).Visible = False
    Me.Command8(5).Visible = False
    Me.Command8(6).Visible = False
    
    Me.Command10(1).Visible = False
    Me.Command10(2).Visible = False
    Me.Command10(3).Visible = False
    
    Me.Command11(1).Visible = False
    Me.Command11(2).Visible = False
    Me.Command11(3).Visible = False
    Me.Command11(4).Visible = False
    Me.Command11(5).Visible = False
    Me.Command11(6).Visible = False
    
    Me.Command12(1).Visible = False
    Me.Command12(2).Visible = False
    Me.Command12(3).Visible = False
    Me.Command12(4).Visible = False
    Me.Command12(5).Visible = False
    Me.Command12(6).Visible = False
    
    Me.Command13(1).Visible = False
    Me.Command13(2).Visible = False
    Me.Command13(3).Visible = False
    Me.Command13(4).Visible = False
    Me.Command13(5).Visible = False
    Me.Command13(6).Visible = False
    Me.Command13(7).Visible = False
    Me.Command13(8).Visible = False
    
    Me.Command14(1).Visible = False
    Me.Command14(2).Visible = False
    Me.Command14(3).Visible = False
    Me.Command14(4).Visible = False
    Me.Command14(5).Visible = False
    Me.Command14(6).Visible = False
    Me.Command14(7).Visible = False
    
    Me.Command15(1).Visible = False
    Me.Command15(2).Visible = False
    Me.Command15(3).Visible = False
    Me.Command15(4).Visible = False
End Sub
Private Sub BorraDos()
    Me.Command1(1).Visible = False
    Me.Command1(2).Visible = False
    Me.Command1(3).Visible = False
    Me.Command1(4).Visible = False
    Me.Command1(5).Visible = False
    Me.Command1(6).Visible = False
    Me.Command1(7).Visible = False
    Me.Command1(8).Visible = False
    
    Me.Command2(1).Visible = False
    Me.Command2(2).Visible = False
    Me.Command2(3).Visible = False
    Me.Command2(4).Visible = False
    Me.Command2(5).Visible = False
    Me.Command2(6).Visible = False
    Me.Command2(7).Visible = False
    Me.Command2(8).Visible = False
    
    Me.Command3(1).Visible = False
    Me.Command3(2).Visible = False
    Me.Command3(3).Visible = False
    Me.Command3(4).Visible = False
    Me.Command3(5).Visible = False
    Me.Command3(6).Visible = False
    Me.Command3(7).Visible = False
    Me.Command3(8).Visible = False
    
    Me.Command4(1).Visible = False
    Me.Command4(2).Visible = False
    Me.Command4(3).Visible = False
    Me.Command4(4).Visible = False
    Me.Command4(5).Visible = False
    Me.Command4(6).Visible = False
    Me.Command4(7).Visible = False
    Me.Command4(8).Visible = False
    Me.Command4(9).Visible = False
    Me.Command4(10).Visible = False
    Me.Command4(11).Visible = False
    Me.Command4(12).Visible = False
    
    Me.Command5(1).Visible = False
    
    Me.Command6(1).Visible = False
    Me.Command6(2).Visible = False
    Me.Command6(3).Visible = False
    Me.Command6(4).Visible = False
    Me.Command6(5).Visible = False
    Me.Command6(6).Visible = False
    Me.Command6(7).Visible = False
    Me.Command6(8).Visible = False
    
    Me.Command7(1).Visible = False
    Me.Command7(2).Visible = False
    Me.Command7(3).Visible = False
    Me.Command7(4).Visible = False
    Me.Command7(5).Visible = False
    Me.Command7(6).Visible = False
    Me.Command7(7).Visible = False
    Me.Command7(8).Visible = False
    
    Me.Command8(1).Visible = False
    Me.Command8(2).Visible = False
    Me.Command8(3).Visible = False
    Me.Command8(4).Visible = False
    Me.Command8(5).Visible = False
    Me.Command8(6).Visible = False
    
    
    Me.Command10(1).Visible = False
    Me.Command10(2).Visible = False
    Me.Command10(3).Visible = False
    
    Me.Command11(1).Visible = False
    Me.Command11(2).Visible = False
    Me.Command11(3).Visible = False
    Me.Command11(4).Visible = False
    Me.Command11(5).Visible = False
    Me.Command11(6).Visible = False
    
    Me.Command12(1).Visible = False
    Me.Command12(2).Visible = False
    Me.Command12(3).Visible = False
    Me.Command12(4).Visible = False
    Me.Command12(5).Visible = False
    Me.Command12(6).Visible = False
    
    Me.Command13(1).Visible = False
    Me.Command13(2).Visible = False
    Me.Command13(3).Visible = False
    Me.Command13(4).Visible = False
    Me.Command13(5).Visible = False
    Me.Command13(6).Visible = False
    Me.Command13(7).Visible = False
    Me.Command13(8).Visible = False
    
    Me.Command14(1).Visible = False
    Me.Command14(2).Visible = False
    Me.Command14(3).Visible = False
    Me.Command14(4).Visible = False
    Me.Command14(5).Visible = False
    Me.Command14(6).Visible = False
    Me.Command14(7).Visible = False
    
    Me.Command15(1).Visible = False
    Me.Command15(2).Visible = False
    Me.Command15(3).Visible = False
    Me.Command15(4).Visible = False
End Sub

