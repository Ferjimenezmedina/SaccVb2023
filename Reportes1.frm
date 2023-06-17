VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Reportes1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "REPORTES"
   ClientHeight    =   9375
   ClientLeft      =   3450
   ClientTop       =   960
   ClientWidth     =   11895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   11895
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10800
      TabIndex        =   131
      Top             =   5640
      Width           =   975
      Begin VB.Image Image2 
         Height          =   675
         Left            =   120
         MouseIcon       =   "Reportes1.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Reportes1.frx":030A
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ventas"
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
         TabIndex        =   132
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame33 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10800
      TabIndex        =   120
      Top             =   8040
      Width           =   975
      Begin VB.Image cmdCancelar 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Reportes1.frx":1EB8
         MousePointer    =   99  'Custom
         Picture         =   "Reportes1.frx":21C2
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Salir"
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
         TabIndex        =   121
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame32 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10800
      TabIndex        =   118
      Top             =   6840
      Width           =   975
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reporte"
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
         TabIndex        =   119
         Top             =   960
         Width           =   975
      End
      Begin VB.Image rpt 
         Height          =   675
         Left            =   120
         MouseIcon       =   "Reportes1.frx":42A4
         MousePointer    =   99  'Custom
         Picture         =   "Reportes1.frx":45AE
         Top             =   240
         Width           =   660
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   14631
      _Version        =   393216
      TabOrientation  =   2
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "DE VENTAS"
      TabPicture(0)   =   "Reportes1.frx":5D24
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame26"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "DE ALMACEN"
      TabPicture(1)   =   "Reportes1.frx":5D40
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame19"
      Tab(1).Control(1)=   "Frame14"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "DE COMPRAS"
      TabPicture(2)   =   "Reportes1.frx":5D5C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame29"
      Tab(2).Control(1)=   "Frame17"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame29 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   105
         Top             =   120
         Width           =   9975
         Begin VB.OptionButton Option7 
            Caption         =   "Factura"
            Height          =   195
            Left            =   6360
            TabIndex        =   134
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            Caption         =   "No. Envio"
            Height          =   195
            Left            =   5160
            TabIndex        =   133
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Clave"
            Height          =   195
            Index           =   0
            Left            =   4320
            TabIndex        =   116
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   1
            Left            =   3000
            TabIndex        =   115
            Top             =   240
            Width           =   1215
         End
         Begin VB.Frame Frame31 
            Height          =   2655
            Left            =   120
            TabIndex        =   111
            Top             =   840
            Width           =   8175
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   120
               MaxLength       =   50
               PasswordChar    =   "*"
               TabIndex        =   112
               Top             =   2040
               Visible         =   0   'False
               Width           =   375
            End
            Begin MSComctlLib.ListView ListView6 
               Height          =   2295
               Left            =   120
               TabIndex        =   113
               Top             =   240
               Width           =   7935
               _ExtentX        =   13996
               _ExtentY        =   4048
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   16777215
               Appearance      =   0
               NumItems        =   0
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "ESPERE MIENTRAS SE LLENA LA LISTA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   720
               TabIndex        =   114
               Top             =   600
               Visible         =   0   'False
               Width           =   1935
            End
         End
         Begin VB.Frame Frame30 
            Caption         =   "Busqueda"
            Height          =   2655
            Left            =   8400
            TabIndex        =   107
            Top             =   780
            Width           =   1455
            Begin VB.CheckBox Check16 
               Caption         =   "Almacen 1"
               Height          =   255
               Left            =   240
               TabIndex        =   110
               Top             =   360
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox Check15 
               Caption         =   "Almacen 2"
               Height          =   255
               Left            =   240
               TabIndex        =   109
               Top             =   600
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox Check14 
               Caption         =   "Almacen 3"
               Height          =   255
               Left            =   240
               TabIndex        =   108
               Top             =   840
               Value           =   1  'Checked
               Width           =   1095
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            CausesValidation=   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   106
            Top             =   480
            Width           =   8055
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Producto"
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
            Left            =   120
            TabIndex        =   117
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame26 
         Height          =   2655
         Left            =   5760
         TabIndex        =   77
         Top             =   5400
         Width           =   4695
         Begin VB.Frame Frame9 
            Caption         =   "Tipo de Pago"
            Height          =   2055
            Left            =   2520
            TabIndex        =   122
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
            Begin VB.OptionButton Option19 
               Caption         =   "Pagos de Contado"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   128
               Top             =   1380
               Width           =   1815
            End
            Begin VB.OptionButton Option19 
               Caption         =   "Credito"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   127
               Top             =   1125
               Width           =   1335
            End
            Begin VB.OptionButton Option19 
               Caption         =   "Tarjeta de Credito"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   126
               Top             =   870
               Width           =   1575
            End
            Begin VB.OptionButton Option19 
               Caption         =   "Cheque"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   125
               Top             =   615
               Width           =   1335
            End
            Begin VB.OptionButton Option19 
               Caption         =   "Efectivo"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   124
               Top             =   360
               Width           =   1335
            End
            Begin VB.OptionButton Option19 
               Caption         =   "Todos los Pagos"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   123
               Top             =   1635
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin VB.Frame Frame27 
            Caption         =   "Tipo de Orden"
            Height          =   2055
            Left            =   120
            TabIndex        =   78
            Top             =   360
            Width           =   2295
            Begin VB.OptionButton Option9 
               Caption         =   "Por Sucursales"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   84
               Top             =   1635
               Value           =   -1  'True
               Width           =   1575
            End
            Begin VB.OptionButton Option9 
               Caption         =   "Por Cliente"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   83
               Top             =   360
               Width           =   1335
            End
            Begin VB.OptionButton Option9 
               Caption         =   "Por Producto"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   82
               Top             =   615
               Width           =   1335
            End
            Begin VB.OptionButton Option9 
               Caption         =   "Por Agente"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   81
               Top             =   870
               Width           =   1455
            End
            Begin VB.OptionButton Option9 
               Caption         =   "Por Fechas"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   80
               Top             =   1125
               Width           =   1335
            End
            Begin VB.OptionButton Option9 
               Caption         =   "Por Folio de Facturas"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   79
               Top             =   1380
               Width           =   1815
            End
         End
      End
      Begin VB.Frame Frame17 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   68
         Top             =   3720
         Width           =   9975
         Begin VB.Frame Frame23 
            Caption         =   "Estado de la Orden de Compra"
            Height          =   1575
            Left            =   6960
            TabIndex        =   97
            Top             =   1560
            Visible         =   0   'False
            Width           =   2775
            Begin VB.CheckBox Check13 
               Caption         =   "Pendientes de Pagar"
               Height          =   255
               Left            =   240
               TabIndex        =   99
               Top             =   240
               Value           =   1  'Checked
               Width           =   2055
            End
            Begin VB.CheckBox Check12 
               Caption         =   "Pagadas"
               Height          =   255
               Left            =   240
               TabIndex        =   98
               Top             =   480
               Value           =   1  'Checked
               Width           =   2055
            End
         End
         Begin VB.Frame Frame25 
            Caption         =   "Surtidas"
            Height          =   1575
            Left            =   7320
            TabIndex        =   100
            Top             =   1800
            Visible         =   0   'False
            Width           =   2535
            Begin VB.OptionButton Option13 
               Caption         =   "No importan"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   104
               Top             =   960
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton Option13 
               Caption         =   "Parcial"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   103
               Top             =   480
               Width           =   1575
            End
            Begin VB.OptionButton Option13 
               Caption         =   "Total"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   102
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton Option13 
               Caption         =   "Sin Surtir"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   101
               Top             =   720
               Width           =   1455
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Busqueda"
            Height          =   975
            Left            =   4440
            TabIndex        =   94
            Top             =   720
            Width           =   5415
            Begin VB.OptionButton Option12 
               Caption         =   "No Importan los Productos Recibidos"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   96
               Top             =   600
               Value           =   -1  'True
               Width           =   3255
            End
            Begin VB.OptionButton Option12 
               Caption         =   "Por Productos Recibidos"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   95
               Top             =   360
               Width           =   2775
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            CausesValidation=   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   69
            Top             =   480
            Width           =   4095
         End
         Begin VB.Frame Frame18 
            Height          =   2775
            Left            =   120
            TabIndex        =   70
            Top             =   720
            Width           =   4215
            Begin MSComctlLib.ListView ListView4 
               Height          =   2415
               Left            =   120
               TabIndex        =   71
               Top             =   240
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   4260
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   16777215
               Appearance      =   0
               NumItems        =   0
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "ESPERE MIENTRAS SE LLENA LA LISTA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   720
               TabIndex        =   72
               Top             =   600
               Visible         =   0   'False
               Width           =   1935
            End
         End
         Begin VB.Frame Frame22 
            Caption         =   "Estado de la Orden de Compra"
            Height          =   1575
            Left            =   4440
            TabIndex        =   73
            Top             =   1800
            Width           =   2775
            Begin VB.CheckBox Check11 
               Caption         =   "Pagadas"
               Height          =   255
               Left            =   240
               TabIndex        =   93
               Top             =   1080
               Value           =   1  'Checked
               Width           =   1815
            End
            Begin VB.CheckBox Check10 
               Caption         =   "Pendientes de Pagar"
               Height          =   255
               Left            =   240
               TabIndex        =   92
               Top             =   840
               Value           =   1  'Checked
               Width           =   2055
            End
            Begin VB.CheckBox Check9 
               Caption         =   "Pendientes de Guardar"
               Height          =   255
               Left            =   240
               TabIndex        =   75
               Top             =   600
               Value           =   1  'Checked
               Width           =   2055
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Pendientes de Autorizar"
               Height          =   255
               Left            =   240
               TabIndex        =   74
               Top             =   360
               Value           =   1  'Checked
               Width           =   2055
            End
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor"
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
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame19 
         Height          =   3735
         Left            =   -74460
         TabIndex        =   57
         Top             =   3720
         Width           =   9735
         Begin VB.Frame Frame24 
            Caption         =   "Tipo de Existencia"
            Height          =   1935
            Left            =   7080
            TabIndex        =   86
            Top             =   645
            Width           =   2535
            Begin VB.OptionButton Option6 
               Caption         =   "Menor al minimo"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   91
               Top             =   360
               Width           =   2295
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Entre el minimo y el maximo"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   90
               Top             =   600
               Width           =   2295
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Mayor al maximo"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   89
               Top             =   840
               Width           =   2175
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Mayor que 0"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   88
               Top             =   1080
               Width           =   1575
            End
            Begin VB.OptionButton Option6 
               Caption         =   "No importa la existencia"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   87
               Top             =   1320
               Value           =   -1  'True
               Width           =   2295
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            CausesValidation=   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   4095
         End
         Begin VB.Frame Frame21 
            Height          =   3015
            Left            =   120
            TabIndex        =   65
            Top             =   600
            Width           =   4215
            Begin MSComctlLib.ListView ListView5 
               Height          =   2655
               Left            =   120
               TabIndex        =   66
               Top             =   240
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   4683
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   16777215
               Appearance      =   0
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Tipo de Movimiento"
            Height          =   1935
            Left            =   4440
            TabIndex        =   58
            Top             =   645
            Width           =   2535
            Begin VB.OptionButton Option10 
               Caption         =   "Existencias"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   85
               Top             =   1560
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton Option10 
               Caption         =   "Producciones"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   63
               Top             =   360
               Width           =   1695
            End
            Begin VB.OptionButton Option10 
               Caption         =   "Scrap"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   62
               Top             =   600
               Width           =   1815
            End
            Begin VB.OptionButton Option10 
               Caption         =   "Requisiciones"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   61
               Top             =   840
               Width           =   1695
            End
            Begin VB.OptionButton Option10 
               Caption         =   "Entradas"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   60
               Top             =   1080
               Width           =   1575
            End
            Begin VB.OptionButton Option10 
               Caption         =   "Pedidos"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   59
               Top             =   1320
               Width           =   1815
            End
         End
         Begin VB.Label Label13 
            Caption         =   "Sucursal"
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
            Left            =   240
            TabIndex        =   67
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame Frame14 
         Height          =   3495
         Left            =   -74460
         TabIndex        =   44
         Top             =   120
         Width           =   9735
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            CausesValidation=   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   7215
         End
         Begin VB.Frame Frame16 
            Caption         =   "Busqueda"
            Height          =   2535
            Left            =   7440
            TabIndex        =   52
            Top             =   720
            Width           =   2175
            Begin VB.CheckBox Check7 
               Caption         =   "Almacen 3"
               Height          =   255
               Left            =   240
               TabIndex        =   55
               Top             =   840
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox Check6 
               Caption         =   "Almacen 2"
               Height          =   255
               Left            =   240
               TabIndex        =   54
               Top             =   600
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Almacen 1"
               Height          =   255
               Left            =   240
               TabIndex        =   53
               Top             =   360
               Value           =   1  'Checked
               Width           =   1095
            End
         End
         Begin VB.Frame Frame15 
            Height          =   2655
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   7215
            Begin VB.TextBox Text3 
               Alignment       =   2  'Center
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   120
               MaxLength       =   50
               PasswordChar    =   "*"
               TabIndex        =   49
               Top             =   960
               Visible         =   0   'False
               Width           =   375
            End
            Begin MSComctlLib.ListView ListView3 
               Height          =   2295
               Left            =   120
               TabIndex        =   50
               Top             =   240
               Width           =   6975
               _ExtentX        =   12303
               _ExtentY        =   4048
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   16777215
               Appearance      =   0
               NumItems        =   0
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "ESPERE MIENTRAS SE LLENA LA LISTA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   720
               TabIndex        =   51
               Top             =   600
               Visible         =   0   'False
               Width           =   1935
            End
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   0
            Left            =   5160
            TabIndex        =   47
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Clave"
            Height          =   195
            Left            =   6480
            TabIndex        =   46
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Producto"
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
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2535
         Left            =   480
         TabIndex        =   32
         Top             =   2880
         Width           =   9975
         Begin VB.OptionButton Option1 
            Caption         =   "Clave"
            Height          =   195
            Index           =   1
            Left            =   5880
            TabIndex        =   130
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Descripcion"
            Height          =   195
            Index           =   2
            Left            =   4320
            TabIndex        =   129
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            CausesValidation=   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   480
            Width           =   7455
         End
         Begin VB.Frame Frame1 
            Height          =   1695
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   7695
            Begin VB.TextBox txtID_User 
               Alignment       =   2  'Center
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   120
               MaxLength       =   50
               PasswordChar    =   "*"
               TabIndex        =   40
               Top             =   1200
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox txtID_Prov 
               Alignment       =   2  'Center
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   600
               MaxLength       =   50
               PasswordChar    =   "*"
               TabIndex        =   39
               Top             =   1200
               Visible         =   0   'False
               Width           =   375
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   1335
               Left            =   120
               TabIndex        =   41
               Top             =   240
               Width           =   7335
               _ExtentX        =   12938
               _ExtentY        =   2355
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   16777215
               Appearance      =   0
               NumItems        =   0
            End
            Begin VB.Label lblEspere 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "ESPERE MIENTRAS SE LLENA LA LISTA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   720
               TabIndex        =   42
               Top             =   600
               Visible         =   0   'False
               Width           =   1935
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Busqueda"
            Height          =   1695
            Left            =   7920
            TabIndex        =   33
            Top             =   720
            Width           =   1935
            Begin VB.CheckBox Check3 
               Caption         =   "Almacen 1"
               Height          =   255
               Left            =   240
               TabIndex        =   36
               Top             =   360
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox Check4 
               Caption         =   "Almacen 2"
               Height          =   255
               Left            =   240
               TabIndex        =   35
               Top             =   600
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox Check5 
               Caption         =   "Almacen 3"
               Height          =   255
               Left            =   240
               TabIndex        =   34
               Top             =   840
               Value           =   1  'Checked
               Width           =   1095
            End
         End
         Begin VB.Label lblProd 
            BackStyle       =   0  'Transparent
            Caption         =   "Producto"
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
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2655
         Left            =   480
         TabIndex        =   25
         Top             =   5400
         Width           =   5175
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            CausesValidation=   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   4935
         End
         Begin VB.Frame Frame2 
            Height          =   1575
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   4935
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   120
               MaxLength       =   50
               PasswordChar    =   "*"
               TabIndex        =   29
               Top             =   840
               Visible         =   0   'False
               Width           =   375
            End
            Begin MSComctlLib.ListView ListView2 
               Height          =   1215
               Left            =   120
               TabIndex        =   30
               Top             =   240
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   2143
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   16777215
               Appearance      =   0
               NumItems        =   0
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Agrupar por Producto"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Agente"
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
            Left            =   240
            TabIndex        =   31
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2895
         Left            =   480
         TabIndex        =   6
         Top             =   60
         Width           =   9975
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            CausesValidation=   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   4695
         End
         Begin VB.Frame Frame5 
            Height          =   2175
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   4815
            Begin MSComctlLib.ListView ListaDocumentos 
               Height          =   1695
               Left            =   120
               TabIndex        =   23
               Top             =   240
               Width           =   4575
               _ExtentX        =   8070
               _ExtentY        =   2990
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   16777215
               Appearance      =   0
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Tipo de Movimiento"
            Height          =   1695
            Left            =   5040
            TabIndex        =   15
            Top             =   240
            Width           =   2655
            Begin VB.OptionButton Option4 
               Caption         =   "Facturadas y No Facturadas"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   20
               Top             =   1320
               Value           =   -1  'True
               Width           =   2415
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Facturas"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   19
               Top             =   1080
               Width           =   1335
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Notas de Venta"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   18
               Top             =   840
               Width           =   1455
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Garantias"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   17
               Top             =   600
               Width           =   2295
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Cotizaciones"
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   16
               Top             =   360
               Width           =   2295
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Busqueda particular"
            Height          =   1695
            Left            =   7800
            TabIndex        =   10
            Top             =   240
            Width           =   2055
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   3
               Left            =   120
               TabIndex        =   12
               Top             =   1200
               Width           =   1815
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "Reportes1.frx":5D78
               Left            =   120
               List            =   "Reportes1.frx":5D85
               Sorted          =   -1  'True
               TabIndex        =   11
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label Label4 
               Caption         =   "# Nota de Venta"
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
               Left            =   120
               TabIndex        =   14
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label5 
               Caption         =   "Sucursal"
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
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Con Comentario "
            Height          =   855
            Left            =   5040
            TabIndex        =   7
            Top             =   1920
            Visible         =   0   'False
            Width           =   4815
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   4
               Left            =   120
               TabIndex        =   8
               Top             =   480
               Width           =   4575
            End
            Begin VB.Label Label7 
               Caption         =   "Comentario"
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
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Label Label8 
            Caption         =   "Cliente"
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
            Left            =   240
            TabIndex        =   24
            Top             =   120
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60620800
         CurrentDate     =   39063
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60620800
         CurrentDate     =   39063
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha Inicio"
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
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha Fin"
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
         Left            =   4080
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Reportes1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
'FUNCION PARA MANEJAR EL ANCHO DE LAS COLUMNAS DEL LISTVIEW1
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_SETCOLUMNWIDTH = &H101E
Private Sub Check1_Click()
    Frame27.Enabled = True
    If Check1.Value Then
        Frame27.Enabled = False
        Option9(5).Enabled = False
        Option9(4).Value = True
        Option9(2).Enabled = False
        Option9(1).Enabled = False
        Option9(0).Enabled = False
    End If
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Combo1_DropDown()
    Dim sBuscarSucursal As String
    Dim rstSucursales As ADODB.Recordset
    sBuscarSucursal = "SELECT NOMBRE FROM SUCURSALES ORDER BY NOMBRE"
    Set rstSucursales = cnn.Execute(sBuscarSucursal)
    With rstSucursales
        Combo1.Clear
        If (.EOF And .BOF) Then
            MsgBox ("NO EXISTEN SUCURSALES")
        Else
            Do While Not .EOF
                Combo1.AddItem (.Fields("NOMBRE"))
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HFFE1E1
End Sub
Private Sub Combo1_LostFocus()
    Combo1.BackColor = &H80000005
End Sub
Private Sub DTPicker1_Change()
    DTPicker2.MinDate = DTPicker1.Value
End Sub
Private Sub Form_Load()
On Error GoTo ManejaError
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    DTPicker2.MinDate = Format(Date, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE PRODUCTO", 1800
        .ColumnHeaders.Add , , "Descripcion"
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE PRODUCTO", 1800
        .ColumnHeaders.Add , , "Descripcion", 4500
    End With
    With ListView6
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE PRODUCTO", 1800
        .ColumnHeaders.Add , , "Descripcion", 4500
    End With
    With ListaDocumentos
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE CLIENTE"
        .ColumnHeaders.Add , , "NOMBRE"
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE AGENTE", 1800
        .ColumnHeaders.Add , , "NOMBRE", 4500
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE PROVEEDOR", 1800
        .ColumnHeaders.Add , , "NOMBRE", 4500
    End With
    With ListView5
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_SUCURSAL", 1800
        .ColumnHeaders.Add , , "SUCURSAL", 3800
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image2_Click()
    FrmRepVentas.Show vbModal
End Sub
Private Sub ListaDocumentos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListaDocumentos.ListItems.Count > 0 Then
        Text1(1).Text = Item.SubItems(1)
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView1.ListItems.Count > 0 Then
        Text1(0).Text = Trim(Item)
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView2.ListItems.Count > 0 Then
        Text1(2).Text = Item.SubItems(1)
        Text2.Text = Item
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView3.ListItems.Count > 0 Then
        Text1(5).Text = Trim(Item)
        Text3.Text = Trim(Item)
    End If
End Sub
Private Sub ListView4_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView4.ListItems.Count > 0 Then
        Text1(6).Text = Trim(Item.SubItems(1))
    End If
End Sub
Private Sub ListView5_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView5.ListItems.Count > 0 Then
        Text1(7).Text = Trim(Item.SubItems(1))
    End If
End Sub
Private Sub ListView6_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView6.ListItems.Count > 0 Then
        Text1(8).Text = Trim(Item.SubItems(1))
        Text4.Text = Trim(Item.SubItems(1))
    End If
End Sub
Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Text1(8).SetFocus
    ElseIf Index = 1 Then
        Text1(0).SetFocus
    End If
End Sub
Private Sub Option10_Click(Index As Integer)
    Frame24.Visible = False
    If Index = 0 Then
        Frame24.Visible = True
    End If
End Sub
Private Sub Option12_Click(Index As Integer)
    If Index = 1 Then
        Frame23.Visible = True
        Frame25.Visible = True
        Frame22.Visible = False
    Else
        Frame23.Visible = False
        Frame25.Visible = False
        Frame22.Visible = True
    End If
End Sub
Private Sub Option3_Click()
    Text1(5).SetFocus
End Sub
Private Sub Option4_Click(Index As Integer)
    Frame12.Visible = False
    Frame9.Visible = False
    If Index < 3 Then
        Frame11.Visible = True
        Frame9.Visible = True
        If Index = 2 Then
            Label4.Caption = "# Nota de Venta"
        ElseIf Index = 1 Then
            Label4.Caption = "# Factura"
            Frame12.Visible = True
        Else
            Label4.Caption = "# Nota de Venta"
        End If
    Else
        Frame11.Visible = False
    End If
    If Index = 4 Then
        Option9(3).Enabled = True
        Option9(0).Enabled = False
    Else
        Option9(3).Enabled = True
        Option9(0).Enabled = True
    End If
    Text1(3).Text = ""
    Text1(4).Text = ""
End Sub
Private Sub Option5_Click(Index As Integer)
    If Index = 2 Then
        Text1(0).SetFocus
    ElseIf Index = 0 Then
        Text1(5).SetFocus
    ElseIf Index = 1 Then
        Text1(8).SetFocus
    End If
End Sub
Private Sub rpt_Click()
    Dim sBuscar As String
    Dim sBuscar2 As String
    Dim sBuscar3 As String
    Dim sCond As String     'se utiliza para hacer la condicionante de busqueda por el rango de fechas
    Dim sCond2 As String
    Dim orden As String
    Dim FechaCadena As String
    If DTPicker1.Value = DTPicker2.Value Then
        sCond = "where (Fecha = '" & DTPicker1.Value & "') "
        FechaCadena = "DEL " & DTPicker1.Value
    Else
        sCond = "where (Fecha >= '" & DTPicker1.Value & "' and Fecha <= '" & DTPicker2.Value & "') "
        FechaCadena = "DEL " & DTPicker1.Value & " AL " & DTPicker2.Value
    End If
    sCond2 = ""
    If SSTab1.Tab = 0 Then
       If (Option4(0).Value = True) Or (Option4(1).Value = True) Or (Option4(2).Value = True) Then
            sBuscar = "Select FECHA,SUCURSAL,ID_USUARIO,ID_VENTA,FOLIO,FACTURADO,ID_CLIENTE,NOMBRE,ID_PRODUCTO,DESCRIPCION,CANTIDAD,PRECIO_VENTA,((CANTIDAD*PRECIO_VENTA)) AS SUBTOTAL,((CANTIDAD*PRECIO_VENTA)*0.16) AS IVA,((CANTIDAD*PRECIO_VENTA)+ ((CANTIDAD*PRECIO_VENTA)*0.16)) AS TOTAL,CONTADO,COMENTARIO,FORMA_PAGO,[NUMERO DE COMANDA  O A.T.],IDDET from vsVentaRep "
            If Text1(4).Text <> "" Then sCond2 = "Comentario = '" & Text1(4).Text & "' " 'Text de comentario
            If Combo1.Text <> "" Then
                If sCond2 = "" Then
                    If Text1(3).Text <> "" Then ' text nota de venta
                        If Option4(1).Value = True Then
                            sCond2 = "Folio = '" & Combo1.ItemData(Combo1.ListIndex) & Text1(3).Text & "' "
                        Else
                            sCond2 = "id_venta = " & Text1(3).Text
                        End If
                        sCond2 = sCond2 & " and sucursal = '" & Combo1.Text & "' "
                    Else
                        sCond2 = "sucursal = '" & Combo1.Text & "' "
                    End If
                Else
                    If Text1(3).Text <> "" Then sCond2 = sCond2 & " and Folio = '" & Combo1.ItemData(Combo1.ListIndex) & Text1(3).Text & "' "
                    sCond2 = sCond2 & " and sucursal = '" & Combo1.Text & "' "
                End If
            Else
                If sCond2 = "" Then
                    If Text1(3).Text <> "" Then
                        If Option4(1).Value = True Then
                            sCond2 = "Folio Like '%" & Text1(3).Text & "' "
                        Else
                            sCond2 = "id_venta = " & Text1(3).Text
                        End If
                    End If
                Else
                    If Text1(3).Text <> "" Then
                        If Option4(1).Value = True Then
                            sCond2 = sCond2 & " and Folio Like '%" & Text1(3).Text & "' "
                        Else
                            sCond2 = sCond2 & " and id_venta = " & Text1(3).Text
                        End If
                    End If
                End If
            End If
        ElseIf (Option4(3).Value = True) Then
            sBuscar = "Select * from VsGarantiaRep"  'No utilizan sucursal y la consulta utiliza sucursal
        ElseIf (Option4(4).Value = True) Then         '
            sBuscar = "Select * from VsCotizaRep"     'No utiliza sucursal y la consulta utiliza sucursal
        End If
        If Text1(0).Text <> "" Then sCond = sCond & "and id_producto = '" & Text1(0).Text & "' "
        If Text1(1).Text <> "" Then sCond = sCond & "and nombre = '" & Text1(1).Text & "' "
        If Text1(2).Text <> "" Then sCond = sCond & "and id_usuario = '" & Text2.Text & "' "
        If Not (sCond = "" And sCond2 = "") Then
            If sCond = "" Then
                sCond2 = " where " & sCond2
                sBuscar = sBuscar & " " & sCond2
            Else
                If sCond2 <> "" Then sCond2 = " and " & sCond2
                sBuscar = sBuscar & " " & sCond & sCond2
            End If
        End If
        If Option9(0).Value Then
            orden = "sucursal"
        ElseIf Option9(1).Value Then
            orden = "Folio"
        ElseIf Option9(2).Value Then
            orden = "Fecha"
        ElseIf Option9(3).Value Then
            orden = "ID_USUARIO"
        ElseIf Option9(4).Value Then
            orden = "id_producto"
        Else
            orden = "Nombre"
        End If
        If sBuscar <> "" Then
            sBuscar = sBuscar & " order by " & orden
            frmVerRPT.txtSQL(0).Text = sBuscar
            frmVerRPT.Label1.Caption = "REPORTE DE VENTAS " & FechaCadena
            frmVerRPT.Show vbModal
            Text1(0).Text = ""
            Text1(1).Text = ""
            Text1(2).Text = ""
            Text2.Text = ""
            ListView1.ListItems.Clear
            ListView2.ListItems.Clear
            ListaDocumentos.ListItems.Clear
         Else
            MsgBox ("NECESARIO SELECCIONAR AL MENOS UN ALMACEN"), vbExclamation, "SACC"
         End If
    ElseIf SSTab1.Tab = 1 Then
        If Option10(0).Value Then
            sBuscar = "select * from  VSInvAlm3 "
            sBuscar2 = "select * from  VSInvAlm2 "
            sBuscar3 = "select * from  VSInvAlm12 "
        ElseIf Option10(1).Value Then
            sBuscar = "select * from vsPedido_PD "
        ElseIf Option10(2).Value Then
            sBuscar = "select * from vsEntrada_ED "
        ElseIf Option10(3).Value Then
            sBuscar = "select * from REQUISICION "
        ElseIf Option10(4).Value Then
            sBuscar = "select * from SCRAP "
        Else
            sBuscar = "select * from vsProducciones "
        End If
        If DTPicker1.Value = DTPicker2.Value Then
            sCond = "where (Fecha = '" & DTPicker1.Value & "') "
            FechaCadena = "DEL " & DTPicker1.Value
        Else
            sCond = "where (Fecha >= '" & DTPicker1.Value & "' and Fecha <= '" & DTPicker2.Value & "') "
            FechaCadena = "DEL " & DTPicker1.Value & " AL " & DTPicker2.Value
        End If
        If Frame24.Visible Then
            If Option6(0).Value Then
                sCond = ""
            ElseIf Option6(1).Value Then
                sCond = "WHERE CANTIDAD > 0 "
            ElseIf Option6(2).Value Then
                sCond = "WHERE CANTIDAD > C_MAXIMA "
            ElseIf Option6(3).Value Then
                sCond = "WHERE (CANTIDAD < C_MAXIMA AND CANTIDAD > C_MINIMA) "
            ElseIf Option6(4).Value Then
                sCond = "WHERE CANTIDAD < C_MINIMA "
            End If
            If Text1(7).Text <> "" Then
                If sCond = "" Then
                    sCond = "WHERE SUCURSAL = '" & Text1(7).Text & "' "
                Else
                    sCond = sCond & "AND SUCURSAL = '" & Text1(7).Text & "' "
                End If
            End If
            If Text3.Text <> "" Then
                If sCond = "" Then
                    sCond = " where id_producto = '" & Text3.Text & "' "
                Else
                    sCond = sCond & "and id_producto = '" & Text3.Text & "' "
                End If
            End If
        Else
            If Text1(7).Text <> "" Then sCond = sCond & "AND SUCURSAL = '" & Text1(7).Text & "' "
            If Text1(5).Text <> "" Then sCond = sCond & "and id_producto = '" & Text1(5).Text & "' "
        End If
        If Option10(0).Value Then
            If Check2.Value = 0 Then
                sBuscar3 = ""
            Else
                sBuscar3 = sBuscar3 & sCond 'estamos sustituyendo la variable  sBuscar2 por sBuscarAlmacen
            End If
            
            If Check6.Value = 0 Then
                sBuscar2 = ""
            Else
                sBuscar2 = sBuscar2 & sCond
            End If
            
            If Check7.Value = 0 Then
                sBuscar = ""
            Else
                sBuscar = sBuscar & sCond
            End If
        Else
            sBuscar = sBuscar & sCond
            sBuscar2 = ""
            sBuscar3 = ""
        End If
        If sBuscar <> "" Or sBuscar2 <> "" Or sBuscar3 <> "" Then
            frmVerRPT.txtSQL(0).Text = sBuscar
            frmVerRPT.txtSQL(1).Text = sBuscar2
            frmVerRPT.txtSQL(2).Text = sBuscar3
            frmVerRPT.Label1.Caption = "REPORTE DE ALMACEN " & FechaCadena
            frmVerRPT.Show vbModal
            Text1(5).Text = ""
            Text1(7).Text = ""
            ListView3.ListItems.Clear
            ListView5.ListItems.Clear
        Else
            MsgBox "NESESITA ELEGIR AL MENOS UN ALMACEN PARA BUSCAR", vbCritical, "SACC"
        End If
    Else
        'Terminar este proceso de reportes el dia 14/12/2006 a mas tardar
        If DTPicker1.Value = DTPicker2.Value Then
            sCond = "where (Fecha = '" & DTPicker1.Value & "') "
            FechaCadena = "DEL " & DTPicker1.Value
        Else
            sCond = "where (Fecha >= '" & DTPicker1.Value & "' and Fecha <= '" & DTPicker2.Value & "') "
            FechaCadena = "DEL " & DTPicker1.Value & " AL " & DTPicker2.Value
        End If
        sCond2 = ""
        If Option12(0).Value Then
            sBuscar = "SELECT V.NUM_ORDEN AS FOLIO, P.NOMBRE, V.FECHA, V.CONFIRMADA AS EDO, V.ID_PRODUCTO AS CLAVE, V.Descripcion, V.CANTIDAD, V.PRECIO, V.TOTAL  FROM VSORDENES AS V JOIN PROVEEDOR AS P ON P.ID_PROVEEDOR = V.ID_PROVEEDOR "
            If Check8.Value = 1 Then sCond2 = "CONFIRMADA = 'N'"
            If Check9.Value = 1 Then
                If sCond2 <> "" Then
                    sCond2 = sCond2 & " OR "
                End If
                sCond2 = sCond2 & "CONFIRMADA = 'S'"
            End If
            If Check10.Value = 1 Then
                If sCond2 <> "" Then
                    sCond2 = sCond2 & " OR "
                End If
                sCond2 = sCond2 & "CONFIRMADA = 'X'"
            End If
            If Check11.Value = 1 Then
                If sCond2 <> "" Then
                    sCond2 = sCond2 & " OR "
                End If
                sCond2 = sCond2 & "CONFIRMADA = 'Y'"
            End If
        Else
            sBuscar = "SELECT V.NUM_ORDEN AS FOLIO, P.NOMBRE, V.FECHA, V.CONFIRMADA AS EDO, V.ID_PRODUCTO AS CLAVE, V.Descripcion, V.CANTIDAD, V.CANTIDADP AS PENDIENTE, V.PRECIO, V.TOTAL  FROM VSORDENES AS V JOIN PROVEEDOR AS P ON P.ID_PROVEEDOR = V.ID_PROVEEDOR "
            If Check13.Value = 1 Then sCond2 = "CONFIRMADA = 'X'"
            If Check12.Value = 1 Then
                If sCond2 <> "" Then
                    sCond2 = sCond2 & " OR "
                End If
                sCond2 = sCond2 & "CONFIRMADA = 'Y'"
            End If
            If Option13(1).Value Then
                sCond = sCond & "AND V.TOTAL_PEND = V.TOTAL_PEDIDO "
            ElseIf Option13(2).Value Then
                sCond = sCond & "AND V.TOTAL_PEND < V.TOTAL_PEDIDO "
            ElseIf Option13(3).Value Then
                sCond = sCond & "AND V.TOTAL_PEND = 0 "
            End If
        End If
        If Text1(6).Text <> "" Then sCond = sCond & "AND P.NOMBRE = '" & Text1(6).Text & "' "
        If Text4.Text <> "" Then sCond = sCond & "and V.id_producto = '" & Text4.Text & "' "
        If sCond2 <> "" Then
            Text4.Text = ""
            sCond2 = "AND (" & sCond2 & ")"
            frmVerRPT.txtSQL(0).Text = sBuscar & sCond & sCond2
            frmVerRPT.Label1.Caption = "REPORTE DE COMPRAS " & FechaCadena
            frmVerRPT.Show vbModal
            Text1(6).Text = ""
            Text1(8).Text = ""
            ListView6.ListItems.Clear
            ListView4.ListItems.Clear
        Else
            MsgBox "NESESITA ELEGIR AL MENOS UN ESTADO DE LA ORDEN DE COMPRA", vbCritical, "SACC"
        End If
    End If
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).BackColor = &HFFE1E1
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = &H80000005
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And Text1(Index).Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        Dim sBus As String
        Dim CadClien As String
        Dim i As Integer
        If Index = 0 Then
            ListView1.ListItems.Clear
            If Check5.Value Then
                If Option1(1).Value Then
                    sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text1(0).Text & "%'"
                Else
                    sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE Descripcion LIKE '%" & Text1(0).Text & "%'"
                End If
                    Set tRs = cnn.Execute(sBus)
                With tRs
                    Do While Not .EOF
                        Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                        .MoveNext
                    Loop
                End With
            End If
            If Check4.Value Then
                sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text1(0).Text & "%'"
                Set tRs = cnn.Execute(sBus)
                With tRs
                    Do While Not .EOF
                        Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                        .MoveNext
                    Loop
                End With
            End If
            If Check3.Value Then
                sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & Text1(0).Text & "%'"
                Set tRs = cnn.Execute(sBus)
                With tRs
                    Do While Not .EOF
                        Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                        .MoveNext
                    Loop
                End With
            End If
            ' CICLO PARA CAMBIAR DEL ANCHO DE COLUMNAS DEL LISTVIEW1
            For i = 0 To ListView1.ColumnHeaders.Count - 1
                SendMessage ListView1.hWnd, LVM_SETCOLUMNWIDTH, i, -3
            Next i
            If sBus <> "" Then
                ListView1.SetFocus
            Else
                MsgBox "NESESITA ELEGIR AL MENOS UN ALMACEN PARA BUSCAR", vbCritical, "SACC"
            End If
        ElseIf Index = 1 Then
            sBus = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1(Index).Text & "%'"
            Set tRs = cnn.Execute(sBus)
            With tRs
                ListaDocumentos.ListItems.Clear
                Do While Not .EOF
                    Set tLi = ListaDocumentos.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & ""
                    .MoveNext
                Loop
                ListaDocumentos.SetFocus
                ' CICLO PARA CAMBIAR DEL ANCHO DE COLUMNAS DEL ListDocumentos
                For i = 0 To ListaDocumentos.ColumnHeaders.Count - 1
                    SendMessage ListaDocumentos.hWnd, LVM_SETCOLUMNWIDTH, i, -3
                Next i
            End With
        ElseIf Index = 2 Then
            sBus = "SELECT ID_USUARIO, NOMBRE, APELLIDOS FROM USUARIOS WHERE NOMBRE LIKE '%" & Text1(Index).Text & "%'"
            Set tRs = cnn.Execute(sBus)
            With tRs
                ListView2.ListItems.Clear
                Do While Not .EOF
                    Set tLi = ListView2.ListItems.Add(, , .Fields("ID_USUARIO") & "")
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & " " & .Fields("APELLIDOS")
                    .MoveNext
                Loop
                ListView2.SetFocus
                ' CICLO PARA CAMBIAR DEL ANCHO DE COLUMNAS DEL ListView2
                    For i = 0 To ListView2.ColumnHeaders.Count - 1
                        SendMessage ListView2.hWnd, LVM_SETCOLUMNWIDTH, i, -3
                    Next i
            End With
        ElseIf Index = 5 Then
            ListView3.ListItems.Clear
            If Check7.Value Then
                If Option3.Value = True Then
                    sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text1(5).Text & "%'"
                Else
                    sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE Descripcion LIKE '%" & Text1(5).Text & "%'"
                End If
                Set tRs = cnn.Execute(sBus)
                With tRs
                    Do While Not .EOF
                        Set tLi = ListView3.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                        .MoveNext
                     Loop
                End With
            End If
            If Check6.Value Then
                If Option3.Value = True Then
                    sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text1(5).Text & "%'"
                Else
                    sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN2 WHERE Descripcion LIKE '%" & Text1(5).Text & "%'"
                End If
                Set tRs = cnn.Execute(sBus)
                With tRs
                    Do While Not .EOF
                        Set tLi = ListView3.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                        .MoveNext
                    Loop
                End With
            End If
            If Check2.Value Then
                If Option3.Value = True Then
                    sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & Text1(5).Text & "%'"
                Else
                    sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN1 WHERE Descripcion LIKE '%" & Text1(5).Text & "%'"
                End If
                Set tRs = cnn.Execute(sBus)
                With tRs
                    Do While Not .EOF
                        Set tLi = ListView3.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                        .MoveNext
                    Loop
                End With
            End If
            ' CICLO PARA CAMBIAR DEL ANCHO DE COLUMNAS DEL LISTVIEW3
             For i = 0 To ListView3.ColumnHeaders.Count - 1
                 SendMessage ListView3.hWnd, LVM_SETCOLUMNWIDTH, i, -3
             Next i
            ListView3.SetFocus
        ElseIf Index = 6 Then
            sBus = "SELECT ID_PROVEEDOR, NOMBRE FROM PROVEEDOR WHERE NOMBRE LIKE '%" & Text1(Index).Text & "%'"
            Set tRs = cnn.Execute(sBus)
            With tRs
                ListView4.ListItems.Clear
                Do While Not .EOF
                    Set tLi = ListView4.ListItems.Add(, , .Fields("ID_PROVEEDOR") & "")
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & ""
                    .MoveNext
                Loop
                ' CICLO PARA CAMBIAR DEL ANCHO DE COLUMNAS DEL LISTVIEW5
                    For i = 0 To ListView4.ColumnHeaders.Count - 1
                        SendMessage ListView4.hWnd, LVM_SETCOLUMNWIDTH, i, -3
                    Next i
                ListView4.SetFocus
            End With
        ElseIf Index = 7 Then
            sBus = "SELECT ID_SUCURSAL, NOMBRE FROM SUCURSALES WHERE NOMBRE LIKE '%" & Text1(Index).Text & "%'"
            Set tRs = cnn.Execute(sBus)
            With tRs
                ListView5.ListItems.Clear
                Do While Not .EOF
                    Set tLi = ListView5.ListItems.Add(, , .Fields("ID_SUCURSAL") & "")
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & ""
                    .MoveNext
                Loop
                ListView4.SetFocus
            End With
        ElseIf Index = 8 Then
            ListView6.ListItems.Clear
            'VsFacProdAlm1
            If Check14.Value Then
                If Option3.Value = True Then
                    sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text1(8).Text & "%'"
                Else
                    If Option2.Value = True Then
                        sBus = "SELECT ID_PRODUCTO, Descripcion FROM VsFacProdAlm3 WHERE NO_ENVIO = '" & Text1(8).Text & "'"
                    Else
                        If Option7.Value = True Then
                            sBus = "SELECT ID_PRODUCTO, Descripcion FROM VsFacProdAlm3 WHERE FACT_PROVE = '" & Text1(8).Text & "'"
                        Else
                            sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE Descripcion LIKE '%" & Text1(8).Text & "%'"
                        End If
                    End If
                End If
                Set tRs = cnn.Execute(sBus)
                With tRs
                    Do While Not .EOF
                        Set tLi = ListView6.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                        .MoveNext
                    Loop
                End With
            End If
            If Check15.Value Then
                If Option3.Value = True Then
                    sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text1(8).Text & "%'"
                Else
                    If Option2.Value = True Then
                        sBus = "SELECT ID_PRODUCTO, Descripcion FROM VsFacProdAlm2 WHERE NO_ENVIO = '" & Text1(8).Text & "'"
                    Else
                        If Option7.Value = True Then
                            sBus = "SELECT ID_PRODUCTO, Descripcion FROM VsFacProdAlm2 WHERE FACT_PROVE = '" & Text1(8).Text & "'"
                        Else
                            sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN2 WHERE Descripcion LIKE '%" & Text1(8).Text & "%'"
                        End If
                    End If
                End If
                Set tRs = cnn.Execute(sBus)
                With tRs
                    Do While Not .EOF
                        Set tLi = ListView6.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                        .MoveNext
                    Loop
                End With
            End If
            If Check16.Value Then
                If Option3.Value = True Then
                    sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & Text1(8).Text & "%'"
                Else
                    If Option2.Value = True Then
                        sBus = "SELECT ID_PRODUCTO, Descripcion FROM VsFacProdAlm1 WHERE NO_ENVIO = '" & Text1(8).Text & "'"
                    Else
                        If Option7.Value = True Then
                            sBus = "SELECT ID_PRODUCTO, Descripcion FROM VsFacProdAlm1 WHERE FACT_PROVE = '" & Text1(8).Text & "'"
                        Else
                            sBus = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN1 WHERE Descripcion LIKE '%" & Text1(8).Text & "%'"
                        End If
                    End If
                End If
                Set tRs = cnn.Execute(sBus)
                With tRs
                    Do While Not .EOF
                        Set tLi = ListView6.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                        .MoveNext
                    Loop
                End With
            End If
            ' CICLO PARA CAMBIAR DEL ANCHO DE COLUMNAS DEL LISTVIEW5
            For i = 0 To ListView6.ColumnHeaders.Count - 1
                SendMessage ListView6.hWnd, LVM_SETCOLUMNWIDTH, i, -3
            Next i
            ListView6.SetFocus
        End If
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890.% "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
