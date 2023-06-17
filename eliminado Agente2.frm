VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form EliAgente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminar o Modificar Usuario"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "eliminado Agente2.frx":0000
      Left            =   960
      List            =   "eliminado Agente2.frx":001F
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   10
      Top             =   3360
      Width           =   975
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   15
         Top             =   0
         Width           =   975
         Begin VB.Image Image17 
            Height          =   705
            Left            =   120
            MouseIcon       =   "eliminado Agente2.frx":006E
            MousePointer    =   99  'Custom
            Picture         =   "eliminado Agente2.frx":0378
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
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
            TabIndex        =   16
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   13
         Top             =   1320
         Width           =   975
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "eliminado Agente2.frx":1E2A
            MousePointer    =   99  'Custom
            Picture         =   "eliminado Agente2.frx":2134
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Aceptar"
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
            TabIndex        =   14
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   11
         Top             =   1320
         Width           =   975
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Guardar"
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
            TabIndex        =   12
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image15 
            Height          =   720
            Left            =   120
            MouseIcon       =   "eliminado Agente2.frx":395E
            MousePointer    =   99  'Custom
            Picture         =   "eliminado Agente2.frx":3C68
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Image Image18 
         Height          =   750
         Left            =   120
         MouseIcon       =   "eliminado Agente2.frx":562A
         MousePointer    =   99  'Custom
         Picture         =   "eliminado Agente2.frx":5934
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Eliminar"
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
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   8
      Top             =   2160
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "eliminado Agente2.frx":765E
         MousePointer    =   99  'Custom
         Picture         =   "eliminado Agente2.frx":7968
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Guardar"
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
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   6
      Top             =   4560
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "eliminado Agente2.frx":932A
         MousePointer    =   99  'Custom
         Picture         =   "eliminado Agente2.frx":9634
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label26 
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
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
      Left            =   3240
      Picture         =   "eliminado Agente2.frx":B716
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7858
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "ID_AGENTE"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   4560
      TabIndex        =   20
      Top             =   120
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Almacén"
      TabPicture(0)   =   "eliminado Agente2.frx":E0E8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Check68"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Ventas"
      TabPicture(1)   =   "eliminado Agente2.frx":E104
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Admini."
      TabPicture(2)   =   "eliminado Agente2.frx":E120
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Varios"
      TabPicture(3)   =   "eliminado Agente2.frx":E13C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(2)=   "Frame4"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Generales"
      TabPicture(4)   =   "eliminado Agente2.frx":E158
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label4(0)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label7"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label8"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label9"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label10"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label11"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label13"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label4(1)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Combo1"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Te1"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Te2"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "Te3"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "Te4"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "Te6"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "Te5"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "Combo3"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).ControlCount=   16
      Begin VB.Frame Frame1 
         Caption         =   "Finanzas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   111
         Top             =   4320
         Width           =   4575
         Begin VB.CheckBox Check15 
            Caption         =   "Cancelar Pago a Proveedor"
            Height          =   195
            Left            =   120
            TabIndex        =   114
            Top             =   720
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox Check52 
            Caption         =   "Administrar Cuentas por Cobrar"
            Height          =   195
            Left            =   120
            TabIndex        =   113
            Top             =   480
            Width           =   2775
         End
         Begin VB.CheckBox Check66 
            Caption         =   "Pagar a Proveedor "
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "De Producción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   104
         Top             =   3480
         Width           =   4575
         Begin VB.CheckBox Check64 
            Caption         =   "Revisar Compra de Almacén 1"
            Height          =   195
            Left            =   120
            TabIndex        =   110
            Top             =   1440
            Width           =   2655
         End
         Begin VB.CheckBox Check61 
            Caption         =   "Surtir Material Extra en Producción"
            Height          =   195
            Left            =   120
            TabIndex        =   109
            Top             =   1200
            Width           =   3015
         End
         Begin VB.CheckBox Check37 
            Caption         =   "Ver Juego de Reparación"
            Height          =   195
            Left            =   120
            TabIndex        =   108
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox Check36 
            Caption         =   "Controlar Calidad de Procucción"
            Height          =   195
            Left            =   120
            TabIndex        =   107
            Top             =   720
            Width           =   2895
         End
         Begin VB.CheckBox Check35 
            Caption         =   "Ver Producción"
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox Check34 
            Caption         =   "Revición de Producción"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "De Soporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74880
         TabIndex        =   102
         Top             =   3600
         Width           =   4575
         Begin VB.CheckBox Check23 
            Caption         =   "Ver Asistencias Técnicas"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   -73440
         TabIndex        =   100
         Top             =   4320
         Width           =   3135
      End
      Begin VB.TextBox Te5 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -73440
         PasswordChar    =   "*"
         TabIndex        =   98
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox Te6 
         Height          =   285
         Left            =   -73440
         TabIndex        =   96
         Top             =   3360
         Width           =   3135
      End
      Begin VB.TextBox Te4 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -73440
         PasswordChar    =   "*"
         TabIndex        =   95
         Top             =   2400
         Width           =   3135
      End
      Begin VB.TextBox Te3 
         Height          =   285
         Left            =   -73440
         TabIndex        =   94
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox Te2 
         Height          =   285
         Left            =   -73440
         TabIndex        =   93
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox Te1 
         Height          =   285
         Left            =   -73440
         TabIndex        =   92
         Top             =   960
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -73440
         TabIndex        =   85
         Top             =   3840
         Width           =   3135
      End
      Begin VB.Frame Frame3 
         Caption         =   "De Almacen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         TabIndex        =   73
         Top             =   480
         Width           =   4575
         Begin VB.CheckBox Check50 
            Caption         =   "Revisar Faltantes"
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   2640
            Width           =   2295
         End
         Begin VB.CheckBox Check49 
            Caption         =   "Pedir Maximos y Minimos Almacén3"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   2400
            Width           =   3015
         End
         Begin VB.CheckBox Check48 
            Caption         =   "Surtir Ventas Programadas"
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   1680
            Width           =   2295
         End
         Begin VB.CheckBox Check46 
            Caption         =   "Rastrear Orden de Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   1920
            Width           =   2775
         End
         Begin VB.CheckBox Check21 
            Caption         =   "Capturar Inventarios"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CheckBox Check14 
            Caption         =   "Entrada Orden de Compra/Producción/Rápida"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Traspasar Inventarios / Surtir"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   480
            Width           =   2655
         End
         Begin VB.CheckBox Check19 
            Caption         =   "Revisar Entradas de productos"
            Height          =   195
            Left            =   120
            TabIndex        =   77
            Top             =   720
            Width           =   2655
         End
         Begin VB.CheckBox Check20 
            Caption         =   "Revisar Ventas Canceladas"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   960
            Width           =   3015
         End
         Begin VB.CheckBox Check41 
            Caption         =   "Dar Salidas para uso interno"
            Height          =   195
            Left            =   120
            TabIndex        =   75
            Top             =   1200
            Width           =   2415
         End
         Begin VB.CheckBox Check67 
            Caption         =   "Producir o Reemplazar Productos"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   2160
            Width           =   2895
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "De Administrador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   -74280
         TabIndex        =   52
         Top             =   480
         Width           =   3375
         Begin VB.CheckBox Check42 
            Caption         =   "Administración del Sistema"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   4800
            Width           =   2655
         End
         Begin VB.CheckBox Check33 
            Caption         =   "Agregar Juego de Reparación"
            Height          =   195
            Left            =   120
            TabIndex        =   71
            Top             =   1680
            Width           =   2535
         End
         Begin VB.CheckBox Check16 
            Caption         =   "Agregar Producto Almacén 1 y 2"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   1200
            Width           =   2895
         End
         Begin VB.CheckBox Check17 
            Caption         =   "Agregar Producto Almacén 3"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   1440
            Width           =   2655
         End
         Begin VB.CheckBox Check38 
            Caption         =   "Crear Promoción"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   3600
            Width           =   1935
         End
         Begin VB.CheckBox Check39 
            Caption         =   "Eliminar Producto de Almacén 1/2"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   2640
            Width           =   2775
         End
         Begin VB.CheckBox Check40 
            Caption         =   "Capturar Tipo de Cambio"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   3840
            Width           =   2415
         End
         Begin VB.CheckBox Check31 
            Caption         =   "Aplicar Sanción a Venta"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   3360
            Width           =   2895
         End
         Begin VB.CheckBox Check32 
            Caption         =   "Eliminar Producto de Almacén 3"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   2880
            Width           =   2655
         End
         Begin VB.CheckBox Check24 
            Caption         =   "Agregar Usuario y/o Mensajero"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   2775
         End
         Begin VB.CheckBox Check25 
            Caption         =   "Agregar Clientes"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox Check26 
            Caption         =   "Agregar Sucursal"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox Check27 
            Caption         =   "Agregar Proveedor"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox Check28 
            Caption         =   "Eliminar Usuario y/o Mensajero"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1920
            Width           =   2535
         End
         Begin VB.CheckBox Check29 
            Caption         =   "Eliminar Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   2160
            Width           =   2175
         End
         Begin VB.CheckBox Check30 
            Caption         =   "Eliminar Proveedor"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CheckBox Check59 
            Caption         =   "Autorizar Garantia y Remanufactura"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   4080
            Width           =   2895
         End
         Begin VB.CheckBox Check65 
            Caption         =   "Aprobar Compra Proveedores Varios"
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   4320
            Width           =   3015
         End
         Begin VB.CheckBox Check53 
            Caption         =   "Eliminar Sucursal"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   3120
            Width           =   2895
         End
         Begin VB.CheckBox Check60 
            Caption         =   "Editar Juegos de Reparación"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   4560
            Width           =   2415
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "De Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   -74280
         TabIndex        =   34
         Top             =   480
         Width           =   3375
         Begin VB.CheckBox Check13 
            Caption         =   "Hacer Notas de Crédito"
            Height          =   195
            Left            =   120
            TabIndex        =   101
            Top             =   2880
            Width           =   2175
         End
         Begin VB.CheckBox Check51 
            Caption         =   "Captura de Domicilios"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   3360
            Width           =   2175
         End
         Begin VB.CheckBox Check47 
            Caption         =   "Ver Detalle de Ventas Programadas"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   3120
            Width           =   3015
         End
         Begin VB.CheckBox Check44 
            Caption         =   "Hacer Ventas Programadas"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   2640
            Width           =   2415
         End
         Begin VB.CheckBox Check43 
            Caption         =   "Hacer Ventas Especiales"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Hacer Comanda"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Hacer Notas de Venta"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   1935
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Hacer Notas de Crédito y Vales de Caja"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   3135
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Hacer Notas de garantia"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Capturar Asistencias tecnicas"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1200
            Width           =   2535
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Buscar Producto"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Buscar Existencias"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Hacer Pedidos a Bodega"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   1920
            Width           =   2295
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Hacer Requisiciones y Orden de Prod."
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   2160
            Width           =   3015
         End
         Begin VB.CheckBox Check62 
            Caption         =   "Cancelar por Refacturación"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   3600
            Width           =   2895
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Administrador en ventas"
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
            TabIndex        =   36
            Top             =   4080
            Width           =   2895
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Realizar Corte de Caja"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   3840
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "(Cancelaciones, Permisos, Licitaciones)"
            Height          =   255
            Left            =   360
            TabIndex        =   51
            Top             =   4320
            Width           =   2895
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "De Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   4575
         Begin VB.CheckBox Check58 
            Caption         =   "Solicitar Cotización "
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   2775
         End
         Begin VB.CheckBox Check57 
            Caption         =   "Imprimir Orden de Compra"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CheckBox Check55 
            Caption         =   "Autorizar Cotización"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox Check54 
            Caption         =   "Revisar Orden de Compra"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   2535
         End
         Begin VB.CheckBox Check22 
            Caption         =   "Hacer Pre-Orden de Compra"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   2655
         End
         Begin VB.CheckBox Check69 
            Caption         =   "Aprobar Orden de Compra"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CheckBox Check70 
            Caption         =   "Regresar Orden de Compra"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1680
            Width           =   2295
         End
         Begin VB.CheckBox Check63 
            Caption         =   "Crear Compras a Proveedores Varios"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   2400
            Width           =   3255
         End
         Begin VB.CheckBox Check45 
            Caption         =   "Modificar Orden Rápida"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   2160
            Width           =   2055
         End
         Begin VB.CheckBox Check56 
            Caption         =   "Crear Orden Rápida"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1920
            Width           =   2415
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Cancelar Orden de Compra/Rápida/Proveedores Varios"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   2640
            Width           =   4335
         End
      End
      Begin VB.CheckBox Check68 
         Caption         =   "No Asignado"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   5280
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Departamento :"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   99
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Rep. Contraseña :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   97
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Id Usuario :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   91
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Contraseña :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   90
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Puesto :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   89
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Apellidos :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   88
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   87
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Sucursal :"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   86
         Top             =   3840
         Width           =   855
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Empresa :"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clave :"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   5400
      Width           =   615
   End
End
Attribute VB_Name = "EliAgente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private cnn2 As ADODB.Connection
Dim IdSuc As String
Private Sub Combo1_DropDown()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Combo1.Clear
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    Do While Not (tRs.EOF)
        Combo1.AddItem tRs.Fields("NOMBRE")
        tRs.MoveNext
    Loop
End Sub
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HFFE1E1
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Combo1_LostFocus()
    Combo1.BackColor = &H80000005
End Sub
Private Sub Combo2_LostFocus()
    'cnn2.Close
    Set cnn2 = New ADODB.Connection
    With cnn2
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & Combo2.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
End Sub
Private Sub Command1_Click()
    Buscar
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Id Agente", 0
        .ColumnHeaders.Add , , "Nombre", 4300, lvwColumnCenter
    End With
    Combo2.Text = NvoMen.TxtBaseDatos.Text
    sBuscar = "SELECT DEPARTAMENTO FROM DEPARTAMENTOS WHERE ESTATUS = 'A' AND TIPO = 'T' ORDER BY DEPARTAMENTO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            If Not IsNull(tRs.Fields("DEPARTAMENTO")) Then Combo3.AddItem tRs.Fields("DEPARTAMENTO")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image18_Click()
    Dim sBuscar As String
    sBuscar = "UPDATE USUARIOS SET ESTADO = 'I', ID_SUCURSAL = 0 WHERE ID_USUARIO = " & Text1.Text
    cnn.Execute (sBuscar)
    MsgBox "USUARIO ELIMINADO!", vbInformation, "SACC"
    Text1.Text = ""
    Text5.Text = ""
    ListView1.ListItems.Clear
End Sub
Private Sub Image8_Click()
    Dim Cont As Integer
    Dim cuPer As Integer
    Dim tRs As ADODB.Recordset
    cuPer = 0
    Dim sBuscar As String
    If Combo1.Text <> "" Then
        Dim tRs4 As ADODB.Recordset
        sBuscar = "SELECT ID_SUCURSAL FROM SUCURSALES WHERE NOMBRE LIKE '" & Combo1.Text & "'"
        Set tRs4 = cnn.Execute(sBuscar)
        If Not (tRs4.EOF And tRs4.BOF) Then
            IdSuc = tRs4.Fields("ID_SUCURSAL")
        End If
    End If
    Dim V(1 To 70) As String
    If Check1.Value = 1 Then
        V(1) = "S"
        cuPer = cuPer + 1
    Else
        V(1) = "N"
    End If
    If Check2.Value = 1 Then
        V(2) = "S"
        cuPer = cuPer + 1
    Else
        V(2) = "N"
    End If
    If Check3.Value = 1 Then
        V(3) = "S"
        cuPer = cuPer + 1
    Else
        V(3) = "N"
    End If
    If Check4.Value = 1 Then
        V(4) = "S"
        cuPer = cuPer + 1
    Else
        V(4) = "N"
    End If
    If Check5.Value = 1 Then
        V(5) = "S"
        cuPer = cuPer + 1
    Else
        V(5) = "N"
    End If
    If Check6.Value = 1 Then
        V(6) = "S"
        cuPer = cuPer + 1
    Else
        V(6) = "N"
    End If
    If Check7.Value = 1 Then
        V(7) = "S"
        cuPer = cuPer + 1
    Else
        V(7) = "N"
    End If
    If Check8.Value = 1 Then
        V(8) = "S"
        cuPer = cuPer + 1
    Else
        V(8) = "N"
    End If
    If Check9.Value = 1 Then
        V(9) = "S"
        cuPer = cuPer + 1
    Else
        V(9) = "N"
    End If
    If Check10.Value = 1 Then
        V(10) = "S"
        cuPer = cuPer + 1
    Else
        V(10) = "N"
    End If
    If Check11.Value = 1 Then
        V(11) = "S"
        cuPer = cuPer + 1
    Else
        V(11) = "N"
    End If
    If Check12.Value = 1 Then
        V(12) = "S"
        cuPer = cuPer + 1
    Else
        V(12) = "N"
    End If
    If Check13.Value = 1 Then
        V(13) = "S"
        cuPer = cuPer + 1
    Else
        V(13) = "N"
    End If
    If Check14.Value = 1 Then
        V(14) = "S"
        cuPer = cuPer + 1
    Else
        V(14) = "N"
    End If
    If Check15.Value = 1 Then
        V(15) = "S"
        cuPer = cuPer + 1
    Else
        V(15) = "N"
    End If
    If Check16.Value = 1 Then
        V(16) = "S"
        cuPer = cuPer + 1
    Else
        V(16) = "N"
    End If
    If Check17.Value = 1 Then
        V(17) = "S"
        cuPer = cuPer + 1
    Else
        V(17) = "N"
    End If
    If Check18.Value = 1 Then
        V(18) = "S"
        cuPer = cuPer + 1
    Else
        V(18) = "N"
    End If
    If Check19.Value = 1 Then
        V(19) = "S"
        cuPer = cuPer + 1
    Else
        V(19) = "N"
    End If
    If Check20.Value = 1 Then
        V(20) = "S"
        cuPer = cuPer + 1
    Else
        V(20) = "N"
    End If
    If Check21.Value = 1 Then
        V(21) = "S"
        cuPer = cuPer + 1
    Else
        V(21) = "N"
    End If
    If Check22.Value = 1 Then
        V(22) = "S"
        cuPer = cuPer + 1
    Else
        V(22) = "N"
    End If
    If Check23.Value = 1 Then
        V(23) = "S"
        cuPer = cuPer + 1
    Else
        V(23) = "N"
    End If
    If Check24.Value = 1 Then
        V(24) = "S"
        cuPer = cuPer + 1
    Else
        V(24) = "N"
    End If
    If Check25.Value = 1 Then
        V(25) = "S"
        cuPer = cuPer + 1
    Else
        V(25) = "N"
    End If
    If Check26.Value = 1 Then
        V(26) = "S"
        cuPer = cuPer + 1
    Else
        V(26) = "N"
    End If
    If Check27.Value = 1 Then
        V(27) = "S"
        cuPer = cuPer + 1
    Else
        V(27) = "N"
    End If
    If Check28.Value = 1 Then
        V(28) = "S"
        cuPer = cuPer + 1
    Else
        V(28) = "N"
    End If
    If Check29.Value = 1 Then
        V(29) = "S"
        cuPer = cuPer + 1
    Else
        V(29) = "N"
    End If
    If Check30.Value = 1 Then
        V(30) = "S"
        cuPer = cuPer + 1
    Else
        V(30) = "N"
    End If
    If Check31.Value = 1 Then
        V(31) = "S"
        cuPer = cuPer + 1
    Else
        V(31) = "N"
    End If
    If Check32.Value = 1 Then
        V(32) = "S"
        cuPer = cuPer + 1
    Else
        V(32) = "N"
    End If
    If Check33.Value = 1 Then
        V(33) = "S"
        cuPer = cuPer + 1
    Else
        V(33) = "N"
    End If
    If Check34.Value = 1 Then
        V(34) = "S"
        cuPer = cuPer + 1
    Else
        V(34) = "N"
    End If
    If Check35.Value = 1 Then
        V(35) = "S"
        cuPer = cuPer + 1
    Else
        V(35) = "N"
    End If
    If Check36.Value = 1 Then
        V(36) = "S"
        cuPer = cuPer + 1
    Else
        V(36) = "N"
    End If
    If Check37.Value = 1 Then
        V(37) = "S"
        cuPer = cuPer + 1
    Else
        V(37) = "N"
    End If
    If Check38.Value = 1 Then
        V(38) = "S"
        cuPer = cuPer + 1
    Else
        V(38) = "N"
    End If
    If Check39.Value = 1 Then
        V(39) = "S"
        cuPer = cuPer + 1
    Else
        V(39) = "N"
    End If
    If Check40.Value = 1 Then
        V(40) = "S"
        cuPer = cuPer + 1
    Else
        V(40) = "N"
    End If
    If Check41.Value = 1 Then
        V(41) = "S"
        cuPer = cuPer + 1
    Else
        V(41) = "N"
    End If
    If Check42.Value = 1 Then
        V(42) = "S"
        cuPer = cuPer + 1
    Else
        V(42) = "N"
    End If
    If Check43.Value = 1 Then
        V(43) = "S"
        cuPer = cuPer + 1
    Else
        V(43) = "N"
    End If
    If Check44.Value = 1 Then
        V(44) = "S"
        cuPer = cuPer + 1
    Else
        V(44) = "N"
    End If
    If Check45.Value = 1 Then
        V(45) = "S"
        cuPer = cuPer + 1
    Else
        V(45) = "N"
    End If
    If Check46.Value = 1 Then
        V(46) = "S"
        cuPer = cuPer + 1
    Else
        V(46) = "N"
    End If
    If Check47.Value = 1 Then
        V(47) = "S"
        cuPer = cuPer + 1
    Else
        V(47) = "N"
    End If
    If Check48.Value = 1 Then
        V(48) = "S"
        cuPer = cuPer + 1
    Else
        V(48) = "N"
    End If
    If Check49.Value = 1 Then
        V(49) = "S"
        cuPer = cuPer + 1
    Else
        V(49) = "N"
    End If
    If Check50.Value = 1 Then
        V(50) = "S"
        cuPer = cuPer + 1
    Else
        V(50) = "N"
    End If
    If Check51.Value = 1 Then
        V(51) = "S"
        cuPer = cuPer + 1
    Else
        V(51) = "N"
    End If
    If Check52.Value = 1 Then
        V(52) = "S"
        cuPer = cuPer + 1
    Else
        V(52) = "N"
    End If
    If Check53.Value = 1 Then
        V(53) = "S"
        cuPer = cuPer + 1
    Else
        V(53) = "N"
    End If
    If Check54.Value = 1 Then
        V(54) = "S"
        cuPer = cuPer + 1
    Else
        V(54) = "N"
    End If
    If Check55.Value = 1 Then
        V(55) = "S"
        cuPer = cuPer + 1
    Else
        V(55) = "N"
    End If
    If Check56.Value = 1 Then
        V(56) = "S"
        cuPer = cuPer + 1
    Else
        V(56) = "N"
    End If
    If Check57.Value = 1 Then
        V(57) = "S"
        cuPer = cuPer + 1
    Else
        V(57) = "N"
    End If
    If Check58.Value = 1 Then
        V(58) = "S"
        cuPer = cuPer + 1
    Else
        V(58) = "N"
    End If
    If Check59.Value = 1 Then
        V(59) = "S"
        cuPer = cuPer + 1
    Else
        V(59) = "N"
    End If
    If Check60.Value = 1 Then
        V(60) = "S"
        cuPer = cuPer + 1
    Else
        V(60) = "N"
    End If
    If Check61.Value = 1 Then
        V(61) = "S"
        cuPer = cuPer + 1
    Else
        V(61) = "N"
    End If
    If Check62.Value = 1 Then
        V(62) = "S"
        cuPer = cuPer + 1
    Else
        V(62) = "N"
    End If
    If Check63.Value = 1 Then
        V(63) = "S"
        cuPer = cuPer + 1
    Else
        V(63) = "N"
    End If
    If Check64.Value = 1 Then
        V(64) = "S"
        cuPer = cuPer + 1
    Else
        V(64) = "N"
    End If
    If Check65.Value = 1 Then
        V(65) = "S"
        cuPer = cuPer + 1
    Else
        V(65) = "N"
    End If
    If Check66.Value = 1 Then
        V(66) = "S"
        cuPer = cuPer + 1
    Else
        V(66) = "N"
    End If
    If Check67.Value = 1 Then
        V(67) = "S"
        cuPer = cuPer + 1
    Else
        V(67) = "N"
    End If
    If Check68.Value = 1 Then
        V(68) = "S"
        cuPer = cuPer + 1
    Else
        V(68) = "N"
    End If
    If Check69.Value = 1 Then
        V(69) = "S"
        cuPer = cuPer + 1
    Else
        V(69) = "N"
    End If
    If Check70.Value = 1 Then
        V(70) = "S"
        cuPer = cuPer + 1
    Else
        V(70) = "N"
    End If
    If Te1.Text <> "" And Te2.Text <> "" And Te3.Text <> "" And Te4.Text <> "" And Te6.Text <> "" And Combo1.Text <> "" Then
        If cuPer <> 0 Then
            Dim tRs2 As ADODB.Recordset
            Dim num As String
            sBuscar = "SELECT ID_USUARIO FROM USUARIOS WHERE NOMBRE = '" & Te1.Text & "'"
            Set tRs2 = cnn.Execute(sBuscar)
            If Not (tRs2.EOF And tRs2.BOF) Then
                sBuscar = "UPDATE USUARIOS SET NOMBRE = '" & Te1.Text & "', APELLIDOS = '" & Te2.Text & "', PUESTO = '" & Te3.Text & "', PASSWORD = '" & Te4.Text & "', ID_USUARIO = '" & Te6.Text & "', ID_SUCURSAL = '" & IdSuc & "', Pe1 = '" & V(1) & "', Pe2 = '" & V(2) & "', Pe3 = '" & V(3) & "', Pe4 = '" & V(4) & "', Pe5 = '" & V(5) & "', Pe6 = '" & V(6) & "', Pe7 = '" & V(7) & "', Pe8 = '" & V(8) & "', Pe9 = '" & V(9) & "', Pe10 = '" & V(10) & "', Pe11 = '" & V(11) & "', Pe12 = '" & V(12) & "', Pe13 = '" & V(13) & "', Pe14 = '" & V(14) & "', Pe15 = '" & V(15) & "', Pe16 = '" & V(16) & "', Pe17 = '" & V(17) & "', Pe18 = '" & V(18) & "', Pe19 = '" & V(19) & "', Pe20 = '" & V(20) & "', Pe21 = '" & V(21) & "', Pe22 = '" & V(22) & "', Pe23 = '" & V(23) & "', Pe24 = '" & V(24) & "', Pe25 = '" & V(25) & "', Pe26 = '" & V(26) & "', Pe27 = '" & V(27) & "', Pe28 = '" & V(28) & "', Pe29 = '" & V(29) & "', Pe30 = '" & V(30) & "', Pe31 = '" & V(31) & "', Pe32 = '" & V(32) & "', Pe33 = '" _
                & V(33) & "', Pe34 = '" & V(34) & "', Pe35 = '" & V(35) & "', Pe36 = '" & V(36) & "', Pe37 = '" & V(37) & "', Pe38 = '" & V(38) & "', Pe40 = '" & V(40) & "', Pe41 = '" & V(41) & "', Pe42 = '" & V(42) & "', Pe43 = '" & V(43) & "', Pe44 = '" & V(44) & "', Pe45 = '" & V(45) & "', Pe46 = '" & V(46) & "', Pe47 = '" & V(47) & "', Pe48 = '" & V(48) & "', Pe49 = '" & V(49) & "', Pe50 = '" & V(50) & "', Pe51 = '" & V(51) & "', Pe52 = '" & V(52) & "', Pe53 = '" & V(53) & "', Pe54 = '" & V(54) & "', Pe55 = '" & V(55) & "', Pe57 = '" & V(57) & "', Pe58 = '" & V(58) & "', Pe59 = '" & V(59) & "', Pe60 = '" & V(60) & "', Pe61 = '" & V(61) & "', Pe62 = '" & V(62) & "', Pe63 = '" & V(63) & "', Pe64 = '" & V(64) & "', Pe65 = '" & V(65) & "', Pe66 = '" & V(66) & "', Pe67 = '" & V(67) & "', Pe68 = '" & V(68) & "', Pe69 = '" & V(69) & "', Pe70 = '" & V(70) & "' WHERE ID_USUARIO = " & Text1.Text
            Else
                sBuscar = "INSERT INTO USUARIOS (NOMBRE, APELLIDOS, PUESTO, PASSWORD, ID_USUARIO, ID_SUCURSAL, DEPARTAMENTO, Pe1, Pe2, Pe3, Pe4, Pe5, Pe6, Pe7, Pe8, Pe9, Pe10, Pe11, Pe12, Pe13, Pe14, Pe15, Pe16, Pe17, Pe18, Pe19, Pe20, Pe21, Pe22, Pe23, Pe24, Pe25, Pe26, Pe27, Pe28, Pe29, Pe30, Pe31, Pe32, Pe33, Pe34, Pe35, Pe36, Pe37, Pe38, Pe39, Pe40, Pe41, Pe42, Pe43, Pe44, Pe45, Pe46, Pe47, Pe48, Pe49, Pe50, Pe51, Pe52, Pe53, Pe54, Pe55, Pe56, Pe57, Pe58, Pe59, Pe60, Pe61, Pe62, Pe63, Pe64, Pe65, Pe66, Pe67, Pe68, Pe69, Pe70) VALUES ('" & Te1.Text & "', '" & Te2.Text & "', '" & Te3.Text & "', '" & Te4.Text & "', '" & Te6.Text & "', '" & IdSuc & "', '" & Combo3.Text & "', '" & V(1) & "', '" & V(2) & "', '" & V(3) & "', '" & V(4) & "', '" & V(5) & "', '" & V(6) & "', '" & V(7) & "', '" & V(8) & "', '" & V(9) & "', '" & V(10) & "', '" & V(11) & "', '" _
                & V(12) & "', '" & V(13) & "', '" & V(14) & "', '" & V(15) & "', '" & V(16) & "', '" & V(17) & "', '" & V(17) & "', '" & V(19) & "', '" & V(20) & "', '" & V(21) & "', '" & V(22) & "', '" & V(23) & "', '" & V(24) & "', '" & V(25) & "', '" & V(26) & "', '" & V(27) & "', '" & V(28) & "', '" & V(29) & "', '" & V(30) & "', '" & V(31) & "', '" & V(32) & "', '" & V(33) & "', '" & V(34) & "', '" & V(35) & "', '" & V(36) & "', '" & V(37) & "', '" & V(38) & "', '" & V(39) & "', '" & V(40) & "', '" & V(41) & "', '" & V(42) & "', '" & V(43) & "', '" & V(44) & "', '" & V(45) & "', '" & V(46) & "', '" & V(47) & "', '" & V(48) & "', '" & V(49) & "', '" & V(50) & "', '" & V(51) & "', '" & V(52) & "', '" & V(53) & "', '" & V(54) & "', '" & V(55) & "', '" & V(56) & "', '" & V(57) & "', '" & V(58) & "', '" & V(59) & "', '" & V(60) & "', '" & V(61) & "', '" & V(62) & "', '" & V(63) & "', '" & V(64) & "', '" & V(65) & "', '" & V(66) & "', '" & V(67) & "', '" & V(68) & "', '" & V(69) & "', '" & V(70) & "');"
            End If
            Set tRs2 = cnn.Execute(sBuscar)
            'Set tRs2 = cnn2.Execute(sBuscar)
            Te1.Text = ""
            Te2.Text = ""
            Te3.Text = ""
            Te4.Text = ""
            Te6.Text = ""
            Combo1.Text = ""
            IdSuc = ""
            Check1.Value = 0
            Check2.Value = 0
            Check3.Value = 0
            Check4.Value = 0
            Check5.Value = 0
            Check6.Value = 0
            Check7.Value = 0
            Check8.Value = 0
            Check9.Value = 0
            Check10.Value = 0
            Check11.Value = 0
            Check12.Value = 0
            Check13.Value = 0
            Check14.Value = 0
            Check15.Value = 0
            Check16.Value = 0
            Check17.Value = 0
            Check18.Value = 0
            Check19.Value = 0
            Check20.Value = 0
            Check21.Value = 0
            Check22.Value = 0
            Check23.Value = 0
            Check24.Value = 0
            Check25.Value = 0
            Check26.Value = 0
            Check27.Value = 0
            Check28.Value = 0
            Check29.Value = 0
            Check30.Value = 0
            Check31.Value = 0
            Check32.Value = 0
            Check33.Value = 0
            Check34.Value = 0
            Check35.Value = 0
            Check36.Value = 0
            Check37.Value = 0
            Check38.Value = 0
            Check39.Value = 0
            Check40.Value = 0
            Check41.Value = 0
            Check42.Value = 0
            Check43.Value = 0
            Check44.Value = 0
            Check45.Value = 0
            Check46.Value = 0
            Check47.Value = 0
            Check48.Value = 0
            Check49.Value = 0
            Check50.Value = 0
            Check51.Value = 0
            Check52.Value = 0
            Check53.Value = 0
            Check54.Value = 0
            Check55.Value = 0
            Check56.Value = 0
            Check57.Value = 0
            Check58.Value = 0
            Check59.Value = 0
            Check60.Value = 0
            Check61.Value = 0
            Check62.Value = 0
            Check63.Value = 0
            Check64.Value = 0
            Check65.Value = 0
            Check66.Value = 0
            Check67.Value = 0
            Check68.Value = 0
            Check69.Value = 0
            Check70.Value = 0
        Else
            MsgBox "No ha dado permisos a este usuario, es necesario dar al menos un permiso"
        End If
    Else
        MsgBox "Falta informacion necesaria para la alta del Usuario"
    End If
    cuPer = 0
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Combo1.Clear
    sBuscar = "SELECT * FROM USUARIOS WHERE ID_USUARIO = " & Item
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.EOF) Then
        Te1.Text = tRs.Fields("NOMBRE")
        Te2.Text = tRs.Fields("APELLIDOS")
        Te3.Text = tRs.Fields("PUESTO")
        Te4.Text = tRs.Fields("PASSWORD")
        Te6.Text = tRs.Fields("ID_USUARIO")
        Combo3.Text = tRs.Fields("DEPARTAMENTO")
        sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ID_SUCURSAL = " & tRs.Fields("ID_SUCURSAL") & " AND ELIMINADO = 'N'"
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.EOF) Then
            Combo1.Text = tRs2.Fields("NOMBRE")
        Else
            Combo1.Text = "Sucursale Eliminada"
        End If
        If tRs.Fields("Pe1") = "S" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
        If tRs.Fields("Pe2") = "S" Then
            Check2.Value = 1
        Else
            Check2.Value = 0
        End If
        If tRs.Fields("Pe3") = "S" Then
            Check3.Value = 1
        Else
            Check3.Value = 0
        End If
        If tRs.Fields("Pe4") = "S" Then
            Check4.Value = 1
        Else
            Check4.Value = 0
        End If
        If tRs.Fields("Pe5") = "S" Then
            Check5.Value = 1
        Else
            Check5.Value = 0
        End If
        If tRs.Fields("Pe6") = "S" Then
            Check6.Value = 1
        Else
            Check6.Value = 0
        End If
        If tRs.Fields("Pe7") = "S" Then
            Check7.Value = 1
        Else
            Check7.Value = 0
        End If
        If tRs.Fields("Pe8") = "S" Then
            Check8.Value = 1
        Else
            Check8.Value = 0
        End If
        If tRs.Fields("Pe9") = "S" Then
            Check9.Value = 1
        Else
            Check9.Value = 0
        End If
        If tRs.Fields("Pe10") = "S" Then
            Check10.Value = 1
        Else
            Check10.Value = 0
        End If
        If tRs.Fields("Pe11") = "S" Then
            Check11.Value = 1
        Else
            Check11.Value = 0
        End If
        If tRs.Fields("Pe12") = "S" Then
            Check12.Value = 1
        Else
            Check12.Value = 0
        End If
        If tRs.Fields("Pe13") = "S" Then
            Check13.Value = 1
        Else
            Check13.Value = 0
        End If
        If tRs.Fields("Pe14") = "S" Then
            Check14.Value = 1
        Else
            Check14.Value = 0
        End If
        If tRs.Fields("Pe15") = "S" Then
            Check15.Value = 1
        Else
            Check15.Value = 0
        End If
        If tRs.Fields("Pe16") = "S" Then
            Check16.Value = 1
        Else
            Check16.Value = 0
        End If
        If tRs.Fields("Pe17") = "S" Then
            Check17.Value = 1
        Else
            Check17.Value = 0
        End If
        If tRs.Fields("Pe18") = "S" Then
            Check18.Value = 1
        Else
            Check18.Value = 0
        End If
        If tRs.Fields("Pe19") = "S" Then
            Check19.Value = 1
        Else
            Check19.Value = 0
        End If
        If tRs.Fields("Pe20") = "S" Then
            Check20.Value = 1
        Else
            Check20.Value = 0
        End If
        If tRs.Fields("Pe21") = "S" Then
            Check21.Value = 1
        Else
            Check21.Value = 0
        End If
        If tRs.Fields("Pe22") = "S" Then
            Check22.Value = 1
        Else
            Check22.Value = 0
        End If
        If tRs.Fields("Pe23") = "S" Then
            Check23.Value = 1
        Else
            Check23.Value = 0
        End If
        If tRs.Fields("Pe24") = "S" Then
            Check24.Value = 1
        Else
            Check24.Value = 0
        End If
        If tRs.Fields("Pe25") = "S" Then
            Check25.Value = 1
        Else
            Check25.Value = 0
        End If
        If tRs.Fields("Pe26") = "S" Then
            Check26.Value = 1
        Else
            Check26.Value = 0
        End If
        If tRs.Fields("Pe27") = "S" Then
            Check27.Value = 1
        Else
            Check27.Value = 0
        End If
        If tRs.Fields("Pe28") = "S" Then
            Check28.Value = 1
        Else
            Check28.Value = 0
        End If
        If tRs.Fields("Pe29") = "S" Then
            Check29.Value = 1
        Else
            Check29.Value = 0
        End If
        If tRs.Fields("Pe30") = "S" Then
            Check30.Value = 1
        Else
            Check30.Value = 0
        End If
        If tRs.Fields("Pe31") = "S" Then
            Check31.Value = 1
        Else
            Check31.Value = 0
        End If
        If tRs.Fields("Pe32") = "S" Then
            Check32.Value = 1
        Else
            Check32.Value = 0
        End If
        If tRs.Fields("Pe33") = "S" Then
            Check33.Value = 1
        Else
            Check33.Value = 0
        End If
        If tRs.Fields("Pe34") = "S" Then
            Check34.Value = 1
        Else
            Check34.Value = 0
        End If
        If tRs.Fields("Pe35") = "S" Then
            Check35.Value = 1
        Else
            Check35.Value = 0
        End If
        If tRs.Fields("Pe36") = "S" Then
            Check36.Value = 1
        Else
            Check36.Value = 0
        End If
        If tRs.Fields("Pe37") = "S" Then
            Check37.Value = 1
        Else
            Check37.Value = 0
        End If
        If tRs.Fields("Pe38") = "S" Then
            Check38.Value = 1
        Else
            Check38.Value = 0
        End If
        If tRs.Fields("Pe39") = "S" Then
            Check39.Value = 1
        Else
            Check39.Value = 0
        End If
        If tRs.Fields("Pe40") = "S" Then
            Check40.Value = 1
        Else
            Check40.Value = 0
        End If
        If tRs.Fields("Pe41") = "S" Then
            Check41.Value = 1
        Else
            Check41.Value = 0
        End If
        If tRs.Fields("Pe42") = "S" Then
            Check42.Value = 1
        Else
            Check42.Value = 0
        End If
        If tRs.Fields("Pe43") = "S" Then
            Check43.Value = 1
        Else
            Check43.Value = 0
        End If
        If tRs.Fields("Pe44") = "S" Then
            Check44.Value = 1
        Else
            Check44.Value = 0
        End If
        If tRs.Fields("Pe45") = "S" Then
            Check45.Value = 1
        Else
            Check45.Value = 0
        End If
        If tRs.Fields("Pe46") = "S" Then
            Check46.Value = 1
        Else
            Check46.Value = 0
        End If
        If tRs.Fields("Pe47") = "S" Then
            Check47.Value = 1
        Else
            Check47.Value = 0
        End If
        If tRs.Fields("Pe48") = "S" Then
            Check48.Value = 1
        Else
            Check48.Value = 0
        End If
        If tRs.Fields("Pe49") = "S" Then
            Check49.Value = 1
        Else
            Check49.Value = 0
        End If
        If tRs.Fields("Pe50") = "S" Then
            Check50.Value = 1
        Else
            Check50.Value = 0
        End If
        If tRs.Fields("Pe51") = "S" Then
            Check51.Value = 1
        Else
            Check51.Value = 0
        End If
        If tRs.Fields("Pe52") = "S" Then
            Check52.Value = 1
        Else
            Check52.Value = 0
        End If
        If tRs.Fields("Pe53") = "S" Then
            Check53.Value = 1
        Else
            Check53.Value = 0
        End If
        If tRs.Fields("Pe54") = "S" Then
            Check54.Value = 1
        Else
            Check54.Value = 0
        End If
        If tRs.Fields("Pe55") = "S" Then
            Check55.Value = 1
        Else
            Check55.Value = 0
        End If
        If tRs.Fields("Pe56") = "S" Then
            Check56.Value = 1
        Else
            Check56.Value = 0
        End If
        If tRs.Fields("Pe57") = "S" Then
            Check57.Value = 1
        Else
            Check57.Value = 0
        End If
        If tRs.Fields("Pe58") = "S" Then
            Check58.Value = 1
        Else
            Check58.Value = 0
        End If
        If tRs.Fields("Pe59") = "S" Then
            Check59.Value = 1
        Else
            Check59.Value = 0
        End If
        If tRs.Fields("Pe60") = "S" Then
            Check60.Value = 1
        Else
            Check60.Value = 0
        End If
        If tRs.Fields("Pe61") = "S" Then
            Check61.Value = 1
        Else
            Check61.Value = 0
        End If
        If tRs.Fields("Pe62") = "S" Then
            Check62.Value = 1
        Else
            Check62.Value = 0
        End If
        If tRs.Fields("Pe63") = "S" Then
            Check63.Value = 1
        Else
            Check63.Value = 0
        End If
        If tRs.Fields("Pe64") = "S" Then
            Check64.Value = 1
        Else
            Check64.Value = 0
        End If
        If tRs.Fields("Pe65") = "S" Then
            Check65.Value = 1
        Else
            Check65.Value = 0
        End If
        If tRs.Fields("Pe66") = "S" Then
            Check66.Value = 1
        Else
            Check66.Value = 0
        End If
        If tRs.Fields("Pe67") = "S" Then
            Check67.Value = 1
        Else
            Check67.Value = 0
        End If
        If tRs.Fields("Pe68") = "S" Then
            Check68.Value = 1
        Else
            Check68.Value = 0
        End If
        If tRs.Fields("Pe69") = "S" Then
            Check69.Value = 1
        Else
            Check69.Value = 0
        End If
        If tRs.Fields("Pe70") = "S" Then
            Check70.Value = 1
        Else
            Check70.Value = 0
        End If
    End If
End Sub
Private Sub Te1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Te2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Te3_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Te4_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Te6_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Buscar()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_USUARIO, NOMBRE, APELLIDOS FROM USUARIOS WHERE NOMBRE LIKE '%" & Text5.Text & "%' AND ESTADO = 'A'"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_USUARIO") & "")
            If Not IsNull(tRs.Fields("NOMBRE")) And Not IsNull(tRs.Fields("APELLIDOS")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE") & " " & tRs.Fields("APELLIDOS")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Text5_GotFocus()
    Text5.BackColor = &HFFE1E1
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Buscar
    End If
End Sub
Private Sub Text5_LostFocus()
    Text5.BackColor = &H80000005
End Sub
