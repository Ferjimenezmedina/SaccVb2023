VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmRequisicion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REQUISICIÓN"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Juntar productos con requisiciones pendientes."
      Height          =   255
      Left            =   6120
      TabIndex        =   55
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   43
      Top             =   3240
      Width           =   975
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
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
         TabIndex        =   44
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "frmRequisicion.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisicion.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "EN SUCURSAL"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "EXISTENCIA"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame26 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   37
      Top             =   4440
      Width           =   975
      Begin VB.Image Image24 
         Height          =   765
         Left            =   240
         MouseIcon       =   "frmRequisicion.frx":1EDC
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisicion.frx":21E6
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "En Compra"
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
         TabIndex        =   38
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9960
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtMarca 
      Height          =   285
      Left            =   5400
      TabIndex        =   36
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   34
      Top             =   5640
      Width           =   975
      Begin VB.Label Label8 
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
         TabIndex        =   35
         Top             =   960
         Width           =   975
      End
      Begin VB.Image cmdEnviar 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmRequisicion.frx":3C74
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisicion.frx":3F7E
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   32
      Top             =   6840
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmRequisicion.frx":5940
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisicion.frx":5C4A
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label27 
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
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "COMENTARIO"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtComentario 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   2520
      Width           =   8295
   End
   Begin VB.TextBox txtAlmacen 
      Height          =   285
      Left            =   6000
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtIdProducto 
      Height          =   285
      Left            =   6600
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox txtIndice 
      Height          =   285
      Left            =   5400
      TabIndex        =   28
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdSurtir1 
      Caption         =   "Surtir"
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
      Left            =   8400
      Picture         =   "frmRequisicion.frx":7D2C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Requi"
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
      Left            =   6960
      Picture         =   "frmRequisicion.frx":A6FE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   8281
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Pedidos"
      TabPicture(0)   =   "frmRequisicion.frx":D0D0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdQuitar2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdQuitar1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Requisición"
      TabPicture(1)   =   "frmRequisicion.frx":D0EC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "cmdBorrar1"
      Tab(1).Control(2)=   "lvProd"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Surtidos"
      TabPicture(2)   =   "frmRequisicion.frx":D108
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "cmdBorrar2"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame5 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   53
         Top             =   480
         Width           =   9135
         Begin MSComctlLib.ListView lvwSurtir 
            Height          =   3255
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   5741
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.CommandButton cmdBorrar2 
         Caption         =   "Borrar"
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
         Left            =   -66840
         Picture         =   "frmRequisicion.frx":D124
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   49
         Top             =   480
         Width           =   9135
         Begin MSComctlLib.ListView lvwOrdenCompra 
            Height          =   3255
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   5741
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.CommandButton cmdBorrar1 
         Caption         =   "Borrar"
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
         Left            =   -66840
         Picture         =   "frmRequisicion.frx":FAF6
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Height          =   3615
         Left            =   240
         TabIndex        =   46
         Top             =   480
         Width           =   9135
         Begin MSComctlLib.ListView lvwDirectas 
            Height          =   3255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   5741
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483633
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.CommandButton cmdQuitar1 
         Caption         =   "Quitar"
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
         Left            =   8160
         Picture         =   "frmRequisicion.frx":124C8
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuitar2 
         Caption         =   "Quitar"
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
         Left            =   600
         Picture         =   "frmRequisicion.frx":14E9A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   4200
         Visible         =   0   'False
         Width           =   495
         Begin MSComctlLib.ListView lvwIndirectas 
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin MSComctlLib.ListView lvProd 
         Height          =   270
         Left            =   -74760
         TabIndex        =   51
         Top             =   4200
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   476
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Appearance      =   0
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NUMERO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID_PRODUCTO"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DESCRIPCIÓN"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "CANTIDAD"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID PEDIDO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "CANTIDAD_PEDIDO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Origen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Sucursal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Agente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Fecha"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "ALMACEN"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "TIPO"
            Object.Width           =   176
         EndProperty
      End
   End
   Begin VB.TextBox txtCantidad_Pedido 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5400
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtId 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6000
      TabIndex        =   24
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total pedidos"
      Height          =   855
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtPI 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "DIRECTOS"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtPD 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "INDIRECTOS"
         Top             =   360
         Visible         =   0   'False
         Width           =   75
      End
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "DESCRIPCIÓN"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtDescripcion 
      BackColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   6600
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox txtCantidad 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6600
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "CANTIDAD"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "ID PRODUCTO"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "FECHA"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtAgente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "AGENTE"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtSucursal 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "SUCURSAL"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "PEDIDO"
      Top             =   1080
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFechaRequisicion 
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51118081
      CurrentDate     =   38840
   End
   Begin VB.TextBox txtPedido 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   10080
      Top             =   2640
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "frmRequisicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim ItMx As ListItem
Dim bLvw As Byte
Dim nCantidad_Pedido As Double
Dim nExistencia As Double
Dim bLOR As Byte
Dim Con As Integer
Dim NumReg As Integer
Dim Cantidad_Acumulada As Double
Dim i As Integer
Private Sub cmdAgregar_Click()
On Error GoTo ManejaError
    Dim nrOrden As Integer
    Dim contOrden As Integer
    Dim ItMx As ListItem
    Dim bNo As Boolean
    Dim Origen As Integer
    If bLvw = 1 Then
        If Val(txtCantidad.Text) >= Val(lvwDirectas.ListItems.Item(Val(txtIndice.Text)).ListSubItems(5)) Then
            lvwDirectas.ListItems.Remove (Val(txtIndice.Text))
        Else
            lvwDirectas.ListItems.Item(Val(txtIndice.Text)).ListSubItems(5) = Val(lvwDirectas.ListItems.Item(Val(txtIndice.Text)).ListSubItems(5)) - Val(txtCantidad.Text)
        End If
        Origen = 1
    Else
        If bLvw = 2 Then
            If Val(txtCantidad.Text) >= Val(lvwIndirectas.ListItems.Item(Val(txtIndice.Text)).ListSubItems(5)) Then
                lvwIndirectas.ListItems.Remove (Val(txtIndice.Text))
            Else
                lvwIndirectas.ListItems.Item(Val(txtIndice.Text)).ListSubItems(5) = Val(lvwIndirectas.ListItems.Item(Val(txtIndice.Text)).ListSubItems(5)) - Val(txtCantidad.Text)
            End If
            Origen = 2
        End If
    End If
    bNo = False
    If Puede_Agregar Then
        Agregar_Lista_Ordenes (Origen)
        Me.cmdSurtir1.Enabled = False
        Me.cmdAgregar.Enabled = False
        Me.txtAgente.Text = ""
        Me.txtCantidad.Text = ""
        Me.txtDescripcion.Text = ""
        Me.txtFecha.Text = ""
        Me.txtIdProducto.Text = ""
        Me.txtPedido.Text = ""
        Me.txtPI.Text = ""
        Me.txtSucursal.Text = ""
        txtIndice.Text = ""
        txtMarca.Text = ""
        txtPD.Text = lvwDirectas.ListItems.Count
        txtPI.Text = lvwIndirectas.ListItems.Count
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdBorrar1_Click()
On Error GoTo ManejaError
    Dim tLi As ListItem
    Dim Cont As Integer
    Dim Tipo As String
    Dim Index As Integer
    Cont = 1
    If bLOR = 1 Then
        Tipo = lvwOrdenCompra.ListItems.Item(i).SubItems(12)
        If lvwOrdenCompra.SelectedItem.SubItems(7) = 2 Then
            If lvwOrdenCompra.ListItems.Item(i).SubItems(6) <= lvwOrdenCompra.ListItems.Item(i).SubItems(3) Then
                Set tLi = lvwIndirectas.ListItems.Add(, , lvwOrdenCompra.ListItems.Item(i).SubItems(5) & "") 'Me.txtPedido.Text
                    tLi.SubItems(1) = lvwOrdenCompra.ListItems.Item(i).SubItems(8)
                    tLi.SubItems(2) = lvwOrdenCompra.ListItems.Item(i).SubItems(9)
                    tLi.SubItems(3) = lvwOrdenCompra.ListItems.Item(i).SubItems(10)
                    tLi.SubItems(4) = lvwOrdenCompra.ListItems.Item(i).SubItems(1)
                    tLi.SubItems(5) = lvwOrdenCompra.ListItems.Item(i).SubItems(3)
                    tLi.SubItems(6) = lvwOrdenCompra.ListItems.Item(i).SubItems(2)
                    tLi.SubItems(7) = lvwOrdenCompra.ListItems.Item(i).SubItems(4)
            Else
                If lvwIndirectas.ListItems.Count > 0 Then
                    Do While (lvwOrdenCompra.ListItems.Item(i).SubItems(5) <> lvwIndirectas.ListItems.Item(Cont)) And Cont <= lvwIndirectas.ListItems.Count
                        Cont = Cont + 1
                    Loop
                End If
                If Cont > lvwIndirectas.ListItems.Count Or lvwIndirectas.ListItems.Count = 0 Then
                    Set tLi = lvwIndirectas.ListItems.Add(, , lvwOrdenCompra.ListItems.Item(i).SubItems(5) & "") 'Me.txtPedido.Text
                    tLi.SubItems(1) = lvwOrdenCompra.ListItems.Item(i).SubItems(8)
                    tLi.SubItems(2) = lvwOrdenCompra.ListItems.Item(i).SubItems(9)
                    tLi.SubItems(3) = lvwOrdenCompra.ListItems.Item(i).SubItems(10)
                    tLi.SubItems(4) = lvwOrdenCompra.ListItems.Item(i).SubItems(1)
                    tLi.SubItems(5) = lvwOrdenCompra.ListItems.Item(i).SubItems(3)
                    tLi.SubItems(6) = lvwOrdenCompra.ListItems.Item(i).SubItems(2)
                    tLi.SubItems(7) = lvwOrdenCompra.ListItems.Item(i).SubItems(4)
                Else
                    lvwIndirectas.ListItems.Item(Cont).SubItems(5) = Val(lvwIndirectas.ListItems.Item(Cont).SubItems(5)) + Val(lvwOrdenCompra.ListItems.Item(i).SubItems(3))
                End If
            End If
        Else
            If lvwOrdenCompra.ListItems.Item(i).SubItems(6) <= lvwOrdenCompra.ListItems.Item(i).SubItems(3) Then
                Set tLi = lvwDirectas.ListItems.Add(, , lvwOrdenCompra.ListItems.Item(i).SubItems(5) & "") 'Me.txtPedido.Text
                    tLi.SubItems(1) = lvwOrdenCompra.ListItems.Item(i).SubItems(8)
                    tLi.SubItems(2) = lvwOrdenCompra.ListItems.Item(i).SubItems(9)
                    tLi.SubItems(3) = lvwOrdenCompra.ListItems.Item(i).SubItems(10)
                    tLi.SubItems(4) = lvwOrdenCompra.ListItems.Item(i).SubItems(1)
                    tLi.SubItems(5) = lvwOrdenCompra.ListItems.Item(i).SubItems(3)
                    tLi.SubItems(6) = lvwOrdenCompra.ListItems.Item(i).SubItems(2)
                    tLi.SubItems(7) = lvwOrdenCompra.ListItems.Item(i).SubItems(4)
            Else
                If lvwDirectas.ListItems.Count > 0 Then
                    Do While (lvwOrdenCompra.ListItems.Item(i).SubItems(5) <> lvwDirectas.ListItems.Item(Cont)) And Cont <= lvwDirectas.ListItems.Count
                        Cont = Cont + 1
                    Loop
                End If
                If Cont > lvwDirectas.ListItems.Count Or lvwDirectas.ListItems.Count = 0 Then
                    Set tLi = lvwDirectas.ListItems.Add(, , lvwOrdenCompra.ListItems.Item(i).SubItems(5) & "") 'Me.txtPedido.Text
                    tLi.SubItems(1) = lvwOrdenCompra.ListItems.Item(i).SubItems(8)
                    tLi.SubItems(2) = lvwOrdenCompra.ListItems.Item(i).SubItems(9)
                    tLi.SubItems(3) = lvwOrdenCompra.ListItems.Item(i).SubItems(10)
                    tLi.SubItems(4) = lvwOrdenCompra.ListItems.Item(i).SubItems(1)
                    tLi.SubItems(5) = lvwOrdenCompra.ListItems.Item(i).SubItems(3)
                    tLi.SubItems(6) = lvwOrdenCompra.ListItems.Item(i).SubItems(2)
                    tLi.SubItems(7) = lvwOrdenCompra.ListItems.Item(i).SubItems(4)
                Else
                    lvwDirectas.ListItems.Item(Cont).SubItems(5) = Val(lvwIndirectas.ListItems.Item(Cont).SubItems(5)) + Val(lvwOrdenCompra.ListItems.Item(i).SubItems(3))
                End If
            End If
        End If
        Me.lvwOrdenCompra.ListItems.Remove (Me.lvwOrdenCompra.SelectedItem.Index)
    Else
        i = lvwSurtir.SelectedItem.Index
        If lvwSurtir.SelectedItem.SubItems(7) = 2 Then
            If lvwSurtir.ListItems.Item(i).SubItems(6) <= lvwSurtir.ListItems.Item(i).SubItems(3) Then
                Set tLi = lvwIndirectas.ListItems.Add(, , lvwSurtir.ListItems.Item(i).SubItems(5) & "") 'Me.txtPedido.Text
                    tLi.SubItems(1) = lvwSurtir.ListItems.Item(i).SubItems(8)
                    tLi.SubItems(2) = lvwSurtir.ListItems.Item(i).SubItems(9)
                    tLi.SubItems(3) = lvwSurtir.ListItems.Item(i).SubItems(10)
                    tLi.SubItems(4) = lvwSurtir.ListItems.Item(i).SubItems(1)
                    tLi.SubItems(5) = lvwSurtir.ListItems.Item(i).SubItems(3)
                    tLi.SubItems(6) = lvwSurtir.ListItems.Item(i).SubItems(2)
                    tLi.SubItems(7) = lvwSurtir.ListItems.Item(i).SubItems(4)
            Else
                If lvwIndirectas.ListItems.Count > 0 Then
                    Do While (lvwSurtir.ListItems.Item(i).SubItems(5) <> lvwIndirectas.ListItems.Item(Cont)) And Cont <= lvwIndirectas.ListItems.Count
                        Cont = Cont + 1
                    Loop
                End If
                If Cont > lvwIndirectas.ListItems.Count Or lvwIndirectas.ListItems.Count = 0 Then
                    Set tLi = lvwIndirectas.ListItems.Add(, , lvwSurtir.ListItems.Item(i).SubItems(5) & "") 'Me.txtPedido.Text
                    tLi.SubItems(1) = lvwSurtir.ListItems.Item(i).SubItems(8)
                    tLi.SubItems(2) = lvwSurtir.ListItems.Item(i).SubItems(9)
                    tLi.SubItems(3) = lvwSurtir.ListItems.Item(i).SubItems(10)
                    tLi.SubItems(4) = lvwSurtir.ListItems.Item(i).SubItems(1)
                    tLi.SubItems(5) = lvwSurtir.ListItems.Item(i).SubItems(3)
                    tLi.SubItems(6) = lvwSurtir.ListItems.Item(i).SubItems(2)
                    tLi.SubItems(7) = lvwSurtir.ListItems.Item(i).SubItems(4)
                Else
                    lvwIndirectas.ListItems.Item(Cont).SubItems(5) = Val(lvwIndirectas.ListItems.Item(Cont).SubItems(5)) + Val(lvwSurtir.ListItems.Item(i).SubItems(3))
                End If
            End If
        Else
            If lvwSurtir.ListItems.Item(i).SubItems(6) <= lvwSurtir.ListItems.Item(i).SubItems(3) Then
                Set tLi = lvwDirectas.ListItems.Add(, , lvwSurtir.ListItems.Item(i).SubItems(5) & "") 'Me.txtPedido.Text
                    tLi.SubItems(1) = lvwSurtir.ListItems.Item(i).SubItems(8)
                    tLi.SubItems(2) = lvwSurtir.ListItems.Item(i).SubItems(9)
                    tLi.SubItems(3) = lvwSurtir.ListItems.Item(i).SubItems(10)
                    tLi.SubItems(4) = lvwSurtir.ListItems.Item(i).SubItems(1)
                    tLi.SubItems(5) = lvwSurtir.ListItems.Item(i).SubItems(3)
                    tLi.SubItems(6) = lvwSurtir.ListItems.Item(i).SubItems(2)
                    tLi.SubItems(7) = lvwSurtir.ListItems.Item(i).SubItems(4)
            Else
                If lvwDirectas.ListItems.Count > 0 Then
                    Do While (lvwSurtir.ListItems.Item(i).SubItems(5) <> lvwDirectas.ListItems.Item(Cont)) And Cont <= lvwDirectas.ListItems.Count
                        Cont = Cont + 1
                    Loop
                End If
                If Cont > lvwDirectas.ListItems.Count Or lvwDirectas.ListItems.Count = 0 Then
                    Set tLi = lvwDirectas.ListItems.Add(, , lvwSurtir.ListItems.Item(i).SubItems(5) & "") 'Me.txtPedido.Text
                    tLi.SubItems(1) = lvwSurtir.ListItems.Item(i).SubItems(8)
                    tLi.SubItems(2) = lvwSurtir.ListItems.Item(i).SubItems(9)
                    tLi.SubItems(3) = lvwSurtir.ListItems.Item(i).SubItems(10)
                    tLi.SubItems(4) = lvwSurtir.ListItems.Item(i).SubItems(1)
                    tLi.SubItems(5) = lvwSurtir.ListItems.Item(i).SubItems(3)
                    tLi.SubItems(6) = lvwSurtir.ListItems.Item(i).SubItems(2)
                    tLi.SubItems(7) = lvwSurtir.ListItems.Item(i).SubItems(4)
                Else
                    lvwDirectas.ListItems.Item(Cont).SubItems(5) = Val(lvwIndirectas.ListItems.Item(Cont).SubItems(5)) + Val(lvwSurtir.ListItems.Item(i).SubItems(3))
                End If
            End If
        End If
        Me.lvwSurtir.ListItems.Remove (Me.lvwSurtir.SelectedItem.Index)
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdBorrar2_Click()
    cmdBorrar1.Value = True
End Sub
Private Sub cmdEnviar_Click()
On Error GoTo ManejaError
    Dim C As Integer
    Dim NR As Integer
    Dim ID_PRODUCTO As String
    Dim Descripcion As String
    Dim CANTIDAD As Double
    Dim CANTIDAD_REQUISICION As Double
    Dim ID As Double
    Dim ID_PEDIDO As Double
    Dim Almacen As String
    Dim Marca As String
    Dim Sucursal As String
    Dim sBuscar As String
    Dim sBuscar2 As String
    Dim tRs As ADODB.Recordset
    Dim Control As Boolean
    Dim Id_Prod As Integer
    Dim cTipo As String
    Dim D As Integer
    Dim Requi As Boolean
    Dim X As Integer
    Dim Y As Integer
    Dim contar As Integer
    Dim numCopies As Integer
    Y = 2
    If Me.lvwSurtir.ListItems.Count <> 0 Then
        NR = Me.lvwSurtir.ListItems.Count
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "                                                                                          PRODUCTOS SURTIDOS"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        POSY = 2200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 2200
        Printer.Print "Descripcion"
        Printer.CurrentY = POSY
        Printer.CurrentX = 9000
        Printer.Print "CANTIDAD"
        Printer.CurrentY = POSY
        Printer.CurrentX = 9500
        Printer.Print "SUCURSAL"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        POSY = POSY + 200
        'CommonDialog1.Copies = 1
        For C = 1 To NR
            ID_PRODUCTO = Me.lvwSurtir.ListItems.Item(C).SubItems(1)
            CANTIDAD = Me.lvwSurtir.ListItems.Item(C).SubItems(3)
            Sucursal = Me.lvwSurtir.ListItems.Item(C).SubItems(8)
            ID = Me.lvwSurtir.ListItems.Item(C).SubItems(4)
            ID_PEDIDO = Me.lvwSurtir.ListItems.Item(C).SubItems(5)
            sBuscar = "INSERT INTO SURTIDOS(FECHA, ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', '" & ID_PRODUCTO & "', " & CANTIDAD & ", '" & Sucursal & "');"
            cnn.Execute (sBuscar)
            sBuscar = "UPDATE DETALLE_PEDIDO SET ENTREGADO = 'S', CANTIDAD = " & CANTIDAD & " WHERE ID = " & ID & " AND ID_PEDIDO = " & ID_PEDIDO
            cnn.Execute (sBuscar)
            sBuscar = "SELECT * FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ID_PRODUCTO & "' AND SUCURSAL = '" & Sucursal & "' AND CANTIDAD >= " & CANTIDAD
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs.Fields("CANTIDAD")) - CDbl(CANTIDAD) & " WHERE ID_PRODUCTO = '" & ID_PRODUCTO & "' AND SUCURSAL = 'BODEGA'"
                cnn.Execute (sBuscar)
                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs.Fields("CANTIDAD")) + CDbl(CANTIDAD) & " WHERE ID_PRODUCTO = '" & ID_PRODUCTO & "' AND SUCURSAL = '" & Sucursal & "'"
                cnn.Execute (sBuscar)
            Else
                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD - " & CDbl(CANTIDAD) & " WHERE ID_PRODUCTO = '" & ID_PRODUCTO & "' AND SUCURSAL = 'BODEGA'"
                cnn.Execute (sBuscar)
                sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & ID_PRODUCTO & "', " & CANTIDAD & ", '" & Sucursal & "')"
                cnn.Execute (sBuscar)
            End If
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print ID_PRODUCTO
            Printer.CurrentY = POSY
            Printer.CurrentX = 2200
            Printer.Print lvwSurtir.ListItems.Item(C).SubItems(2)
            Printer.CurrentY = POSY
            Printer.CurrentX = 9000
            Printer.Print CANTIDAD
            Printer.CurrentY = POSY
            Printer.CurrentX = 9500
            Printer.Print Sucursal
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 0
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            If POSY >= 14200 Then
                Printer.NewPage
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                Printer.Print VarMen.Text5(0).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print "                                                                                          PRODUCTOS SURTIDOS"
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                POSY = 2200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print "Producto"
                Printer.CurrentY = POSY
                Printer.CurrentX = 2200
                Printer.Print "Descripcion"
                Printer.CurrentY = POSY
                Printer.CurrentX = 9000
                Printer.Print "CANTIDAD"
                Printer.CurrentY = POSY
                Printer.CurrentX = 9500
                Printer.Print "SUCURSAL"
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                POSY = POSY + 200
            End If
        Next C
        Printer.EndDoc
        Me.lvwSurtir.ListItems.Clear
    End If
    D = 0
    If Me.lvwOrdenCompra.ListItems.Count <> 0 Then
        NR = Me.lvwOrdenCompra.ListItems.Count
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "                                                                                          PRODUCTOS EN REQUISICION"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        POSY = 2200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 2200
        Printer.Print "Descripcion"
        Printer.CurrentY = POSY
        Printer.CurrentX = 9000
        Printer.Print "CANTIDAD"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        POSY = POSY + 200
        For C = 1 To NR
            ID_PRODUCTO = lvwOrdenCompra.ListItems.Item(C).SubItems(1)
            Descripcion = lvwOrdenCompra.ListItems.Item(C).SubItems(2)
            CANTIDAD = lvwOrdenCompra.ListItems.Item(C).SubItems(3)
            ID = lvwOrdenCompra.ListItems.Item(C).SubItems(4)
            ID_PEDIDO = lvwOrdenCompra.ListItems.Item(C).SubItems(5)
            Almacen = lvwOrdenCompra.ListItems.Item(C).SubItems(11)
            Marca = lvwOrdenCompra.ListItems.Item(C).SubItems(13)
            Requi = False
            If Trim(lvwOrdenCompra.ListItems.Item(C).SubItems(12)) = "SIMPLE" Then
                Requi = True
            Else
                If MsgBox("DESEA MANDAR EL " & ID_PRODUCTO & " A REQUISICION?", vbYesNo, "SACC") = vbYes Then
                    Requi = True
                End If
            End If
            If Requi Then
                If Check1.Value = 1 Then
                    sBuscar = "SELECT ID_REQUISICION, CANTIDAD FROM REQUISICION WHERE ACTIVO = 0 AND URGENTE = 'N' AND ID_PRODUCTO = '" & ID_PRODUCTO & "'"
                    Set tRs = cnn.Execute(sBuscar)
                    If tRs.BOF And tRs.EOF Then
                        sBuscar = "INSERT INTO REQUISICION (FECHA,ID_PRODUCTO,Descripcion,CANTIDAD,CONTADOR,ALMACEN,MARCA,COMENTARIO) Values('" & Format(Date, "dd/mm/yyyy") & "', '" & ID_PRODUCTO & "', '" & Descripcion & "', " & CANTIDAD & ", 0, '" & Almacen & "', '& Marca &','" & txtComentario.Text & "')"
                        sBuscar2 = ""
                    Else
                        sBuscar = "UPDATE REQUISICION SET CANTIDAD = " & Replace(Val(tRs.Fields("CANTIDAD")) + Val(CANTIDAD), ",", "") & "WHERE ID_REQUISICION = " & tRs.Fields("ID_REQUISICION") & " AND ID_PRODUCTO = '" & ID_PRODUCTO & "'"
                        sBuscar2 = "INSERT INTO RASTREOREQUI (ID_REQUI, SOLICITO, CANTIDAD, COMENTARIO,FECHA) Values(" & tRs.Fields("ID_REQUISICION") & ", '" & VarMen.Text1(1).Text & "'  ," & CANTIDAD & ", 'REQUISICION HECHA DESDE UN PEDIDO','" & Format(Date, "dd/mm/yyyy") & "')"
                    End If
                Else
                    sBuscar = "INSERT INTO REQUISICION (FECHA,ID_PRODUCTO,Descripcion,CANTIDAD,CONTADOR,ALMACEN,MARCA,COMENTARIO) Values('" & Format(Date, "dd/mm/yyyy") & "', '" & ID_PRODUCTO & "', '" & Descripcion & "', '" & CANTIDAD & "', 0, '" & Almacen & "', '& Marca &','" & txtComentario.Text & "')"
                    sBuscar2 = ""
                End If
              '  tRs.Close
                cnn.Execute (sBuscar)
                If sBuscar2 = "" Then
                    sBuscar = "SELECT ID_REQUISICION FROM REQUISICION ORDER BY ID_REQUISICION DESC"
                    Set tRs = cnn.Execute(sBuscar)
                    sBuscar2 = "INSERT INTO RASTREOREQUI (ID_REQUI, SOLICITO, CANTIDAD, COMENTARIO,FECHA) Values(" & tRs.Fields("ID_REQUISICION") & ", '" & VarMen.Text1(1).Text & "'  ," & CANTIDAD & ", 'REQUISICION HECHA DESDE UN PEDIDO','" & Format(Date, "dd/mm/yyyy") & "')"
                    tRs.Close
                End If
                cnn.Execute (sBuscar2)
                sBuscar = "UPDATE DETALLE_PEDIDO SET ENTREGADO = 'R', CANTIDAD = " & CANTIDAD & " WHERE ID = " & ID & " AND ID_PEDIDO = " & ID_PEDIDO
                cnn.Execute (sBuscar)
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print ID_PRODUCTO
                Printer.CurrentY = POSY
                Printer.CurrentX = 2200
                Printer.Print Descripcion
                Printer.CurrentY = POSY
                Printer.CurrentX = 9000
                Printer.Print CANTIDAD
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 0
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Else
                D = lvProd.ListItems.Count + 1
                Set ItMx = lvProd.ListItems.Add(, , D)
                ItMx.SubItems(1) = lvwOrdenCompra.ListItems.Item(C).SubItems(1)
                ItMx.SubItems(2) = lvwOrdenCompra.ListItems.Item(C).SubItems(2)
                ItMx.SubItems(3) = lvwOrdenCompra.ListItems.Item(C).SubItems(3)
                ItMx.SubItems(4) = lvwOrdenCompra.ListItems.Item(C).SubItems(4)
                ItMx.SubItems(5) = lvwOrdenCompra.ListItems.Item(C).SubItems(5)
                ItMx.SubItems(6) = lvwOrdenCompra.ListItems.Item(C).SubItems(6)
                ItMx.SubItems(7) = lvwOrdenCompra.ListItems.Item(C).SubItems(7)
                ItMx.SubItems(8) = lvwOrdenCompra.ListItems.Item(C).SubItems(8)
                ItMx.SubItems(9) = lvwOrdenCompra.ListItems.Item(C).SubItems(9)
                ItMx.SubItems(10) = lvwOrdenCompra.ListItems.Item(C).SubItems(10)
                ItMx.SubItems(11) = lvwOrdenCompra.ListItems.Item(C).SubItems(11)
                ItMx.SubItems(12) = lvwOrdenCompra.ListItems.Item(C).SubItems(12)
            End If
            If POSY >= 14200 Then
                Printer.NewPage
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                Printer.Print VarMen.Text5(0).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print "                                                                                          PRODUCTOS EN REQUISICION"
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                POSY = 2200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print "Producto"
                Printer.CurrentY = POSY
                Printer.CurrentX = 2200
                Printer.Print "Descripcion"
                Printer.CurrentY = POSY
                Printer.CurrentX = 9000
                Printer.Print "CANTIDAD"
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                POSY = POSY + 200
            End If
        Next C
    'pr.Copies = numCopies
        Printer.EndDoc
        'Next contar
        Me.lvwOrdenCompra.ListItems.Clear
    End If
    If lvProd.ListItems.Count > 0 Then
        sqlQuery = "INSERT INTO COMANDAS_2 (FECHA_INICIO, ID_AGENTE, ID_SUCURSAL, TIPO, SUCURSAL) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', " & VarMen.Text1(0).Text & ", " & VarMen.Text1(5).Text & ", 'P', '" & VarMen.Text4(0).Text & "')"
        cnn.Execute (sqlQuery)
        sqlQuery = "SELECT TOP 1 ID_COMANDA FROM COMANDAS_2 ORDER BY ID_COMANDA DESC"
        Set tRs = cnn.Execute(sqlQuery)
        Id_Prod = tRs.Fields("ID_COMANDA")
        For C = 1 To NR
            ID_PRODUCTO = lvProd.ListItems.Item(C).SubItems(1)
            Descripcion = lvProd.ListItems.Item(C).SubItems(2)
            CANTIDAD = lvProd.ListItems.Item(C).SubItems(3)
            ID = lvProd.ListItems.Item(C).SubItems(4)
            ID_PEDIDO = lvProd.ListItems.Item(C).SubItems(5)
            Almacen = lvProd.ListItems.Item(C).SubItems(11)
            If Mid(ID_PRODUCTO, 3, 1) = "T" Then
                cTipo = "T" 'Toner
            ElseIf Mid(ID_PRODUCTO, 3, 1) = "I" Then
                cTipo = "I" 'Tinta
            Else
                cTipo = "X" 'Error
            End If
            sqlQuery = "INSERT INTO COMANDAS_DETALLES_2 (ID_COMANDA, ARTICULO, ID_PRODUCTO, CANTIDAD, TIPO) VALUES (" & Id_Prod & ", " & C & ", '" & ID_PRODUCTO & "', " & CANTIDAD & ", '" & cTipo & "');"
            cnn.Execute (sqlQuery)
            sqlQuery = "INSERT INTO PRODPEND (ID_COMANDA, ARTICULO) VALUES (" & Id_Prod & ", " & C & ");"
            cnn.Execute (sqlQuery)
            sBuscar = "UPDATE DETALLE_PEDIDO SET ENTREGADO = 'R', CANTIDAD = " & CANTIDAD & " WHERE ID = " & ID & " AND ID_PEDIDO = " & ID_PEDIDO
            cnn.Execute (sBuscar)
        Next C
        Printer.Print "     " & VarMen.Text5(0).Text
        Printer.Print "           ORDEN DE PRODUCCIÓN"
        Printer.Print "FECHA : " & Now
        Printer.Print "No. DE COMANDA : " & Id_Prod
        Printer.Print "ORDEN HECHA POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                           ORDEN DE TINTA"
        NR = lvProd.ListItems.Count
        Dim Con As Integer
        POSY = 2200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3000
        Printer.Print "Cant."
        For Con = 1 To NR
            If Mid(lvProd.ListItems.Item(Con).SubItems(1), 3, 1) = "I" Then
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print lvProd.ListItems.Item(Con).SubItems(1)
                Printer.CurrentY = POSY
                Printer.CurrentX = 2900
                Printer.Print lvProd.ListItems.Item(Con).SubItems(3)
            End If
        Next Con
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                           ORDEN DE TONER"
        POSY = POSY + 600
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3000
        Printer.Print "Cant."
        For Con = 1 To NR
            If Mid(lvProd.ListItems.Item(Con).SubItems(1), 3, 1) = "T" Then
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print lvProd.ListItems.Item(Con).SubItems(1)
                Printer.CurrentY = POSY
                Printer.CurrentX = 2900
                Printer.Print lvProd.ListItems.Item(Con).SubItems(3)
            End If
        Next Con
        Printer.Print ""
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print ""
        Printer.EndDoc
        'NESESITA 2 TICKETS
        Printer.Print "    " & VarMen.Text5(0).Text
        Printer.Print "           ORDEN DE PRODUCCIÓN"
        Printer.Print "FECHA : " & Now
        Printer.Print "No. DE COMANDA : " & Id_Prod
        Printer.Print "ORDEN HECHA POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                           ORDEN DE TINTA"
        NR = lvProd.ListItems.Count
        POSY = 2200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3000
        Printer.Print "Cant."
        For Con = 1 To NR
            If Mid(lvProd.ListItems.Item(Con).SubItems(1), 3, 1) = "I" Then
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print lvProd.ListItems.Item(Con).SubItems(1)
                Printer.CurrentY = POSY
                Printer.CurrentX = 2900
                Printer.Print lvProd.ListItems.Item(Con).SubItems(3)
            End If
        Next Con
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                           ORDEN DE TONER"
        POSY = POSY + 600
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3000
        Printer.Print "Cant."
        For Con = 1 To NR
            If Mid(lvProd.ListItems.Item(Con).SubItems(1), 3, 1) = "T" Then
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print lvProd.ListItems.Item(Con).SubItems(1)
                Printer.CurrentY = POSY
                Printer.CurrentX = 2900
                Printer.Print lvProd.ListItems.Item(Con).SubItems(3)
            End If
        Next Con
        Printer.Print ""
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print ""
        Printer.EndDoc
        lvProd.ListItems.Clear
    End If
    Me.Llenar_Lista_Pedidos_Directos
    Me.Llenar_Lista_Pedidos_Indirectos
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdQuitar1_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    sBuscar = "UPDATE DETALLE_PEDIDO SET ENTREGADO = 'I', CANTIDAD = " & Me.txtCantidad.Text & " WHERE ID = " & Me.txtID.Text & " AND ID_PEDIDO = " & Me.txtPedido.Text
    cnn.Execute (sBuscar)
    'lvwIndirectas.ListItems.Remove lvwIndirectas.SelectedItem.Index
    txtPD.Text = lvwDirectas.ListItems.Count
    txtPI.Text = lvwIndirectas.ListItems.Count
    Llenar_Lista_Pedidos_Directos
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdQuitar2_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    sBuscar = "UPDATE DETALLE_PEDIDO SET ENTREGADO = 'I', CANTIDAD = " & Me.txtCantidad.Text & " WHERE ID = " & Me.txtID.Text & " AND ID_PEDIDO = " & Me.txtPedido.Text
    cnn.Execute (sBuscar)
    lvwDirectas.ListItems.Remove lvwDirectas.SelectedItem.Index
    txtPD.Text = lvwDirectas.ListItems.Count
    txtPI.Text = lvwIndirectas.ListItems.Count
    Llenar_Lista_Pedidos_Indirectos
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdSurtir1_Click()
On Error GoTo ManejaError
    Dim Origen As Integer
    Dim Quitar As Boolean
    If Puede_Surtir Then
        Quitar = False
        If Val(Me.txtCantidad_Pedido.Text) >= Val(Me.txtCantidad.Text) Then
            If MsgBox("DESEA HACER UNA REQUISICION POR " & Val(Me.txtCantidad_Pedido.Text), vbYesNo + vbDefaultButton1 + vbQuestion, "SACC") = vbYes Then
                If bLvw = 1 Then
                    Origen = 1
                Else
                    If bLvw = 2 Then
                        Origen = 2
                    End If
                End If
                
                Dim C As Integer
                If txtAlmacen.Text = "A3" Then
                    sBuscar = "SELECT isNull(TIPO, 'SIMPLE') AS TIPO FROM ALMACEN3 WHERE ID_PRODUCTO = '" & txtIdProducto.Text & "'"
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        Tipo = tRs.Fields("TIPO")
                    End If
                End If
                C = Me.lvwOrdenCompra.ListItems.Count + 1
                Set ItMx = Me.lvwOrdenCompra.ListItems.Add(, , C)
                ItMx.SubItems(1) = Me.txtIdProducto.Text
                ItMx.SubItems(2) = Me.txtDescripcion.Text
                ItMx.SubItems(3) = Val(Me.txtCantidad_Pedido.Text) - Val(Me.txtCantidad.Text)
                ItMx.SubItems(4) = Me.txtID.Text
                ItMx.SubItems(5) = Me.txtPedido.Text
                ItMx.SubItems(6) = Me.txtCantidad_Pedido.Text
                ItMx.SubItems(7) = Origen
                ItMx.SubItems(8) = txtSucursal.Text
                ItMx.SubItems(9) = txtAgente.Text
                ItMx.SubItems(10) = txtFecha.Text
                ItMx.SubItems(11) = txtAlmacen.Text
                ItMx.SubItems(12) = Tipo
                Quitar = True
            End If
        End If
        If bLvw = 1 Then
             If Val(txtCantidad.Text) >= Val(lvwDirectas.ListItems.Item(Val(txtIndice.Text)).ListSubItems(5)) Or Quitar Then
                lvwDirectas.ListItems.Remove (Val(txtIndice.Text))
            Else
                lvwDirectas.ListItems.Item(Val(txtIndice.Text)).ListSubItems(5) = Val(lvwDirectas.ListItems.Item(Val(txtIndice.Text)).ListSubItems(5)) - Val(txtCantidad.Text)
            End If
            Origen = 1
        Else
            If bLvw = 2 Then
                If Val(txtCantidad.Text) >= Val(lvwIndirectas.ListItems.Item(Val(txtIndice.Text)).ListSubItems(5)) Or Quitar Then
                    lvwIndirectas.ListItems.Remove (Val(txtIndice.Text))
                Else
                    lvwIndirectas.ListItems.Item(Val(txtIndice.Text)).ListSubItems(5) = Val(lvwIndirectas.ListItems.Item(Val(txtIndice.Text)).ListSubItems(5)) - Val(txtCantidad.Text)
                End If
                Origen = 2
            End If
        End If
        Me.Agregar_Lista_Surtir (Origen)
        Me.cmdSurtir1.Enabled = False
        Me.cmdAgregar.Enabled = False
        Me.txtAgente.Text = ""
        Me.txtCantidad.Text = ""
        Me.txtDescripcion.Text = ""
        Me.txtFecha.Text = ""
        Me.txtIdProducto.Text = ""
        Me.txtPedido.Text = ""
        Me.txtPI.Text = ""
        Me.txtSucursal.Text = ""
        txtPD.Text = lvwDirectas.ListItems.Count
        txtPI.Text = lvwIndirectas.ListItems.Count
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Activate()
On Error GoTo ManejaError
    If Hay_Pedidos_Directos Then
        Llenar_Lista_Pedidos_Directos
        txtPD.Text = lvwDirectas.ListItems.Count
    Else
        Me.txtPD.Text = "0"
    End If
    If Hay_Pedidos_Indirectos Then
        Llenar_Lista_Pedidos_Indirectos
        txtPI.Text = lvwIndirectas.ListItems.Count
    Else
        Me.txtPI.Text = "0"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With lvwDirectas
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "SUCURSAL", 2000
        .ColumnHeaders.Add , , "AGENTE", 0
        .ColumnHeaders.Add , , "FECHA", 0
        .ColumnHeaders.Add , , "PRODUCTO", 2500
        .ColumnHeaders.Add , , "CANTIDAD", 500
        .ColumnHeaders.Add , , "Descripcion", 3000
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "", 0
        .ColumnHeaders.Add , , "ALMACEN", 0
        .ColumnHeaders.Add , , "COMENTARIO", 0
        .ColumnHeaders.Add , , "MARCA", 0
    End With
    With lvwIndirectas
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "SUCURSAL", 2000
        .ColumnHeaders.Add , , "AGENTE", 0
        .ColumnHeaders.Add , , "FECHA", 0
        .ColumnHeaders.Add , , "PRODUCTO", 2500
        .ColumnHeaders.Add , , "CANTIDAD", 500
        .ColumnHeaders.Add , , "Descripcion", 3000
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "", 0
        .ColumnHeaders.Add , , "ALMACEN", 0
        .ColumnHeaders.Add , , "COMENTARIO", 0
        .ColumnHeaders.Add , , "MARCA", 0
    End With
    With lvwOrdenCompra
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "NUMERO", 0
        .ColumnHeaders.Add , , "CLAVE", 2500
        .ColumnHeaders.Add , , "Descripcion", 2000
        .ColumnHeaders.Add , , "CANTIDAD", 500
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_PEDIDO", 0
        .ColumnHeaders.Add , , "CANTIDAD_PEDIDO", 0
        .ColumnHeaders.Add , , "ORIGEN", 0
        .ColumnHeaders.Add , , "SUCURSAL", 0
        .ColumnHeaders.Add , , "AGENTE", 0
        .ColumnHeaders.Add , , "FECHA", 0
        .ColumnHeaders.Add , , "ALMACEN", 1440
        .ColumnHeaders.Add , , "TIPO", 0
        .ColumnHeaders.Add , , "MARCA", 0
    End With
    With lvwSurtir
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "NUMERO", 0
        .ColumnHeaders.Add , , "CLAVE", 2500
        .ColumnHeaders.Add , , "Descripcion", 2000
        .ColumnHeaders.Add , , "CANTIDAD", 500
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ID_PEDIDO", 0
        .ColumnHeaders.Add , , "CANTIDAD_PEDIDO", 0
        .ColumnHeaders.Add , , "ORIGEN", 0
        .ColumnHeaders.Add , , "SUCURSAL", 0
        .ColumnHeaders.Add , , "AGENTE", 0
        .ColumnHeaders.Add , , "FECHA", 0
    End With
    Me.dtpFechaRequisicion.Value = Format(Date, "dd/mm/yyyy")
End Sub
Public Sub Llenar_Lista_Pedidos_Directos()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Me.lvwDirectas.ListItems.Clear
    sBuscar = "Select  P.Id_Pedido,P.Sucursal,P.Pidio,P.Fecha,DP.ID,DP.Id_Producto,DP.Cantidad,DP.Descripcion, DP.Almacen, P.COMENTARIO, DP.MARCA From Pedido AS P Join Detalle_Pedido AS DP ON DP.Id_Pedido = P.Id_Pedido Where P.Tipo= 'D' AND DP.Entregado='0'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            While Not .EOF
                Set ItMx = Me.lvwDirectas.ListItems.Add(, , .Fields("ID_PEDIDO"))
                If Not IsNull(.Fields("Sucursal")) Then ItMx.SubItems(1) = Trim(.Fields("Sucursal"))
                If Not IsNull(.Fields("Pidio")) Then ItMx.SubItems(2) = Trim(.Fields("Pidio"))
                If Not IsNull(.Fields("fecha")) Then ItMx.SubItems(3) = Trim(.Fields("fecha"))
                If Not IsNull(.Fields("ID_PRODUCTO")) Then ItMx.SubItems(4) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("CANTIDAD")) Then ItMx.SubItems(5) = Trim(.Fields("CANTIDAD"))
                If Not IsNull(.Fields("Descripcion")) Then ItMx.SubItems(6) = Trim(.Fields("Descripcion"))
                If Not IsNull(.Fields("ID")) Then ItMx.SubItems(7) = Trim(.Fields("ID"))
                If Not IsNull(.Fields("Almacen")) Then ItMx.SubItems(9) = Trim(.Fields("Almacen"))
                If Not IsNull(.Fields("COMENTARIO")) Then ItMx.SubItems(10) = .Fields("COMENTARIO")
                If Not IsNull(.Fields("MARCA")) Then ItMx.SubItems(11) = .Fields("MARCA")
                .MoveNext
            Wend
            .Close
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Public Sub Llenar_Lista_Pedidos_Indirectos()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Me.lvwIndirectas.ListItems.Clear
    sBuscar = "Select  P.Id_Pedido,P.Sucursal,P.Pidio,P.Fecha,DP.ID,DP.Id_Producto,DP.Cantidad,DP.Descripcion, DP.Almacen, P.COMENTARIO, DP.MARCA From Pedido AS P Join Detalle_Pedido AS DP ON DP.Id_Pedido = P.Id_Pedido Where P.Tipo= 'I' AND DP.Entregado='0'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            While Not .EOF
                Set ItMx = Me.lvwIndirectas.ListItems.Add(, , .Fields("ID_PEDIDO"))
                If Not IsNull(.Fields("Sucursal")) Then ItMx.SubItems(1) = Trim(.Fields("Sucursal"))
                If Not IsNull(.Fields("Pidio")) Then ItMx.SubItems(2) = Trim(.Fields("Pidio"))
                If Not IsNull(.Fields("fecha")) Then ItMx.SubItems(3) = Trim(.Fields("fecha"))
                If Not IsNull(.Fields("ID_PRODUCTO")) Then ItMx.SubItems(4) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("CANTIDAD")) Then ItMx.SubItems(5) = Trim(.Fields("CANTIDAD"))
                If Not IsNull(.Fields("Descripcion")) Then ItMx.SubItems(6) = Trim(.Fields("Descripcion"))
                If Not IsNull(.Fields("ID")) Then ItMx.SubItems(7) = Trim(.Fields("ID"))
                If Not IsNull(.Fields("Almacen")) Then ItMx.SubItems(9) = Trim(.Fields("Almacen"))
                If Not IsNull(.Fields("COMENTARIO")) Then ItMx.SubItems(10) = .Fields("COMENTARIO")
                If Not IsNull(.Fields("MARCA")) Then ItMx.SubItems(11) = .Fields("MARCA")
                .MoveNext
            Wend
            .Close
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image1_Click()
    If Me.SSTab1.Tab = 1 Or SSTab1.Tab = 0 Then
        FunImprime
    Else
        MsgBox "Seleccione en los Tabs entre pedidos Directos o Indirectos", vbExclamation, "SACC"
    End If
End Sub
Private Sub Image24_Click()
    frmRequisiciones.SSTab1.TabEnabled(1) = False
    frmRequisiciones.SSTab1.TabEnabled(2) = False
    frmRequisiciones.SSTab1.TabEnabled(3) = False
    frmRequisiciones.Command11.Enabled = False
    frmRequisiciones.Command12.Enabled = False
    frmRequisiciones.Frame7.Visible = False
    frmRequisiciones.Frame8.Visible = False
    frmRequisiciones.Frame4.Visible = False
    frmRequisiciones.Frame14.Visible = False
    frmRequisiciones.Show vbModal
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwDirectas_Click()
On Error GoTo ManejaError
    If lvwDirectas.ListItems.Count > 0 Then
        bLvw = 1
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        If Me.lvwDirectas.SelectedItem.SubItems(8) = "" Then
            Me.cmdAgregar.Enabled = True
            Me.cmdSurtir1.Enabled = True
        Else
            Me.cmdAgregar.Enabled = False
            Me.cmdSurtir1.Enabled = False
        End If
        Me.txtPedido.Text = Me.lvwDirectas.SelectedItem
        Me.txtSucursal.Text = Me.lvwDirectas.SelectedItem.SubItems(1)
        Me.txtAgente.Text = Me.lvwDirectas.SelectedItem.SubItems(2)
        Me.txtFecha.Text = Me.lvwDirectas.SelectedItem.SubItems(3)
        Me.txtIdProducto.Text = Me.lvwDirectas.SelectedItem.SubItems(4)
        Me.txtCantidad.Text = Me.lvwDirectas.SelectedItem.SubItems(5)
        Me.txtCantidad_Pedido.Text = Me.lvwDirectas.SelectedItem.SubItems(5)
        Me.txtDescripcion.Text = Me.lvwDirectas.SelectedItem.SubItems(6)
        Me.txtID.Text = Me.lvwDirectas.SelectedItem.SubItems(7)
        Me.txtAlmacen.Text = Me.lvwDirectas.SelectedItem.SubItems(9)
        Me.txtComentario.Text = Me.lvwDirectas.SelectedItem.SubItems(10)
        Me.txtMarca.Text = Me.lvwDirectas.SelectedItem.SubItems(11)
        txtIndice = lvwDirectas.SelectedItem.Index
        sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & lvwDirectas.SelectedItem.SubItems(4) & "' AND SUCURSAL = 'BODEGA'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Text1.Text = tRs.Fields("CANTIDAD")
        Else
            Text1.Text = "0"
        End If
        sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & lvwDirectas.SelectedItem.SubItems(4) & "' AND SUCURSAL = '" & lvwDirectas.SelectedItem.SubItems(1) & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Text11.Text = tRs.Fields("CANTIDAD")
        Else
            Text11.Text = "0"
        End If
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwIndirectas_Click()
On Error GoTo ManejaError
    bLvw = 2
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If Me.lvwIndirectas.SelectedItem.SubItems(8) = "" Then
        Me.cmdAgregar.Enabled = True
        Me.cmdSurtir1.Enabled = True
    Else
        Me.cmdAgregar.Enabled = False
        Me.cmdSurtir1.Enabled = False
    End If
    Me.txtPedido.Text = Me.lvwIndirectas.SelectedItem
    Me.txtSucursal.Text = Me.lvwIndirectas.SelectedItem.SubItems(1)
    Me.txtAgente.Text = Me.lvwIndirectas.SelectedItem.SubItems(2)
    Me.txtFecha.Text = Me.lvwIndirectas.SelectedItem.SubItems(3)
    Me.txtIdProducto.Text = Me.lvwIndirectas.SelectedItem.SubItems(4)
    Me.txtCantidad.Text = Me.lvwIndirectas.SelectedItem.SubItems(5)
    Me.txtCantidad_Pedido.Text = Me.lvwIndirectas.SelectedItem.SubItems(5)
    Me.txtDescripcion.Text = Me.lvwIndirectas.SelectedItem.SubItems(6)
    Me.txtID.Text = Me.lvwIndirectas.SelectedItem.SubItems(7)
    Me.txtAlmacen.Text = Me.lvwIndirectas.SelectedItem.SubItems(9)
    Me.txtComentario.Text = Me.lvwIndirectas.SelectedItem.SubItems(10)
    Me.txtMarca.Text = Me.lvwIndirectas.SelectedItem.SubItems(11)
    txtIndice = lvwIndirectas.SelectedItem.Index
    sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & lvwIndirectas.SelectedItem.SubItems(4) & "' AND SUCURSAL = 'BODEGA'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Text1.Text = tRs.Fields("CANTIDAD")
    Else
        Text1.Text = "0"
    End If
    sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & lvwIndirectas.SelectedItem.SubItems(4) & "' AND SUCURSAL = '" & lvwIndirectas.SelectedItem.SubItems(1) & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Text11.Text = tRs.Fields("CANTIDAD")
    Else
        Text11.Text = "0"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Function Puede_Agregar() As Boolean
On Error GoTo ManejaError
    If Me.txtDescripcion.Text = "" Then
        MsgBox "ESCRIBA LA Descripcion", vbInformation, "SACC"
        Me.txtDescripcion.SetFocus
        Puede_Agregar = False
        Exit Function
    End If
    
    If Me.txtCantidad.Text = "" Then
        MsgBox "ESCRIBA LA CANTIDAD", vbInformation, "SACC"
        Me.txtCantidad.SetFocus
        Puede_Agregar = False
        Exit Function
    End If
    Puede_Agregar = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Agregar_Lista_Ordenes(Origen As Integer)
On Error GoTo ManejaError
    Dim C As Integer
    Dim D As Integer
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim Tipo As String
    If txtAlmacen.Text = "A3" Then
        sBuscar = "SELECT isNull(TIPO, 'SIMPLE') AS TIPO FROM ALMACEN3 WHERE ID_PRODUCTO = '" & txtIdProducto.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Tipo = tRs.Fields("TIPO")
        End If
    Else
        Tipo = "SIMPLE"
    End If
    C = Me.lvwOrdenCompra.ListItems.Count + 1
    D = lvProd.ListItems.Count + 1
    Set ItMx = Me.lvwOrdenCompra.ListItems.Add(, , C)
    ItMx.SubItems(1) = Me.txtIdProducto.Text
    ItMx.SubItems(2) = Me.txtDescripcion.Text
    ItMx.SubItems(3) = Me.txtCantidad.Text
    ItMx.SubItems(4) = Me.txtID.Text
    ItMx.SubItems(5) = Me.txtPedido.Text
    ItMx.SubItems(6) = Me.txtCantidad_Pedido.Text
    ItMx.SubItems(7) = Origen
    ItMx.SubItems(8) = txtSucursal.Text
    ItMx.SubItems(9) = txtAgente.Text
    ItMx.SubItems(10) = txtFecha.Text
    ItMx.SubItems(11) = txtAlmacen.Text
    ItMx.SubItems(12) = Tipo
    ItMx.SubItems(13) = txtMarca.Text
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Agregar_Lista_Surtir(Origen As Integer)
On Error GoTo ManejaError
    Dim C As Integer
    C = Me.lvwSurtir.ListItems.Count + 1
    Set ItMx = Me.lvwSurtir.ListItems.Add(, , C)
    ItMx.SubItems(1) = Me.txtIdProducto.Text
    ItMx.SubItems(2) = txtDescripcion.Text
    ItMx.SubItems(3) = Me.txtCantidad.Text
    ItMx.SubItems(4) = Me.txtID.Text
    ItMx.SubItems(5) = Me.txtPedido.Text
    ItMx.SubItems(6) = Me.txtCantidad_Pedido.Text
    ItMx.SubItems(7) = Origen
    ItMx.SubItems(8) = txtSucursal.Text
    ItMx.SubItems(9) = txtAgente.Text
    ItMx.SubItems(10) = txtFecha.Text
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Function Puede_Surtir() As Boolean
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If Me.txtIdProducto.Text = "" Then
        MsgBox "ESCRIBA EL ID DEL PRODUCTO", vbInformation, "SACC"
        Me.txtIdProducto.SetFocus
        Puede_Surtir = False
        Exit Function
    End If
    If Me.txtCantidad.Text = "" Then
        MsgBox "ESCRIBA LA CANTIDAD", vbInformation, "SACC"
        Me.txtCantidad.SetFocus
        Puede_Surtir = False
        Exit Function
    End If
    nCantidad_Pedido = 0
    nExistencia = 0
    sBuscar = "SELECT ID_EXISTENCIA, ID_PRODUCTO, CANTIDAD From EXISTENCIAS WHERE ID_PRODUCTO= '" & Me.txtIdProducto.Text & "' AND SUCURSAL='BODEGA'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        If Not IsNull(tRs.Fields("CANTIDAD")) Then
            nCantidad_Pedido = tRs.Fields("CANTIDAD")
            nExistencia = tRs.Fields("ID_EXISTENCIA")
        End If
    End If
    If Me.txtCantidad.Text > nCantidad_Pedido Then
        MsgBox "NO HAY SUFICIENTE EXISTENCIA. LA MAXIMA CANTIDAD QUE PUEDE SURTIR ES " & Val(nCantidad_Pedido) & ".", vbInformation, "SACC"
        Puede_Surtir = False
        Exit Function
    End If
    Cantidad_Acumulada = 0
    NumReg = Me.lvwSurtir.ListItems.Count
    For Con = 1 To NumReg
        If Me.lvwSurtir.ListItems.Item(Con).SubItems(1) = Me.txtIdProducto.Text Then
            Cantidad_Acumulada = Cantidad_Acumulada + Me.lvwSurtir.ListItems.Item(Con).SubItems(3)
        End If
    Next Con
    If (Cantidad_Acumulada + Me.txtCantidad.Text) > nCantidad_Pedido Then
        MsgBox "NO HAY SUFICIENTE EXISTENCIA. LA MAXIMA CANTIDAD QUE PUEDE SURTIR ES " & Val(nCantidad_Pedido - Cantidad_Acumulada) & ".", vbInformation, "SACC"
        Puede_Surtir = False
        Exit Function
    End If
    Puede_Surtir = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub lvwOrdenCompra_ItemClick(ByVal Item As MSComctlLib.ListItem)
    bLOR = 1
    i = Item.Index
End Sub
Private Sub lvwSurtir_ItemClick(ByVal Item As MSComctlLib.ListItem)
    bLOR = 2
    i = Item.Index
End Sub
Private Sub TxtCantidad_Change()
    Me.txtCantidad.BackColor = &HFFE1E1
End Sub
Private Sub txtCantidad_GotFocus()
    txtCantidad.BackColor = &HFFE1E1
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys vbTab
End If
End Sub
Private Sub TxtCantidad_LostFocus()
    txtCantidad.BackColor = &H80000005
End Sub
Private Sub txtDescripcion_GotFocus()
    Me.txtDescripcion.BackColor = &HFFE1E1
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.cmdAgregar.Value = True
    End If
End Sub
Private Sub txtDescripcion_LostFocus()
    txtDescripcion.BackColor = &H80000005
End Sub
Private Sub txtIdProducto_GotFocus()
    Me.txtIdProducto.BackColor = &HFFE1E1
End Sub
Private Sub txtIdProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys vbTab
    End If
End Sub
Private Sub txtIdProducto_LostFocus()
    txtIdProducto.BackColor = &H80000005
End Sub
Private Sub FunImprime()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim ConPag As Integer
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    ConPag = 1
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate(App.Path & "\OrdenCompra.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
    ' Encabezado del reporte
    Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    oDoc.LoadImage Image2, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 70, 40, 43, 161, "Logo"
    sBuscar = "SELECT * FROM EMPRESA  "
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 40, 205, 100, 175, tRs.Fields("NOMBRE"), "F3", 8, hCenter
    oDoc.WTextBox 60, 224, 100, 175, tRs.Fields("DIRECCION"), "F3", 8, hLeft
    oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs.Fields("COLONIA"), "F3", 8, hLeft
    oDoc.WTextBox 70, 205, 100, 175, tRs.Fields("ESTADO") & "," & tRs.Fields("CD"), "F3", 8, hCenter
    oDoc.WTextBox 80, 205, 100, 175, tRs.Fields("TELEFONO"), "F3", 8, hCenter
    oDoc.WTextBox 80, 340, 20, 250, "Fecha :" & Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
    'CAJA1
    If SSTab1.Tab = 0 Then
        oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE PEDIDOS INDIRECTOS A BODEGA", "F3", 10, hCenter
    End If
    If SSTab1.Tab = 1 Then
        oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE PEDIDOS DIRECTOS A BODEGA", "F3", 10, hCenter
    End If
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 135
    oDoc.WLineTo 580, 135
    oDoc.LineStroke
    Posi = 135
    ' ENCABEZADO DEL DETALLE
    oDoc.WTextBox Posi, 5, 20, 80, "CLAVE", "F2", 8, hCenter
    oDoc.WTextBox Posi, 90, 20, 50, "CANTIDAD", "F2", 8, hCenter
    oDoc.WTextBox Posi, 144, 20, 70, "SUCURSAL", "F2", 8, hCenter
    oDoc.WTextBox Posi, 208, 20, 50, "FECHA", "F2", 8, hCenter
    oDoc.WTextBox Posi, 262, 20, 70, "AGENTE", "F2", 8, hCenter
    oDoc.WTextBox Posi, 324, 20, 120, "COMENTARIO", "F2", 8, hCenter
    Posi = Posi + 12
    ' Linea
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, Posi
    oDoc.WLineTo 580, Posi
    oDoc.LineStroke
    Posi = Posi + 6
    Cont = 1
    ' DETALLE
    If SSTab1.Tab = 1 Then
        Do While Cont <= lvwDirectas.ListItems.Count
            oDoc.WTextBox Posi, 5, 20, 80, lvwDirectas.ListItems(Cont).SubItems(4), "F3", 7, hLeft
            oDoc.WTextBox Posi, 90, 20, 50, lvwDirectas.ListItems(Cont).SubItems(5), "F3", 7, hCenter
            oDoc.WTextBox Posi, 144, 20, 70, lvwDirectas.ListItems(Cont).SubItems(1), "F3", 7, hCenter
            oDoc.WTextBox Posi, 208, 20, 50, Format(lvwDirectas.ListItems(Cont).SubItems(3), "dd/mm/yyyy"), "F3", 7, hCenter
            sBuscar = "SELECT (NOMBRE + ' ' + APELLIDOS) AS NOMBRE FROM USUARIOS WHERE ID_USUARIO = " & lvwDirectas.ListItems(Cont).SubItems(2)
            Set tRs1 = cnn.Execute(sBuscar)
            If Not (tRs1.EOF And tRs1.BOF) Then
                oDoc.WTextBox Posi, 262, 20, 70, tRs1.Fields("NOMBRE"), "F3", 7, hLeft
            Else
                oDoc.WTextBox Posi, 262, 20, 70, "SISTEMA", "F3", 7, hLeft
            End If
            oDoc.WTextBox Posi, 344, 20, 260, lvwDirectas.ListItems(Cont).SubItems(10), "F3", 7, hLeft
            Posi = Posi + 12
            If Posi >= 750 Then
                oDoc.NewPage A4_Vertical
                oDoc.WImage 70, 40, 43, 161, "Logo"
                oDoc.WTextBox 40, 205, 100, 175, tRs.Fields("NOMBRE"), "F3", 8, hCenter
                oDoc.WTextBox 60, 224, 100, 175, tRs.Fields("DIRECCION"), "F3", 8, hLeft
                oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs.Fields("COLONIA"), "F3", 8, hLeft
                oDoc.WTextBox 70, 205, 100, 175, tRs.Fields("ESTADO") & "," & tRs.Fields("CD"), "F3", 8, hCenter
                oDoc.WTextBox 80, 205, 100, 175, tRs.Fields("TELEFONO"), "F3", 8, hCenter
                oDoc.WTextBox 80, 340, 20, 250, "Fecha :" & Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
                'CAJA1
                oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE PEDIDOS INDIRECTOS A BODEGA", "F3", 10, hCenter
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, 135
                oDoc.WLineTo 580, 135
                oDoc.LineStroke
                Posi = 135
                ' ENCABEZADO DEL DETALLE
                oDoc.WTextBox Posi, 5, 20, 80, "CLAVE", "F2", 8, hCenter
                oDoc.WTextBox Posi, 90, 20, 50, "CANTIDAD", "F2", 8, hCenter
                oDoc.WTextBox Posi, 144, 20, 70, "SUCURSAL", "F2", 8, hCenter
                oDoc.WTextBox Posi, 208, 20, 50, "FECHA", "F2", 8, hCenter
                oDoc.WTextBox Posi, 262, 20, 70, "AGENTE", "F2", 8, hCenter
                oDoc.WTextBox Posi, 324, 20, 120, "COMENTARIO", "F2", 8, hCenter
                Posi = Posi + 12
                ' Linea
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, Posi
                oDoc.WLineTo 580, Posi
                oDoc.LineStroke
                Posi = Posi + 6
            End If
            Cont = Cont + 1
        Loop
    End If
    If SSTab1.Tab = 0 Then
        Do While Cont <= lvwIndirectas.ListItems.Count
            oDoc.WTextBox Posi, 5, 20, 80, lvwIndirectas.ListItems(Cont).SubItems(4), "F3", 7, hLeft
            oDoc.WTextBox Posi, 90, 20, 50, lvwIndirectas.ListItems(Cont).SubItems(5), "F3", 7, hCenter
            oDoc.WTextBox Posi, 144, 20, 70, lvwIndirectas.ListItems(Cont).SubItems(1), "F3", 7, hCenter
            oDoc.WTextBox Posi, 208, 20, 50, Format(lvwIndirectas.ListItems(Cont).SubItems(3), "dd/mm/yyyy"), "F3", 7, hCenter
            sBuscar = "SELECT (NOMBRE + ' ' + APELLIDOS) AS NOMBRE FROM USUARIOS WHERE ID_USUARIO = " & lvwDirectas.ListItems(Cont).SubItems(2)
            Set tRs1 = cnn.Execute(sBuscar)
            If Not (tRs1.EOF And tRs1.BOF) Then
                oDoc.WTextBox Posi, 262, 20, 70, tRs1.Fields("NOMBRE"), "F3", 7, hLeft
            Else
                oDoc.WTextBox Posi, 262, 20, 70, "SISTEMA", "F3", 7, hLeft
            End If
            oDoc.WTextBox Posi, 344, 20, 260, lvwIndirectas.ListItems(Cont).SubItems(10), "F3", 7, hLeft
            Posi = Posi + 12
            tRs3.MoveNext
            If Posi >= 750 Then
                oDoc.NewPage A4_Vertical
                oDoc.WImage 70, 40, 43, 161, "Logo"
                oDoc.WTextBox 40, 205, 100, 175, tRs.Fields("NOMBRE"), "F3", 8, hCenter
                oDoc.WTextBox 60, 224, 100, 175, tRs.Fields("DIRECCION"), "F3", 8, hLeft
                oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs.Fields("COLONIA"), "F3", 8, hLeft
                oDoc.WTextBox 70, 205, 100, 175, tRs.Fields("ESTADO") & "," & tRs.Fields("CD"), "F3", 8, hCenter
                oDoc.WTextBox 80, 205, 100, 175, tRs.Fields("TELEFONO"), "F3", 8, hCenter
                oDoc.WTextBox 80, 340, 20, 250, "Fecha :" & Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
                'CAJA1
                oDoc.WTextBox 115, 20, 100, 525, "REPORTE DE PEDIDOS INDIRECTOS A BODEGA", "F3", 10, hCenter
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, 135
                oDoc.WLineTo 580, 135
                oDoc.LineStroke
                Posi = 135
                ' ENCABEZADO DEL DETALLE
                oDoc.WTextBox Posi, 5, 20, 80, "CLAVE", "F2", 8, hCenter
                oDoc.WTextBox Posi, 90, 20, 50, "CANTIDAD", "F2", 8, hCenter
                oDoc.WTextBox Posi, 144, 20, 70, "SUCURSAL", "F2", 8, hCenter
                oDoc.WTextBox Posi, 208, 20, 50, "FECHA", "F2", 8, hCenter
                oDoc.WTextBox Posi, 262, 20, 70, "AGENTE", "F2", 8, hCenter
                oDoc.WTextBox Posi, 324, 20, 120, "COMENTARIO", "F2", 8, hCenter
                Posi = Posi + 12
                ' Linea
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, Posi
                oDoc.WLineTo 580, Posi
                oDoc.LineStroke
                Posi = Posi + 6
            End If
            Cont = Cont + 1
        Loop
    End If
    ' Linea
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 760
    oDoc.WLineTo 580, 760
    oDoc.WTextBox 780, 324, 20, 120, "Fin del Reporte", "F2", 8, hCenter
    oDoc.LineStroke
    Posi = Posi + 6
    oDoc.PDFClose
    oDoc.Show
End Sub
