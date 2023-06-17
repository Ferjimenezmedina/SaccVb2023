VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Ordenes 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Entrada"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   3960
      TabIndex        =   37
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483644
      TabCaption(0)   =   "Entradas"
      TabPicture(0)   =   "Ordenes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label19"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFolio"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListaSurtidosT"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListaEntradasDetalle"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ListaProdPedidos"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ListaEntradas"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ListaTemp"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "btnQuitar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame16"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtBuscaOrden"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "btnLimpiar"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtNoFactura"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "btnAceptarFactura"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text6"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text7"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text8"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Option3"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Option4"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Rechazados"
      TabPicture(1)   =   "Ordenes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListaRechazados"
      Tab(1).Control(1)=   "Text9"
      Tab(1).Control(2)=   "btnQuitarR"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton btnQuitarR 
         Caption         =   "Quitar"
         Enabled         =   0   'False
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
         Left            =   -68640
         Picture         =   "Ordenes.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   645
         Left            =   -74520
         MultiLine       =   -1  'True
         TabIndex        =   59
         Text            =   "Ordenes.frx":2A0A
         Top             =   5400
         Width           =   7095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Internacio."
         Height          =   255
         Left            =   6240
         TabIndex        =   57
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Nacional"
         Height          =   255
         Left            =   6240
         TabIndex        =   56
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   5160
         TabIndex        =   54
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   4680
         TabIndex        =   53
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   51
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2640
         TabIndex        =   48
         Top             =   6720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton btnAceptarFactura 
         Caption         =   "Aceptar"
         Enabled         =   0   'False
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
         Left            =   2760
         Picture         =   "Ordenes.frx":2A91
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtNoFactura 
         Height          =   285
         Left            =   240
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton btnLimpiar 
         Caption         =   "Limpiar"
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
         Left            =   6480
         Picture         =   "Ordenes.frx":5463
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtBuscaOrden 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Frame Frame16 
         Height          =   1575
         Left            =   240
         TabIndex        =   38
         ToolTipText     =   "En Proceso"
         Top             =   3240
         Width           =   7335
         Begin VB.CommandButton btnRechazar 
            Caption         =   "Rechazar"
            Enabled         =   0   'False
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
            Left            =   6120
            Picture         =   "Ordenes.frx":7E35
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtCodBarras 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   1200
            Width           =   7095
         End
         Begin VB.Frame Frame4 
            Caption         =   "Codigo de Barras"
            Height          =   975
            Left            =   3120
            TabIndex        =   40
            Top             =   120
            Width           =   1575
            Begin VB.OptionButton OpNS 
               Caption         =   "N. Serie"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   600
               Width           =   1095
            End
            Begin VB.OptionButton OpNP 
               Caption         =   "N. Parte"
               Height          =   195
               Left            =   120
               TabIndex        =   11
               Top             =   280
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.TextBox txtPrecio 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1320
            TabIndex        =   10
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton btnAgregar 
            Caption         =   "Agregar"
            Enabled         =   0   'False
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
            Left            =   4800
            Picture         =   "Ordenes.frx":A807
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtCant 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblCod 
            Caption         =   "Codigo de Barras"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Precio Individual"
            Height          =   255
            Left            =   1320
            TabIndex        =   42
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CommandButton btnQuitar 
         Caption         =   "Quitar"
         Enabled         =   0   'False
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
         Left            =   6480
         Picture         =   "Ordenes.frx":D1D9
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6720
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListaTemp 
         Height          =   135
         Left            =   960
         TabIndex        =   17
         Top             =   6720
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   238
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListaEntradas 
         Height          =   1695
         Left            =   240
         TabIndex        =   14
         Top             =   4920
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2990
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListaProdPedidos 
         Height          =   1575
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2778
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListaEntradasDetalle 
         Height          =   135
         Left            =   240
         TabIndex        =   16
         Top             =   6720
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   238
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListaSurtidosT 
         Height          =   135
         Left            =   1800
         TabIndex        =   18
         Top             =   6720
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   238
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListaRechazados 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   58
         Top             =   1200
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5318
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label9 
         Caption         =   "No. Envio"
         Height          =   255
         Left            =   1440
         TabIndex        =   52
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Numero de Orden De Compra"
         Height          =   15
         Left            =   120
         TabIndex        =   50
         Top             =   6120
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblFolio 
         Height          =   255
         Left            =   720
         TabIndex        =   47
         Top             =   6120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Num. Factura"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Orden(es) de Compra"
         Height          =   255
         Left            =   3960
         TabIndex        =   45
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1800
      TabIndex        =   35
      Top             =   6240
      Width           =   975
      Begin VB.Label Label1 
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
         TabIndex        =   36
         Top             =   960
         Width           =   975
      End
      Begin VB.Image btnGuardar 
         Height          =   720
         Left            =   120
         MouseIcon       =   "Ordenes.frx":FBAB
         MousePointer    =   99  'Custom
         Picture         =   "Ordenes.frx":FEB5
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2880
      TabIndex        =   33
      Top             =   6240
      Width           =   975
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
         TabIndex        =   34
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Ordenes.frx":11877
         MousePointer    =   99  'Custom
         Picture         =   "Ordenes.frx":11B81
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   30
      Top             =   3120
      Width           =   3735
      Begin VB.TextBox txtNoTemp 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ListView ListaDocumentos 
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   2566
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clave"
      Height          =   195
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre"
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   3735
      Begin VB.TextBox txtID_User 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
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
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      CausesValidation=   0   'False
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0.00"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      CausesValidation=   0   'False
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0.00"
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      CausesValidation=   0   'False
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0.00"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Documentos Pendientes"
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
      TabIndex        =   31
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label lblProv 
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
      TabIndex        =   27
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
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
      Left            =   1200
      TabIndex        =   26
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
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
      Left            =   1200
      TabIndex        =   25
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SUBTOTAL"
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
      Left            =   840
      TabIndex        =   24
      Top             =   5160
      Width           =   1095
   End
End
Attribute VB_Name = "Ordenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim xA As ListItem
Dim Prods As ListItem
Dim ClvProd As String
Dim DesProd As String
Dim PreProd As String
Dim CLVCLIEN As String
Dim NomClien As String
Dim DesClien As String
Dim DelInd As String
Dim DelDes As String
Dim DelCan As String
Dim DelPre As String
Dim BanCnn As Boolean
Dim sqlQuery As String
Dim DOMI As String
Dim GTIA As String
Dim IdClien As String
Dim NoAsTec As String
Dim TelCasa As String
Dim TelTrabajo As String
Dim Direc As String
Dim NoExte As String
Dim NoInte As String
Dim COLONIA As String
Dim ENTR As Integer
Dim IdProducto As String
Dim ElimProd As Integer
Dim VarDescrip As String
Dim VarPrecio As String
Dim sTipoO As String
Dim VarCant As String
Dim VarId As String
Dim VarSurtido As String
Dim VarIndex As Integer
Dim VarEliminado As Integer
Dim VarNumOrden As Integer
Dim VarIdOrdenCompra As String
Dim sTipo As String
Private Sub btnAceptarFactura_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    If txtBuscaOrden.Text <> "" Then
        Buscar
    Else
        MsgBox "Seleccione Primero los documentos"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub btnAgregar_Click()
On Error GoTo ManejaError
    Dim Selected As Integer
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim ID As String
    Dim Cont As Integer
    Dim PreFinVen As String
    Selected = ListaProdPedidos.SelectedItem.Index
    If Int(txtCant.Text) > Int(ListaProdPedidos.ListItems.Item(Selected).SubItems(3)) Then
        MsgBox "No puede recibir mas productos de los pedidos," & Chr(13) & "                Favor de Verificar"
    Else
        Set tLi = ListaEntradas.ListItems.Add(, , ListaProdPedidos.ListItems.Item(Selected) & "")
        tLi.SubItems(1) = ListaProdPedidos.ListItems.Item(Selected).SubItems(1)
        tLi.SubItems(2) = txtPrecio.Text
        tLi.SubItems(3) = txtCant.Text
        tLi.SubItems(4) = ListaProdPedidos.ListItems.Item(Selected).SubItems(4)
        tLi.SubItems(5) = ListaProdPedidos.ListItems.Item(Selected).SubItems(5)
        PreFinVen = Format(Val(Replace(txtPrecio.Text, ",", "")) * Val(Replace(txtCant.Text, ",", "")), "###,###,##0.00")
        Text3.Text = Format(Val(Replace(Text3.Text, ",", "")) + Val(Replace(PreFinVen, ",", "")), "###,###,##0.00")
        If sTipo <> "I" Then
            Text4.Text = Format(Val(Replace(Text3.Text, ",", "")) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        End If
        Text5.Text = Format(Val(Replace(Text4.Text, ",", "")) + Val(Replace(Text3.Text, ",", "")), "###,###,##0.00")
        If OpNP.Value Then
            Set Prods = ListaTemp.ListItems.Add(, , ListaProdPedidos.ListItems.Item(Selected) & "")
            Prods.SubItems(1) = ListaProdPedidos.ListItems.Item(Selected).SubItems(1)
            Prods.SubItems(2) = txtPrecio.Text
            Prods.SubItems(3) = txtCant.Text
            Prods.SubItems(4) = ListaProdPedidos.ListItems.Item(Selected).SubItems(4)
        Else
            If OpNS.Value Then
                For Cont = 1 To Int(txtCant.Text)
                    Set Prods = ListaTemp.ListItems.Add(, , ListaProdPedidos.ListItems.Item(Selected) & "")
                    Prods.SubItems(1) = ListaProdPedidos.ListItems.Item(Selected).SubItems(1)
                    Prods.SubItems(2) = txtPrecio.Text
                    Prods.SubItems(3) = "1"
                    Prods.SubItems(4) = ListaProdPedidos.ListItems.Item(Selected).SubItems(4)
                Next Cont
            End If
        End If
        Set tLi = ListaSurtidosT.ListItems.Add(, , ListaProdPedidos.ListItems.Item(Selected) & "")
        tLi.SubItems(1) = ListaProdPedidos.ListItems.Item(Selected).SubItems(1)
        tLi.SubItems(2) = ListaProdPedidos.ListItems.Item(Selected).SubItems(2)
        tLi.SubItems(3) = ListaProdPedidos.ListItems.Item(Selected).SubItems(3)
        tLi.SubItems(4) = ListaProdPedidos.ListItems.Item(Selected).SubItems(4)
        tLi.SubItems(5) = ListaProdPedidos.ListItems.Item(Selected).SubItems(5)
        ListaProdPedidos.ListItems.Remove (Selected)
        txtNoFactura.Enabled = False
        txtBuscaOrden.Enabled = False
        btnAceptarFactura.Enabled = False
        Text1.Enabled = False
        ListView1.Enabled = False
        ListaDocumentos.Enabled = False
        Image9.Enabled = False
        btnGuardar.Enabled = False
        btnQuitar.Enabled = False
        txtCant.Enabled = False
        txtPrecio.Enabled = False
        btnAgregar.Enabled = False
        btnRechazar.Enabled = False
        btnLimpiar.Enabled = False
        ListaProdPedidos.Enabled = False
        ListaEntradas.Enabled = False
        txtCodBarras.Enabled = True
        txtCodBarras.SetFocus
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Buscar()
On Error GoTo ManejaError
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBus As String
    Dim SUC As String
    sBus = "SELECT ID, ID_PRODUCTO, DESCRIPCION, PRECIO, CANTIDADP, SURTIDO, CONFIRMADA, TIPO, ID_ORDEN_COMPRA FROM VSORDENES WHERE ID_ORDEN_COMPRA IN (" & txtBuscaOrden.Text & ") AND ( CANTIDADP > 0 AND ID_PROVEEDOR = '" & txtID_Prov & "') GROUP BY ID, ID_PRODUCTO, DESCRIPCION, PRECIO, CANTIDADP, SURTIDO, CONFIRMADA, TIPO, ID_ORDEN_COMPRA"
    Set tRs = cnn.Execute(sBus)
    With tRs
        ListaProdPedidos.ListItems.Clear
        Do While Not .EOF
            Set tLi = ListaProdPedidos.ListItems.Add(, , .Fields("ID_PRODUCTO"))
            If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion")
            If Not IsNull(.Fields("PRECIO")) Then tLi.SubItems(2) = Format(CDbl(.Fields("PRECIO")), "###,###,##0.00")
            If Not IsNull(.Fields("CANTIDADP")) Then tLi.SubItems(3) = .Fields("CANTIDADP")
            tLi.SubItems(4) = .Fields("ID") & ""
            If Not IsNull(.Fields("SURTIDO")) Then
                tLi.SubItems(5) = .Fields("SURTIDO")
            Else
                tLi.SubItems(5) = "0"
            End If
            If Not IsNull(.Fields("ID_ORDEN_COMPRA")) Then tLi.SubItems(6) = .Fields("ID_ORDEN_COMPRA")
            'cv
            btnRechazar.Enabled = False
            If Not IsNull(.Fields("CONFIRMADA")) Then
               If .Fields("CONFIRMADA") = "Y" Then
                  btnRechazar.Enabled = True
               End If
            End If
            sTipo = .Fields("TIPO")
            .MoveNext
        Loop
    End With
    Me.ListaProdPedidos.SetFocus
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub btnLimpiar_Click()
    txtBuscaOrden.Text = ""
End Sub
Private Sub btnQuitar_Click()
On Error GoTo ManejaError
    Dim ID As String
    Dim Selected As Integer
    Dim Cont As Integer
    Dim Precio As String
    Dim cant As String
    Dim tLi As ListItem
    Selected = ListaEntradas.SelectedItem.Index
    ID = ListaEntradas.ListItems.Item(Selected).SubItems(4)
    Precio = ListaEntradas.ListItems.Item(Selected).SubItems(2)
    cant = ListaEntradas.ListItems.Item(Selected).SubItems(3)
    ListaEntradas.ListItems.Remove (Selected)
    Precio = Format(CDbl(Precio) * CDbl(cant), "###,###,##0.00")
    Text3.Text = Format(CDbl(Text3.Text) - CDbl(Precio), "###,###,##0.00")
    Text4.Text = Format(CDbl(Text3.Text) * CDbl("0,15"), "###,###,##0.00")
    Text5.Text = Format(CDbl(Text4.Text) + CDbl(Text3.Text), "###,###,##0.00")
    ListaEntradasDetalle.ListItems.Remove (ElimProd)
    Cont = 1
    Do While Cont <= ListaSurtidosT.ListItems.Count
        If ListaSurtidosT.ListItems.Item(Cont).SubItems(4) = ID Then
            Set tLi = ListaProdPedidos.ListItems.Add(, , ListaSurtidosT.ListItems.Item(Cont) & "")
                tLi.SubItems(1) = ListaSurtidosT.ListItems.Item(Cont).SubItems(1)
                tLi.SubItems(2) = ListaSurtidosT.ListItems.Item(Cont).SubItems(2)
                tLi.SubItems(3) = ListaSurtidosT.ListItems.Item(Cont).SubItems(3)
                tLi.SubItems(4) = ListaSurtidosT.ListItems.Item(Cont).SubItems(4)
                tLi.SubItems(5) = ListaSurtidosT.ListItems.Item(Cont).SubItems(5)
            Cont = ListaSurtidosT.ListItems.Count + 1
        End If
        Cont = Cont + 1
    Loop
    If ListaEntradas.ListItems.Count = 0 Then
        ListaEntradas.ListItems.Clear
        ListaEntradasDetalle.ListItems.Clear
        Text1.Enabled = True
        ListView1.Enabled = True
        ListaDocumentos.Enabled = True
        Image9.Enabled = True
        btnGuardar.Enabled = False
        btnQuitar.Enabled = False
        txtNoFactura.Enabled = True
        txtBuscaOrden.Enabled = False
        btnAceptarFactura.Enabled = False
        btnLimpiar.Enabled = True
        txtNoFactura.Text = ""
        txtBuscaOrden.Text = ""
        Text3.Text = "0.00"
        Text4.Text = "0.00"
        Text5.Text = "0.00"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub btnGuardar_Click()
On Error GoTo ManejaError
    Dim PRODU As String
    Dim SUC As String
    Dim Cont As Integer
    Dim Surtido As String
    Dim ID As String
    Dim sBuscar As String
    Dim sBusca As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim X As Integer
    Dim orde As Integer
    Dim precosto As Double
    SUC = VarMen.Text4(0).Text
    For Cont = 1 To ListaEntradas.ListItems.Count
        PRODU = ListaEntradas.ListItems.Item(Cont)
        precosto = ListaEntradas.ListItems.Item(Cont).SubItems(2)
        ID = ListaEntradas.ListItems.Item(Cont).SubItems(4)
        Surtido = ListaEntradas.ListItems.Item(Cont).SubItems(5)
        sBuscar = "select cantidad From Existencias where Sucursal = '" & SUC & " ' and ID_Producto = '" & PRODU & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD + " & CDbl(ListaEntradas.ListItems.Item(Cont).SubItems(3)) & " WHERE SUCURSAL = '" & SUC & " ' AND ID_PRODUCTO = '" & PRODU & "'"
            cnn.Execute (sBuscar)
        Else
            sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & PRODU & "', " & ListaEntradas.ListItems.Item(Cont).SubItems(3) & ", '" & SUC & "' );"
            cnn.Execute (sBuscar)
        End If
        sBuscar = "UPDATE ORDEN_COMPRA_DETALLE SET SURTIDO = " & CDbl(ListaEntradas.ListItems.Item(Cont).SubItems(3)) + CDbl(Surtido) & ", FACT_PROVE = '" & txtNoFactura.Text & "', NO_ENVIO = '" & Text6.Text & "' WHERE ID = " & ID
        cnn.Execute (sBuscar)
        sBuscar = "SELECT PRECIO_COSTO FROM ALMACEN3 WHERE ID_PRODUCTO = '" & PRODU & "'"
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            If tRs3.Fields("PRECIO_COSTO") < precosto Then
                sBuscar = "UPDATE ALMACEN3 SET PRECIO_COSTO = '" & precosto & "' WHERE  ID_PRODUCTO = '" & PRODU & "'"
                cnn.Execute (sBuscar)
            End If
        Else
            sBuscar = "SELECT PRECIO_COSTO FROM ALMACEN2 WHERE ID_PRODUCTO = '" & PRODU & "'"
            Set tRs3 = cnn.Execute(sBuscar)
            If Not (tRs3.EOF And tRs3.BOF) Then
                If tRs3.Fields("PRECIO_COSTO") < precosto Then
                    sBuscar = "UPDATE ALMACEN2 SET PRECIO_COSTO = '" & precosto & "' WHERE  ID_PRODUCTO = '" & PRODU & "'"
                    cnn.Execute (sBuscar)
                End If
            Else
                sBuscar = "SELECT PRECIO_COSTO FROM ALMACEN1 WHERE ID_PRODUCTO = '" & PRODU & "'"
                Set tRs3 = cnn.Execute(sBuscar)
                If Not (tRs3.EOF And tRs3.BOF) Then
                    If tRs3.Fields("PRECIO_COSTO") < precosto Then
                        sBuscar = "UPDATE ALMACEN1 SET PRECIO_COSTO = '" & precosto & "' WHERE  ID_PRODUCTO = '" & PRODU & "'"
                        cnn.Execute (sBuscar)
                    End If
                End If
            End If
        End If
    Next Cont
    sBuscar = "INSERT INTO ENTRADAS (ID_PROVEEDOR, FECHA, TOTAL, FACTURA, ID_USUARIO, NUM_ORDEN, ID_ORDEN_COMPRA, TIPO_ORDEN) VALUES ('" & txtID_Prov.Text & "', '" & Format(Date, "dd/mm/yyyy") & "', " & Replace(Text5.Text, ",", "") & ", '" & txtNoFactura.Text & "', '" & VarMen.Text1(0).Text & "'," & Text7.Text & ", " & VarIdOrdenCompra & ", '" & sTipoO & "');"
    Set tRs = cnn.Execute(sBuscar)
    If Option3.Value = True Then
        sBuscar = "UPDATE ORDEN_COMPRA SET FACT_PROVE = '" & txtNoFactura.Text & "', TOT_PRO = " & Replace(Text5.Text, ",", "") & " WHERE ID_ORDEN_COMPRA='" & Text8.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        'ALMACEN3 WHERE ID_PRODUCTO
        sBuscar = "UPDATE ALMACEN3 SET PRECIO_COSTO = '" & precosto & "' WHERE  ID_PRODUCTO = '" & PRODU & "'"
        cnn.Execute (sBuscar)
    End If
    If Option4.Value = True Then
        sBuscar = "UPDATE ORDEN_COMPRA SET FACT_PROVE = '" & txtNoFactura.Text & "', TOT_PRO = " & Replace(Text5.Text, ",", "") & "  WHERE ID_ORDEN_COMPRA='" & Text8.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
    End If
    sBuscar = "SELECT MAX(ID_ENTRADA) AS ID_ENTRADA FROM ENTRADAS"
    Set tRs = cnn.Execute(sBuscar)
    lblFolio.Caption = tRs.Fields("ID_ENTRADA")
    ENTR = tRs.Fields("ID_ENTRADA")
    For Cont = 1 To ListaEntradasDetalle.ListItems.Count
        sBuscar = "INSERT INTO ENTRADA_PRODUCTO (ID_ENTRADA, ID_PRODUCTO, CANTIDAD, PRECIO, MONEDA, FECHA, ID_SUCURSAL, CODIGO_BARAS) VALUES"
        sBuscar = sBuscar & "('" & tRs.Fields("ID_ENTRADA") & "', '" & ListaEntradasDetalle.ListItems.Item(Cont)
        sBuscar = sBuscar & "', '" & ListaEntradasDetalle.ListItems.Item(Cont).SubItems(3)
        sBuscar = sBuscar & "', " & ListaEntradasDetalle.ListItems.Item(Cont).SubItems(2)
        sBuscar = sBuscar & ", '', '" & Format(Date, "dd/mm/yyyy")
        sBuscar = sBuscar & "', '" & VarMen.Text4(6).Text
        sBuscar = sBuscar & "', '" & ListaEntradasDetalle.ListItems.Item(Cont).SubItems(5) & "' );"
        Set tRs2 = cnn.Execute(sBuscar)
        sBuscar = "UPDATE ENTRADA_PRODUCTO SET FACT_PROV= '" & txtNoFactura & "', NUM_ORDEN= '" & Text7.Text & "' WHERE ID_ENTRADA = '" & ENTR & "'"
        cnn.Execute (sBuscar)
    Next Cont
    'cv
    If Me.ListaRechazados.ListItems.Count > 0 Then
        For X = 1 To ListaRechazados.ListItems.Count
            sBuscar = "INSERT INTO ORDENES_NO_SURTIDAS (ID_PRODUCTO, DESCRIPCION, PRECIO, CANTIDAD, NUM_ORDEN, TIPO, ESTADO) " & _
                      "                         VALUES ('" & ListaRechazados.ListItems(X) & "', '" & Replace(Me.ListaRechazados.ListItems(X).SubItems(1), Chr(34), "") & "', '" & Me.ListaRechazados.ListItems(X).SubItems(2) & "', '" & Me.ListaRechazados.ListItems(X).SubItems(3) & "', '" & VarNumOrden & "', '" & sTipoO & "', 'I')"
            cnn.Execute (sBuscar)
            sBuscar = "UPDATE ORDEN_COMPRA_DETALLE SET SURTIDO = SURTIDO + " & ListaRechazados.ListItems(X).SubItems(3) & " WHERE ID = " & ListaRechazados.ListItems(X).SubItems(4)
            cnn.Execute (sBuscar)
        Next X
    End If
    Imprimir
    ListaEntradas.ListItems.Clear
    ListaEntradasDetalle.ListItems.Clear
    ListaProdPedidos.ListItems.Clear
    ListView1.ListItems.Clear
    ListaDocumentos.ListItems.Clear
    Text1.Enabled = True
    ListView1.Enabled = True
    ListaDocumentos.Enabled = True
    Image9.Enabled = True
    btnGuardar.Enabled = False
    btnQuitar.Enabled = False
    txtNoFactura.Enabled = True
    txtBuscaOrden.Enabled = False
    btnAceptarFactura.Enabled = False
    txtNoFactura.Text = ""
    txtBuscaOrden.Text = ""
    Text3.Text = "0.00"
    Text4.Text = "0.00"
    Text5.Text = "0.00"
    Text1.Text = ""
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub btnQuitarR_Click()
    If VarEliminado > 0 Then Me.ListaRechazados.ListItems.Remove (VarEliminado)
    VarEliminado = 0
End Sub
Private Sub btnRechazar_Click()
On Error GoTo ManejaError
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBus As String
    Dim SUC As String
    Set tLi = ListaRechazados.ListItems.Add(, , IdProducto)
    tLi.SubItems(1) = VarDescrip
    tLi.SubItems(2) = VarPrecio
    tLi.SubItems(3) = VarCant
    tLi.SubItems(4) = VarId
    tLi.SubItems(5) = VarSurtido
    ListaProdPedidos.ListItems.Remove (VarIndex)
    Me.ListaRechazados.SetFocus
    VarDescrip = ""
    VarPrecio = ""
    VarCant = ""
    VarId = ""
    VarSurtido = ""
    VarIndex = 0
    btnQuitarR.Enabled = True
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "# DEL PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE", 6100
    End With
    With ListaProdPedidos
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLV DEL PRODUCTO", 1600
        .ColumnHeaders.Add , , "Descripcion", 6800
        .ColumnHeaders.Add , , "PRECIO", 1500
        .ColumnHeaders.Add , , "CANTIDAD", 1500
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "SURTIDO", 0
        .ColumnHeaders.Add , , "ID_ORDEN_COMPRA", 0
    End With
    With ListaSurtidosT
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLV DEL PRODUCTO", 1600
        .ColumnHeaders.Add , , "Descripcion", 6800
        .ColumnHeaders.Add , , "PRECIO", 1500
        .ColumnHeaders.Add , , "CANTIDAD", 1500
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "SURTIDO", 0
    End With
    With ListaEntradas
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLV DEL PRODUCTO", 1600
        .ColumnHeaders.Add , , "Descripcion", 6800
        .ColumnHeaders.Add , , "PRECIO", 2000
        .ColumnHeaders.Add , , "CANTIDAD", 2000
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "SURTIDO", 0
    End With
    With ListaEntradasDetalle
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLV DEL PRODUCTO", 1600
        .ColumnHeaders.Add , , "Descripcion", 6800
        .ColumnHeaders.Add , , "PRECIO", 2000
        .ColumnHeaders.Add , , "CANTIDAD", 2000
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "C.BARRAS", 0
    End With
    With ListaTemp
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLV DEL PRODUCTO", 1600
        .ColumnHeaders.Add , , "Descripcion", 6800
        .ColumnHeaders.Add , , "PRECIO", 2000
        .ColumnHeaders.Add , , "CANTIDAD", 2000
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "C.BARRAS", 0
    End With
    With ListaDocumentos
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "# ID", 0
        .ColumnHeaders.Add , , "# ORDEN", 950
        .ColumnHeaders.Add , , "TOTAL", 1400
        .ColumnHeaders.Add , , "FECHA", 1150
    End With
    'cv
    With ListaRechazados
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLV DEL PRODUCTO", 1600
        .ColumnHeaders.Add , , "Descripcion", 6800
        .ColumnHeaders.Add , , "PRECIO", 1500
        .ColumnHeaders.Add , , "CANTIDAD", 1500
        .ColumnHeaders.Add , , "ID", 1500
        .ColumnHeaders.Add , , "TIPO", 1500
    End With
    CLVCLIEN = ""
    txtID_User.Text = VarMen.Text1(0).Text
    txtID_User.Enabled = False
    txtID_Prov.Enabled = False
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    txtCant.Enabled = False
    txtPrecio.Enabled = False
    txtNoFactura.Enabled = False
    txtCodBarras.Enabled = False
    btnLimpiar.Enabled = True
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListaDocumentos_DblClick()
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBus As String
    Dim SUC As String
    Dim Docum As String
    Text7.Text = ""
    Text8.Text = ""
    If (txtBuscaOrden.Text <> "") Then
        Dim Aux As String
        Aux = txtBuscaOrden.Text
        If Strings.Right(Aux, 1) = "," Then
            txtBuscaOrden.Text = txtBuscaOrden.Text & ListaDocumentos.ListItems.Item(ListaDocumentos.SelectedItem.Index)
        Else
            txtBuscaOrden.Text = txtBuscaOrden.Text & ", " & ListaDocumentos.ListItems.Item(ListaDocumentos.SelectedItem.Index)
        End If
    Else
        If ListaDocumentos.ListItems.Count > 0 Then
            txtBuscaOrden.Text = ListaDocumentos.ListItems.Item(ListaDocumentos.SelectedItem.Index)
        End If
    End If
    Text7.Text = ListaDocumentos.SelectedItem.SubItems(1)
    Text8.Text = ListaDocumentos.ListItems.Item(ListaDocumentos.SelectedItem.Index)
    sBus = "SELECT * FROM ORDEN_COMPRA WHERE ID_ORDEN_COMPRA='" & Text8.Text & "'"
    Set tRs = cnn.Execute(sBus)
    If tRs.Fields("TIPO") = "I" Then
        Option4.Value = True
        sTipoO = "I"
    End If
    If tRs.Fields("TIPO") = "N" Then
        Option3.Value = True
        sTipoO = "N"
    End If
End Sub
Private Sub ListaDocumentos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtNoTemp.Text = Item
    VarNumOrden = Item.SubItems(1)
End Sub
Private Sub ListaEntradas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ElimProd = Item.Index
End Sub
Private Sub ListaProdPedidos_DblClick()
    txtCant.Enabled = True
    txtPrecio.Enabled = True
    txtCant.SetFocus
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT PRECIO FROM ORDEN_COMPRA_DETALLE WHERE ID_PRODUCTO = '" & IdProducto & "' AND ID_ORDEN_COMPRA IN (" & txtBuscaOrden.Text & ")"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        txtPrecio.Text = tRs.Fields("PRECIO")
    End If
End Sub
Private Sub ListaProdPedidos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.txtCant.Text = Item.SubItems(3)
    txtCant.Enabled = False
    txtPrecio.Enabled = False
    IdProducto = Item
    VarDescrip = Item.SubItems(1)
    VarPrecio = Item.SubItems(2)
    VarCant = Item.SubItems(3)
    VarId = Item.SubItems(4)
    VarSurtido = Item.SubItems(5)
    VarIdOrdenCompra = Item.SubItems(6)
    VarIndex = Item.Index
End Sub
Private Sub ListaRechazados_ItemClick(ByVal Item As MSComctlLib.ListItem)
    VarEliminado = Item.Index
End Sub
Private Sub ListView1_DblClick()
On Error GoTo ManejaError
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBus As String
    Dim SUC As String
    Dim Docum As String
    If ListaEntradas.ListItems.Count = 0 Then
        txtBuscaOrden.Text = ""
        txtNoFactura.Text = ""
        Docum = "VSORDENESP"
        sBus = "SELECT * FROM " & Docum & " WHERE ID_PROVEEDOR = '" & txtID_Prov & "' ORDER BY NUM_ORDEN"
        Set tRs = cnn.Execute(sBus)
        ListaDocumentos.ListItems.Clear
        ListaProdPedidos.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            With tRs
                Do While Not .EOF
                    Set tLi = ListaDocumentos.ListItems.Add(, , .Fields("ID_ORDEN_COMPRA") & "")
                    tLi.SubItems(1) = .Fields("NUM_ORDEN") & ""
                    If Not IsNull(.Fields("TOTAL")) Then tLi.SubItems(2) = .Fields("TOTAL") & ""
                    tLi.SubItems(3) = .Fields("FECHA") & ""
                    .MoveNext
                Loop
            End With
        End If
    Else
        MsgBox "La Lista de productos recibidos debe estar vacía", , "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtID_Prov.Text = Item
    Text1.Text = Item.SubItems(1)
    txtNoFactura.Enabled = True
End Sub
Private Sub Text1_Change()
    If Text1.Text = "" Then
        txtNoFactura.Enabled = False
    Else
        txtNoFactura.Enabled = True
    End If
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And Text1.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        Dim sBus As String
        Dim CadClien As String
        If Option2.Value = True Then
            CadClien = Text1.Text
            CadClien = Replace(CadClien, " ", "%")
            sBus = "SELECT ID_PROVEEDOR, NOMBRE FROM PROVEEDOR WHERE NOMBRE LIKE '%" & CadClien & "%'"
        Else
            sBus = "SELECT ID_PROVEEDOR, NOMBRE FROM PROVEEDOR WHERE ID_PROVEEDOR = " & Text1.Text
        End If
        Set tRs = cnn.Execute(sBus)
        With tRs
            ListView1.ListItems.Clear
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PROVEEDOR") & "")
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & ""
                .MoveNext
            Loop
            Me.ListView1.SetFocus
        End With
    End If
    Dim Valido As String
    If Option1.Value = True Then
        Valido = "1234567890"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    Else
        Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    ElseIf (KeyAscii = 13) And (txtNoFactura.Text <> "") Then
            btnAceptarFactura.Value = True
    End If
End Sub
Private Sub txtBuscaOrden_GotFocus()
    Me.txtBuscaOrden.BackColor = &HFFE1E1
End Sub
Private Sub txtBuscaOrden_LostFocus()
    txtBuscaOrden.BackColor = &H80000005
End Sub
Private Sub txtCant_Change()
    If (txtCant.Text <> "") And (txtPrecio.Text <> "") Then
        btnAgregar.Enabled = True
        btnRechazar.Enabled = True
    Else
        btnAgregar.Enabled = False
        btnRechazar.Enabled = False
    End If
End Sub
Private Sub txtCant_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    ElseIf KeyAscii = 13 Then
        If (txtPrecio.Text = "") And (txtCant.Text <> "") Then
            txtPrecio.SetFocus
        Else
            If (txtPrecio.Text <> "") And (txtCant.Text <> "") Then
                btnAgregar.Value = True
                btnRechazar.Enabled = True
            End If
        End If
    End If
End Sub
Private Sub txtCant_GotFocus()
    txtCant.BackColor = &HFFE1E1
End Sub
Private Sub txtCant_LostFocus()
    txtCant.BackColor = &H80000005
End Sub
Private Sub txtCodBarras_GotFocus()
    txtCodBarras.BackColor = &HFFE1E1
End Sub
Private Sub txtCodBarras_LostFocus()
    txtCodBarras.BackColor = &H80000005
End Sub
Private Sub txtCodBarras_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    ElseIf (KeyAscii = 13) And (txtCodBarras.Text <> "") Then
        Dim tLi As ListItem
        If ListaTemp.ListItems.Count > 0 Then
            Set tLi = ListaEntradasDetalle.ListItems.Add(, , ListaTemp.ListItems.Item(1) & "")
            tLi.SubItems(1) = ListaTemp.ListItems.Item(1).SubItems(1) & ""
            tLi.SubItems(2) = ListaTemp.ListItems.Item(1).SubItems(2) & ""
            tLi.SubItems(3) = ListaTemp.ListItems.Item(1).SubItems(3) & ""
            tLi.SubItems(4) = ListaTemp.ListItems.Item(1).SubItems(4) & ""
            tLi.SubItems(5) = txtCodBarras.Text
            txtCodBarras.Text = ""
            ListaTemp.ListItems.Remove (1)
            If ListaTemp.ListItems.Count = 0 Then
                txtCodBarras.Enabled = False
                btnAgregar.Enabled = True
                btnRechazar.Enabled = True
                txtCodBarras.Text = ""
                txtCant.Text = ""
                txtPrecio.Text = ""
                ListaProdPedidos.Enabled = True
                ListaEntradas.Enabled = True
                btnGuardar.Enabled = True
                btnQuitar.Enabled = True
            Else
                txtCodBarras.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txtNoFactura_Change()
    If (txtNoFactura.Text <> "") Then
        btnAceptarFactura.Enabled = True
    Else
        btnAceptarFactura.Enabled = False
    End If
End Sub
Private Sub txtNoFactura_GotFocus()
    txtNoFactura.BackColor = &HFFE1E1
End Sub
Private Sub txtNoFactura_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    ElseIf (KeyAscii = 13) And (txtNoFactura.Text <> "") Then
            btnAceptarFactura.Value = True
    End If
End Sub
Private Sub txtNoFactura_LostFocus()
    txtNoFactura.BackColor = &H80000005
End Sub
Private Sub txtPrecio_Change()
    If (txtCant.Text <> "") And (txtPrecio.Text <> "") Then
        btnAgregar.Enabled = True
        btnRechazar.Enabled = True
    Else
        btnAgregar.Enabled = False
        btnRechazar.Enabled = False
    End If
End Sub
Private Sub txtPrecio_GotFocus()
    txtPrecio.BackColor = &HFFE1E1
End Sub
Private Sub txtPrecio_LostFocus()
    txtPrecio.BackColor = &H80000005
End Sub
Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    ElseIf KeyAscii = 13 Then
        If txtCant.Text = "" Then
            txtCant.SetFocus
        Else
            If txtPrecio.Text <> "" Then
                btnAgregar.Value = True
                btnRechazar.Enabled = True
            End If
        End If
    End If
End Sub
Private Sub Imprimir()
On Error GoTo ManejaError
    If ListaEntradas.ListItems.Count > 0 Then
        Dim Total As Double
        Total = 0
        CommonDialog1.Flags = 64
        CommonDialog1.CancelError = True
        CommonDialog1.ShowPrinter
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print ""
        Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
        Printer.Print "             SUCURSAL : " & VarMen.Text4(0).Text
        Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & VarMen.Text1(2).Text
        Printer.Print "             FOLIO: " & lblFolio.Caption
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO")) / 2
        Printer.Print "COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("NOMBRE DEL PROVEEDOR:  " & Text1.Text)) / 2
        Printer.Print "NOMBRE DEL PROVEEDOR:  " & Text1.Text
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Dim NRegistros As Integer
        NRegistros = ListaEntradas.ListItems.Count
        Dim Con As Integer
        Dim POSY As Integer
        POSY = 3800
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Clave del Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3500
        Printer.Print "Cantidad Registrada"
        Printer.CurrentY = POSY
        Printer.CurrentX = 6500
        Printer.Print "Precio"
        Printer.CurrentY = POSY
        Printer.CurrentX = 7500
        Printer.Print "Sucursal"
        Printer.CurrentY = POSY
        Printer.CurrentX = 8800
        Printer.Print "Num Orden"
        Printer.CurrentY = POSY
        Printer.CurrentX = 10000
        Printer.Print "Factura"
        POSY = POSY + 200
        For Con = 1 To NRegistros
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print ListaEntradas.ListItems(Con).Text
            Printer.CurrentY = POSY
            Printer.CurrentX = 4000
            Printer.Print Format(CDbl(ListaEntradas.ListItems(Con).SubItems(3)), "0.00")
            Printer.CurrentY = POSY
            Printer.CurrentX = 6500
            Printer.Print Format(CDbl(ListaEntradas.ListItems(Con).SubItems(2)), "0.000")
            Printer.CurrentY = POSY
            Printer.CurrentX = 8800
            Printer.Print Text7.Text
            Printer.CurrentY = POSY
            Printer.CurrentX = 10000
            Printer.Print txtNoFactura.Text
            Total = Total + (Val(Replace(ListaEntradas.ListItems(Con).SubItems(2), ",", "")) * Val(Replace(ListaEntradas.ListItems(Con).SubItems(3), ",", "")))
            Printer.CurrentY = POSY
            Printer.CurrentX = 7500
            Printer.Print "BODEGA"
            If POSY >= 14200 Then
                POSY = 100
                Printer.NewPage
                Printer.Print ""
                Printer.Print ""
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                Printer.Print VarMen.Text5(0).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print ""
                Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
                Printer.Print "             SUCURSAL : " & VarMen.Text4(0).Text
                Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
                Printer.Print "             FOLIO: " & lblFolio.Caption
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO")) / 2
                Printer.Print "COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO"
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("NOMBRE DEL PROVEEDOR:  " & Text1.Text)) / 2
                Printer.Print "NOMBRE DEL PROVEEDOR:  " & Text1.Text
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                POSY = 3800
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print "Clave del Producto"
                Printer.CurrentY = POSY
                Printer.CurrentX = 3500
                Printer.Print "Cantidad Registrada"
                Printer.CurrentY = POSY
                Printer.CurrentX = 6500
                Printer.Print "Precio"
                Printer.CurrentY = POSY
                Printer.CurrentX = 7500
                Printer.Print "Sucursal"
                POSY = POSY + 200
            End If
        Next Con
        Printer.Print ""
        Printer.Print "             Total = " & Total
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.EndDoc
        CommonDialog1.Copies = 1
    End If
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub ImprimirSaldo()
On Error GoTo ManejaError
    If ListaRechazados.ListItems.Count > 0 Then
        Dim Total As Double
        Total = 0
        CommonDialog1.Flags = 64
        CommonDialog1.CancelError = True
        CommonDialog1.ShowPrinter
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print ""
        Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
        Printer.Print "             SUCURSAL : " & VarMen.Text4(0).Text
        Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & VarMen.Text1(2).Text
        Printer.Print "             FOLIO: " & lblFolio.Caption
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE DE REGISTRO DE SALDO POR FALTA DE RECEPCION DE PRODUCTO")) / 2
        Printer.Print "COMPROBANTE DE REGISTRO DE SALDO POR FALTA DE RECEPCION DE PRODUCTO"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("NOMBRE DEL PROVEEDOR:  " & Text1.Text)) / 2
        Printer.Print "NOMBRE DEL PROVEEDOR:  " & Text1.Text
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Dim NRegistros As Integer
        NRegistros = ListaEntradas.ListItems.Count
        Dim Con As Integer
        Dim POSY As Integer
        POSY = 3800
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Clave del Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3500
        Printer.Print "Cantidad Registrada"
        Printer.CurrentY = POSY
        Printer.CurrentX = 6500
        Printer.Print "Precio"
        Printer.CurrentY = POSY
        Printer.CurrentX = 7500
        Printer.Print "Sucursal"
        Printer.CurrentY = POSY
        Printer.CurrentX = 8800
        Printer.Print "Num Orden"
        Printer.CurrentY = POSY
        Printer.CurrentX = 10000
        Printer.Print "Factura"
        POSY = POSY + 200
        For Con = 1 To NRegistros
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print ListaRechazados.ListItems(Con).Text
            Printer.CurrentY = POSY
            Printer.CurrentX = 4000
            Printer.Print ListaRechazados.ListItems(Con).SubItems(3)
            Printer.CurrentY = POSY
            Printer.CurrentX = 6500
            Printer.Print ListaRechazados.ListItems(Con).SubItems(2)
            Printer.CurrentY = POSY
            Printer.CurrentX = 8800
            Printer.Print Text7.Text
            Printer.CurrentY = POSY
            Printer.CurrentX = 10000
            Printer.Print txtNoFactura.Text
            Total = Total + (Val(Replace(ListaRechazados.ListItems(Con).SubItems(2), ",", "")) * Val(Replace(ListaRechazados.ListItems(Con).SubItems(3), ",", "")))
            Printer.CurrentY = POSY
            Printer.CurrentX = 7500
            Printer.Print "BODEGA"
            If POSY >= 14200 Then
                POSY = 100
                Printer.NewPage
                Printer.Print ""
                Printer.Print ""
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                Printer.Print VarMen.Text5(0).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print ""
                Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
                Printer.Print "             SUCURSAL : " & VarMen.Text4(0).Text
                Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
                Printer.Print "             FOLIO: " & lblFolio.Caption
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE DE REGISTRO DE SALDO POR FALTA DE RECEPCION DE PRODUCTO")) / 2
                Printer.Print "COMPROBANTE DE REGISTRO DE SALDO POR FALTA DE RECEPCION DE PRODUCTO"
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("NOMBRE DEL PROVEEDOR:  " & Text1.Text)) / 2
                Printer.Print "NOMBRE DEL PROVEEDOR:  " & Text1.Text
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                POSY = 3800
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print "Clave del Producto"
                Printer.CurrentY = POSY
                Printer.CurrentX = 3500
                Printer.Print "Cantidad Registrada"
                Printer.CurrentY = POSY
                Printer.CurrentX = 6500
                Printer.Print "Precio"
                Printer.CurrentY = POSY
                Printer.CurrentX = 7500
                Printer.Print "Sucursal"
                POSY = POSY + 200
            End If
        Next Con
        Printer.Print ""
        Printer.Print "             Total = " & Total
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.EndDoc
        CommonDialog1.Copies = 1
    End If
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub

