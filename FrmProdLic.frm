VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmProdLic 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Productos de produccion en Licitaciones"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9360
      TabIndex        =   55
      Top             =   5880
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
         TabIndex        =   56
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmProdLic.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmProdLic.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9360
      TabIndex        =   45
      Top             =   4680
      Width           =   975
      Begin VB.Label Label14 
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
         TabIndex        =   46
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   780
         Left            =   120
         MouseIcon       =   "FrmProdLic.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmProdLic.frx":26F6
         Top             =   120
         Width           =   705
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      PrinterDefault  =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Detalle"
      TabPicture(0)   =   "FrmProdLic.frx":4478
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Producto General"
      TabPicture(1)   =   "FrmProdLic.frx":4494
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo3"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(3)=   "ListView6"
      Tab(1).Control(4)=   "Label17"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Originales"
      TabPicture(2)   =   "FrmProdLic.frx":44B0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Combo2"
      Tab(2).Control(1)=   "Frame8"
      Tab(2).Control(2)=   "Frame5"
      Tab(2).Control(3)=   "ListView5"
      Tab(2).Control(4)=   "Label16"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Almacen 2"
      TabPicture(3)   =   "FrmProdLic.frx":44CC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Combo1"
      Tab(3).Control(1)=   "Frame1"
      Tab(3).Control(2)=   "Frame2"
      Tab(3).Control(3)=   "ListView1"
      Tab(3).Control(4)=   "Label15"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Pedidos"
      TabPicture(4)   =   "FrmProdLic.frx":44E8
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command4"
      Tab(4).Control(1)=   "Command2"
      Tab(4).Control(2)=   "ListView3"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Compras"
      TabPicture(5)   =   "FrmProdLic.frx":4504
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Text13"
      Tab(5).Control(1)=   "ListView7"
      Tab(5).Control(2)=   "ListView8"
      Tab(5).Control(3)=   "Label12"
      Tab(5).ControlCount=   4
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68760
         TabIndex        =   60
         Top             =   6480
         Width           =   2655
      End
      Begin MSComctlLib.ListView ListView7 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   57
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5106
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
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   -74160
         TabIndex        =   53
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -74160
         TabIndex        =   51
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -74160
         TabIndex        =   49
         Top             =   600
         Width           =   2775
      End
      Begin VB.Frame Frame8 
         Caption         =   " Pedir Faltates "
         Height          =   1695
         Left            =   -68520
         TabIndex        =   47
         Top             =   5040
         Width           =   2535
         Begin VB.CommandButton Command9 
            Caption         =   "Todos"
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
            Left            =   720
            Picture         =   "FrmProdLic.frx":4520
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " Pedir Faltates "
         Height          =   1695
         Left            =   -68520
         TabIndex        =   43
         Top             =   5040
         Width           =   2535
         Begin VB.CommandButton Command8 
            Caption         =   "Todos"
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
            Left            =   720
            Picture         =   "FrmProdLic.frx":6EF2
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Detalle "
         Height          =   1695
         Left            =   -74880
         TabIndex        =   35
         Top             =   5040
         Width           =   6255
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   2040
            TabIndex        =   39
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox Text11 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   720
            Width           =   4815
         End
         Begin VB.TextBox Text12 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Agregar"
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
            Picture         =   "FrmProdLic.frx":98C4
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Cantidad Minima a Pedir"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label11 
            Caption         =   "Descripción"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Clave"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " Detalle "
         Height          =   1695
         Left            =   -74880
         TabIndex        =   25
         Top             =   5040
         Width           =   6255
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   2040
            TabIndex        =   29
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   720
            Width           =   4815
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Agregar"
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
            Picture         =   "FrmProdLic.frx":C296
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Cantidad Minima a Pedir"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "Descripción"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Clave"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Detalle "
         Height          =   1575
         Left            =   120
         TabIndex        =   17
         Top             =   5280
         Width           =   6255
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   2040
            TabIndex        =   21
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   720
            Width           =   4815
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Agregar"
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
            Picture         =   "FrmProdLic.frx":EC68
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Cantidad Minima a Pedir"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label5 
            Caption         =   "Descripción"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Clave"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2055
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3625
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
      Begin VB.CommandButton Command4 
         Caption         =   "Borrar Todo"
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
         Picture         =   "FrmProdLic.frx":1163A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
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
         Left            =   -67320
         Picture         =   "FrmProdLic.frx":1400C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   " Detalle "
         Height          =   1695
         Left            =   -74880
         TabIndex        =   3
         Top             =   5040
         Width           =   6255
         Begin VB.CommandButton Command1 
            Caption         =   "Agregar"
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
            Picture         =   "FrmProdLic.frx":169DE
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   720
            Width           =   4815
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   2040
            TabIndex        =   4
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Clave"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Descripción"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Cantidad Minima a Pedir"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1200
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Pedir Faltates "
         Height          =   1695
         Left            =   -68520
         TabIndex        =   1
         Top             =   5040
         Width           =   2535
         Begin VB.CommandButton Command3 
            Caption         =   "Todos"
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
            Left            =   720
            Picture         =   "FrmProdLic.frx":193B0
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   720
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   11
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6800
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
      Begin MSComctlLib.ListView ListView3 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   13
         Top             =   600
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9551
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
      Begin MSComctlLib.ListView ListView4 
         Height          =   2415
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4260
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
      Begin MSComctlLib.ListView ListView5 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   33
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6800
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
      Begin MSComctlLib.ListView ListView6 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   34
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6800
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
      Begin MSComctlLib.ListView ListView8 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   58
         Top             =   3360
         Width           =   8775
         _ExtentX        =   15478
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
      Begin VB.Label Label12 
         Caption         =   "Total :"
         Height          =   255
         Left            =   -69360
         TabIndex        =   59
         Top             =   6480
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Marca"
         Height          =   255
         Left            =   -74760
         TabIndex        =   54
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Marca"
         Height          =   255
         Left            =   -74760
         TabIndex        =   52
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Marca"
         Height          =   255
         Left            =   -74760
         TabIndex        =   50
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   9360
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "FrmProdLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim Exis As String
Dim CanPedi As String
Dim RepVar As String
Dim IndElim As Integer
Private Sub Combo1_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    If Combo1.Text <> "<TODAS>" Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, SUM(CAN_MIN_LIC) AS CAN_MIN_LIC2, SUM(CAN_MAX_LIC) AS CAN_MAX_LIC2, EXISTENCIA FROM VsJuegoRepLicitacion WHERE SUCURSAL = 'BODEGA' AND MARCA = '" & Combo1.Text & "' GROUP BY ID_PRODUCTO, DESCRIPCION, EXISTENCIA ORDER BY ID_PRODUCTO"
    Else
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, SUM(CAN_MIN_LIC) AS CAN_MIN_LIC2, SUM(CAN_MAX_LIC) AS CAN_MAX_LIC2, EXISTENCIA FROM VsJuegoRepLicitacion WHERE SUCURSAL = 'BODEGA' GROUP BY ID_PRODUCTO, DESCRIPCION, EXISTENCIA ORDER BY ID_PRODUCTO"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CAN_MIN_LIC2")) Then tLi.SubItems(2) = tRs.Fields("CAN_MIN_LIC2")
            If Not IsNull(tRs.Fields("CAN_MAX_LIC2")) Then tLi.SubItems(3) = tRs.Fields("CAN_MAX_LIC2")
            If Not IsNull(tRs.Fields("EXISTENCIA")) Then tLi.SubItems(4) = tRs.Fields("EXISTENCIA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Combo1_DropDown()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Combo1.Clear
    sBuscar = "SELECT MARCA FROM VsJuegoRepLicitacion GROUP BY MARCA ORDER BY MARCA"
    Set tRs = cnn.Execute(sBuscar)
    Combo1.AddItem "<TODAS>"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("MARCA")) Then Combo1.AddItem tRs.Fields("MARCA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Combo2_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If Combo2.Text = "<TODAS>" Then
        sBuscar = "SELECT * FROM VsLicProd WHERE TIPO = 'SIMPLE' ORDER BY ID_PRODUCTO"
    Else
        sBuscar = "SELECT * FROM VsLicProd WHERE TIPO = 'SIMPLE' AND MARCA = '" & Combo2.Text & "' ORDER BY ID_PRODUCTO"
    End If
    Set tRs = cnn.Execute(sBuscar)
    ListView5.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView5.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANT_MIN")) Then tLi.SubItems(2) = tRs.Fields("CANT_MIN")
            If Not IsNull(tRs.Fields("CANT_MAX")) Then tLi.SubItems(3) = tRs.Fields("CANT_MAX")
            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(4) = tRs.Fields("PRECIO_VENTA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Combo2_DropDown()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Combo2.Clear
    sBuscar = "SELECT MARCA FROM VsLicProd GROUP BY MARCA ORDER BY MARCA"
    Set tRs = cnn.Execute(sBuscar)
    Combo2.AddItem "<TODAS>"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("MARCA")) Then Combo2.AddItem tRs.Fields("MARCA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Combo3_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If Combo3.Text = "<TODAS>" Then
        sBuscar = "SELECT * FROM VsProdGralLic WHERE SUCURSAL = 'BODEGA' ORDER BY ID_PRODUCTO"
    Else
        sBuscar = "SELECT * FROM VsProdGralLic WHERE SUCURSAL = 'BODEGA' AND MARCA = '" & Combo3.Text & "' ORDER BY ID_PRODUCTO"
    End If
    ListView6.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView6.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANT_MIN")) Then tLi.SubItems(2) = tRs.Fields("CANT_MIN")
            If Not IsNull(tRs.Fields("CANT_MAX")) Then tLi.SubItems(3) = tRs.Fields("CANT_MAX")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(4) = tRs.Fields("CANTIDAD")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Combo3_DropDown()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Combo3.Clear
    sBuscar = "SELECT MARCA FROM VsProdGralLic GROUP BY MARCA ORDER BY MARCA"
    Set tRs = cnn.Execute(sBuscar)
    Combo3.AddItem "<TODAS>"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("MARCA")) Then Combo3.AddItem tRs.Fields("MARCA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command1_Click()
    If Text3.Text = "" Then
        Text3.Text = "0"
    End If
    If CDbl(Text3.Text) > 0 Then
        Dim tLi As ListItem
        If CDbl(CanPedi) > CDbl(Text3.Text) Then
            MsgBox "NO PUEDE PEDIR MENOS DE LO NECESARIO PARA COMPLETAR EL PEDIDO!", vbInformation, "SACC"
            Text3.Text = CanPedi
        Else
            Set tLi = ListView3.ListItems.Add(, , Text1.Text)
            If Not IsNull(Text2.Text) Then tLi.SubItems(1) = Text2.Text
            If Not IsNull(Text3.Text) Then tLi.SubItems(2) = Text3.Text
            If Not IsNull(Exis) Then tLi.SubItems(3) = Exis
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Exis = ""
            CanPedi = ""
        End If
    Else
        MsgBox "NO SE PUEDE PEDIR CANTIDAD CERO!", vbInformation, "SACC"
    End If
End Sub
Private Sub Command2_Click()
    If IndElim > 0 Then
        ListView3.ListItems.Remove (IndElim)
        IndElim = 0
    End If
End Sub
Private Sub Command3_Click()
    Dim tLi As ListItem
    Dim Con As Integer
    Dim NRegistros As Integer
    NRegistros = ListView1.ListItems.Count
    For Con = 1 To NRegistros
        If Format(CDbl(ListView1.ListItems(Con).SubItems(2)) - CDbl(ListView1.ListItems(Con).SubItems(4)), "###,###,##0.00") > 0 Then
            Set tLi = ListView3.ListItems.Add(, , ListView1.ListItems(Con).Text)
            If Not IsNull(ListView1.ListItems(Con).SubItems(1)) Then tLi.SubItems(1) = ListView1.ListItems(Con).SubItems(1)
            If Not IsNull(ListView1.ListItems(Con).SubItems(2)) Or Not IsNull(ListView1.ListItems(Con).SubItems(4)) Then tLi.SubItems(2) = Format(CDbl(ListView1.ListItems(Con).SubItems(2)) - CDbl(ListView1.ListItems(Con).SubItems(4)), "###,###,##0.00")
            If Not IsNull(ListView1.ListItems(Con).SubItems(4)) Then tLi.SubItems(3) = ListView1.ListItems(Con).SubItems(4)
        End If
    Next Con
End Sub
Private Sub Command4_Click()
    ListView3.ListItems.Clear
End Sub
Private Sub Command5_Click()
    Dim tLi As ListItem
    If Text6.Text = "" Then
        Text6.Text = "0"
    End If
    If CDbl(Text6.Text) > 0 Then
        If CDbl(CanPedi) > CDbl(Text6.Text) Then
            MsgBox "NO PUEDE PEDIR MENOS DE LO NECESARIO PARA COMPLETAR EL PEDIDO!", vbInformation, "SACC"
            Text6.Text = CanPedi
        Else
            Set tLi = ListView3.ListItems.Add(, , Text4.Text)
            If Not IsNull(Text2.Text) Then tLi.SubItems(1) = Text5.Text
            If Not IsNull(Text3.Text) Then tLi.SubItems(2) = Text6.Text
            If Not IsNull(Exis) Then tLi.SubItems(3) = Exis
            Text4.Text = ""
            Text5.Text = ""
            Text6.Text = ""
            CanPedi = 0
            Exis = ""
        End If
    Else
        MsgBox "NO SE PUEDE PEDIR CANTIDAD EN CERO!", vbInformation, "SACC"
    End If
End Sub
Private Sub Command6_Click()
    Dim tLi As ListItem
    If Text10.Text = "" Then
        Text10.Text = "0"
    End If
    If CDbl(Text10.Text) > 0 Then
        If CDbl(CanPedi) > CDbl(Text10.Text) Then
            MsgBox "NO PUEDE PEDIR MENOS DE LO NECESARIO PARA COMPLETAR EL PEDIDO!", vbInformation, "SACC"
            Text3.Text = CanPedi
        Else
            Set tLi = ListView3.ListItems.Add(, , Text12.Text)
            If Not IsNull(Text2.Text) Then tLi.SubItems(1) = Text11.Text
            If Not IsNull(Text3.Text) Then tLi.SubItems(2) = Text10.Text
            If Not IsNull(Exis) Then tLi.SubItems(3) = Exis
            Text12.Text = ""
            Text11.Text = ""
            Text10.Text = ""
            Exis = ""
            CanPedi = ""
        End If
    Else
        MsgBox "NO PUEDE PEDIR CANTIDAD EN CEROS!", vbInformation, "SACC"
    End If
End Sub
Private Sub Command7_Click()
    If Text9.Text = "" Then
        Text9.Text = "0"
    End If
    If CDbl(Text9.Text) > 0 Then
        Dim tLi As ListItem
        If CDbl(CanPedi) > CDbl(Text9.Text) Then
            MsgBox "NO PUEDE PEDIR MENOS DE LO NECESARIO PARA COMPLETAR EL PEDIDO!", vbInformation, "SACC"
            Text6.Text = CanPedi
        Else
            Set tLi = ListView3.ListItems.Add(, , Text7.Text)
            If Not IsNull(Text8.Text) Then tLi.SubItems(1) = Text8.Text
            If Not IsNull(Text9.Text) Then tLi.SubItems(2) = Text9.Text
            If Not IsNull(Exis) Then tLi.SubItems(3) = Exis
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            CanPedi = 0
            Exis = ""
        End If
    Else
        MsgBox "NO SE PUEDE PEDIR CANTIDAD EN CERO!", vbInformation, "SACC"
    End If
End Sub
Private Sub Command8_Click()
    Dim tLi As ListItem
    Dim Con As Integer
    Dim NRegistros As Integer
    NRegistros = ListView6.ListItems.Count
    For Con = 1 To NRegistros
        If Format(CDbl(ListView6.ListItems(Con).SubItems(2)) - CDbl(ListView6.ListItems(Con).SubItems(4)), "###,###,##0.00") > 0 Then
            Set tLi = ListView3.ListItems.Add(, , ListView6.ListItems(Con).Text)
            If Not IsNull(ListView6.ListItems(Con).SubItems(1)) Then tLi.SubItems(1) = ListView6.ListItems(Con).SubItems(1)
            If Not IsNull(ListView6.ListItems(Con).SubItems(2)) Or Not IsNull(ListView6.ListItems(Con).SubItems(4)) Then tLi.SubItems(2) = Format(CDbl(ListView6.ListItems(Con).SubItems(2)) - CDbl(ListView6.ListItems(Con).SubItems(4)), "###,###,##0.00")
            If Not IsNull(ListView6.ListItems(Con).SubItems(4)) Then tLi.SubItems(3) = ListView6.ListItems(Con).SubItems(4)
        End If
    Next Con
End Sub
Private Sub Command9_Click()
    Dim tLi As ListItem
    Dim Con As Integer
    Dim NRegistros As Integer
    NRegistros = ListView5.ListItems.Count
    For Con = 1 To NRegistros
        If Format(CDbl(ListView5.ListItems(Con).SubItems(2)) - CDbl(ListView5.ListItems(Con).SubItems(4)), "###,###,##0.00") > 0 Then
            Set tLi = ListView3.ListItems.Add(, , ListView5.ListItems(Con).Text)
            If Not IsNull(ListView5.ListItems(Con).SubItems(1)) Then tLi.SubItems(1) = ListView5.ListItems(Con).SubItems(1)
            If Not IsNull(ListView5.ListItems(Con).SubItems(2)) Or Not IsNull(ListView5.ListItems(Con).SubItems(4)) Then tLi.SubItems(2) = Format(CDbl(ListView5.ListItems(Con).SubItems(2)) - CDbl(ListView5.ListItems(Con).SubItems(4)), "###,###,##0.00")
            If Not IsNull(ListView5.ListItems(Con).SubItems(4)) Then tLi.SubItems(3) = ListView5.ListItems(Con).SubItems(4)
        End If
    Next Con
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBuscar As String
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.txtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Cantidad Minima Pedida", 1500
        .ColumnHeaders.Add , , "Cantidad Maxima Pedida", 1500
        .ColumnHeaders.Add , , "Existencia", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Cliente", 0
        .ColumnHeaders.Add , , "Nombre", 5800
        .ColumnHeaders.Add , , "No. Contrato", 1500
        .ColumnHeaders.Add , , "Fecha Fin", 1500
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Cantidad Para Requisición", 1500
        .ColumnHeaders.Add , , "Existencia", 1500
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Cantidad Minima", 1500
        .ColumnHeaders.Add , , "Cantidad Maxima", 1500
        .ColumnHeaders.Add , , "Precio de Vanta", 1500
        .ColumnHeaders.Add , , "Total Comprado", 1500
    End With
    With ListView5
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Cantidad Minima", 1500
        .ColumnHeaders.Add , , "Cantidad Maxima", 1500
        .ColumnHeaders.Add , , "Precio de Vanta", 1500
    End With
    With ListView6
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Cantidad Minima", 1500
        .ColumnHeaders.Add , , "Cantidad Maxima", 1500
        .ColumnHeaders.Add , , "Existencia", 1500
    End With
    With ListView7
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Cliente", 2000
        .ColumnHeaders.Add , , "Nombre", 5500
        .ColumnHeaders.Add , , "No. Contrato", 1500
        .ColumnHeaders.Add , , "No. Licitacion", 1500
        .ColumnHeaders.Add , , "Fecha Fin", 1500
    End With
    With ListView8
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Total", 1500
        .ColumnHeaders.Add , , "Precio de Venta", 1500
    End With
    sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, SUM(CAN_MIN_LIC) AS CAN_MIN_LIC2, SUM(CAN_MAX_LIC) AS CAN_MAX_LIC2, EXISTENCIA FROM VsJuegoRepLicitacion WHERE SUCURSAL = 'BODEGA' GROUP BY ID_PRODUCTO, DESCRIPCION, EXISTENCIA ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CAN_MIN_LIC2")) Then tLi.SubItems(2) = tRs.Fields("CAN_MIN_LIC2")
            If Not IsNull(tRs.Fields("CAN_MAX_LIC2")) Then tLi.SubItems(3) = tRs.Fields("CAN_MAX_LIC2")
            If Not IsNull(tRs.Fields("EXISTENCIA")) Then tLi.SubItems(4) = tRs.Fields("EXISTENCIA")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT * FROM VsClienLic WHERE FECHA_FIN >= DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())) ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("NO_CONTRATO")) Then tLi.SubItems(2) = tRs.Fields("NO_CONTRATO")
            If Not IsNull(tRs.Fields("FECHA_FIN")) Then tLi.SubItems(3) = tRs.Fields("FECHA_FIN")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT * FROM VsLicProd WHERE TIPO = 'SIMPLE' ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView5.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANT_MIN")) Then tLi.SubItems(2) = tRs.Fields("CANT_MIN")
            If Not IsNull(tRs.Fields("CANT_MAX")) Then tLi.SubItems(3) = tRs.Fields("CANT_MAX")
            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(4) = tRs.Fields("PRECIO_VENTA")
            tRs.MoveNext
        Loop
    End If
    'VsProdGralLic
    sBuscar = "SELECT * FROM VsProdGralLic WHERE SUCURSAL = 'BODEGA' ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView6.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANT_MIN")) Then tLi.SubItems(2) = tRs.Fields("CANT_MIN")
            If Not IsNull(tRs.Fields("CANT_MAX")) Then tLi.SubItems(3) = tRs.Fields("CANT_MAX")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(4) = tRs.Fields("CANTIDAD")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT CLIENTE.ID_CLIENTE, CLIENTE.NOMBRE, LICITACIONES.NO_CONTRATO, LICITACIONES.NO_LICITACION, LICITACIONES.FECHA_FIN FROM LICITACIONES INNER JOIN CLIENTE ON LICITACIONES.ID_CLIENTE = CLIENTE.ID_CLIENTE GROUP BY CLIENTE.ID_CLIENTE, CLIENTE.NOMBRE, LICITACIONES.NO_CONTRATO, LICITACIONES.NO_LICITACION, LICITACIONES.FECHA_FIN ORDER BY LICITACIONES.FECHA_FIN"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView7.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("NO_CONTRATO")) Then tLi.SubItems(2) = tRs.Fields("NO_CONTRATO")
            If Not IsNull(tRs.Fields("NO_LICITACION")) Then tLi.SubItems(3) = tRs.Fields("NO_LICITACION")
            If Not IsNull(tRs.Fields("FECHA_FIN")) Then tLi.SubItems(4) = tRs.Fields("FECHA_FIN")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image3_Click()
    If SSTab1.Tab = 1 Then
        Imprimir ListView6, "         PRODUCTOS PARA LICITACIÓNES"
    Else
        If SSTab1.Tab = 2 Then
            Imprimir ListView5, "PRODUCTOS ORIGINALES PARA LICITACIÓNES"
        Else
            If SSTab1.Tab = 3 Then
                Imprimir ListView1, "MATERIA PRIMA PARA LICITACIÓNES"
            Else
                If SSTab1.Tab = 4 Then
                    Imprimir ListView3, "PRODUCTOS FALTANTES PARA LICITACIÓNES"
                Else
                    If RepVar <> "" Then
                        ImprimirRep
                    End If
                End If
            End If
        End If
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    ListView1.SortOrder = 1 Xor ListView1.SortOrder
End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    ListView2.SortOrder = 1 Xor ListView2.SortOrder
End Sub
Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView3.SortKey = ColumnHeader.Index - 1
    ListView3.Sorted = True
    ListView3.SortOrder = 1 Xor ListView3.SortOrder
End Sub
Private Sub ListView4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView4.SortKey = ColumnHeader.Index - 1
    ListView4.Sorted = True
    ListView4.SortOrder = 1 Xor ListView4.SortOrder
End Sub
Private Sub ListView5_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView5.SortKey = ColumnHeader.Index - 1
    ListView5.Sorted = True
    ListView5.SortOrder = 1 Xor ListView5.SortOrder
End Sub
Private Sub ListView6_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView6.SortKey = ColumnHeader.Index - 1
    ListView6.Sorted = True
    ListView6.SortOrder = 1 Xor ListView6.SortOrder
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item
    Text2.Text = Item.SubItems(1)
    If Format(CDbl(Item.SubItems(2)) - CDbl(Item.SubItems(4)), "###,###,##0.00") < 0 Then
        Text3.Text = "0.00"
    Else
        Text3.Text = Format(CDbl(Item.SubItems(2)) - CDbl(Item.SubItems(4)), "###,###,##0.00")
    End If
    Exis = Item.SubItems(4)
    CanPedi = Format(CDbl(Item.SubItems(2)) - CDbl(Item.SubItems(4)), "###,###,##0.00")
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView4.ListItems.Clear
    RepVar = Item.SubItems(2)
    'sBuscar = "SELECT LICITACIONES.ID_CLIENTE, VENTAS.NOMBRE, LICITACIONES.NO_CONTRATO, LICITACIONES.ID_PRODUCTO, SUM(VENTAS_DETALLE.CANTIDAD) AS CANTIDAD, LICITACIONES.CANT_MAX, LICITACIONES.CANT_MIN FROM VENTAS_DETALLE INNER JOIN VENTAS ON VENTAS_DETALLE.ID_VENTA = VENTAS.ID_VENTA INNER JOIN LICITACIONES ON VENTAS.ID_CLIENTE = LICITACIONES.ID_CLIENTE AND VENTAS_DETALLE.ID_PRODUCTO = LICITACIONES.ID_PRODUCTO AND VENTAS.FECHA <= LICITACIONES.FECHA_FIN AND VENTAS.FECHA >= LICITACIONES.FECHA_INICIO GROUP BY LICITACIONES.ID_CLIENTE, VENTAS.NOMBRE, LICITACIONES.ID_PRODUCTO, LICITACIONES.NO_CONTRATO, LICITACIONES.NO_LICITACION, LICITACIONES.CANT_MAX, LICITACIONES.CANT_MIN"
    sBuscar = "SELECT * FROM VsLicProd WHERE NO_CONTRATO = '" & Item.SubItems(2) & "' ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView4.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANT_MIN")) Then tLi.SubItems(2) = tRs.Fields("CANT_MIN")
            If Not IsNull(tRs.Fields("CANT_MAX")) Then tLi.SubItems(3) = tRs.Fields("CANT_MAX")
            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(4) = tRs.Fields("PRECIO_VENTA")
            If Not IsNull(tRs.Fields("TOTAL_COMPRADO")) Then
                tLi.SubItems(5) = tRs.Fields("TOTAL_COMPRADO")
            Else
                tLi.SubItems(5) = "0"
            End If
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IndElim = Item.Index
End Sub
Private Sub ListView4_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text4.Text = Item
    Text5.Text = Item.SubItems(1)
    Text6.Text = Item.SubItems(2)
    CanPedi = Item.SubItems(2)
End Sub
Private Sub ListView5_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text7.Text = Item
    Text8.Text = Item.SubItems(1)
    Text9.Text = Item.SubItems(2)
    CanPedi = Item.SubItems(2)
End Sub
Private Sub ListView6_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text12.Text = Item
    Text11.Text = Item.SubItems(1)
    If Format(CDbl(Item.SubItems(2)) - CDbl(Item.SubItems(4)), "###,###,##0.00") < 0 Then
        Text10.Text = "0.00"
    Else
        Text10.Text = Format(CDbl(Item.SubItems(2)) - CDbl(Item.SubItems(4)), "###,###,##0.00")
    End If
    Exis = Item.SubItems(4)
    CanPedi = Format(CDbl(Item.SubItems(2)) - CDbl(Item.SubItems(4)), "###,###,##0.00")
End Sub
Private Sub Imprimir(LV As ListView, Titulo As String)
On Error GoTo Nada
    CommonDialog1.Flags = 64
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    Dim POSY As Integer
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
    Printer.Print VarMen.Text5(0).Text
    Printer.Print "                                                                                              " & Titulo
    Printer.Print "                                                                                                                  LISTA DE PRODUCTOS "
    Printer.Print ""
    Printer.Print "     FECHA DE IMPRESION : " & Format(Date, "DD/MM/YYYY")
    Printer.Print "     USUARIO QUE IMPRIMIO : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
    Printer.Print ""
    Printer.Print "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print ""
    Printer.Print ""
    POSY = 2200
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "PRODUCTO"
    Printer.CurrentY = POSY
    Printer.CurrentX = 2200
    Printer.Print "Descripcion"
    Printer.CurrentY = POSY
    Printer.CurrentX = 10000
    Printer.Print "CANTIDAD"
    Dim Con As Integer
    Dim NRegistros As Integer
    NRegistros = LV.ListItems.Count
    For Con = 1 To NRegistros
        POSY = POSY + 200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print LV.ListItems(Con).Text
        Printer.CurrentY = POSY
        Printer.CurrentX = 2200
        Printer.Print LV.ListItems(Con).SubItems(1)
        Printer.CurrentY = POSY
        Printer.CurrentX = 10000
        Printer.Print LV.ListItems(Con).SubItems(2)
        If POSY >= 14200 Then
            Printer.NewPage
            POSY = 200
            Printer.Print ""
            Printer.Print ""
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
            Printer.Print VarMen.Text5(0).Text
            Printer.Print "                                                                                                     & Titulo"
            Printer.Print "                                                                                                                  LISTA DE PRODUCTOS "
            Printer.Print ""
            Printer.Print "     FECHA DE IMPRESION : " & Format(Date, "DD/MM/YYYY")
            Printer.Print "     USUARIO QUE IMPRIMIO : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
            Printer.Print ""
            Printer.Print "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print ""
            Printer.Print ""
            POSY = 2200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print "PRODUCTO"
            Printer.CurrentY = POSY
            Printer.CurrentX = 2200
            Printer.Print "Descripcion"
            Printer.CurrentY = POSY
            Printer.CurrentX = 10000
            Printer.Print "CANTIDAD"
        End If
    Next Con
    POSY = POSY + 200
    Printer.CurrentY = POSY
    Printer.CurrentX = 0
    Printer.Print "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    POSY = POSY + 200
    Printer.CurrentY = POSY
    Printer.CurrentX = 0
    Printer.Print "            FIN DEL DOCUMENTO."
    Printer.EndDoc
    CommonDialog1.Copies = 1
Exit Sub
Nada:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub ImprimirRep()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    ConPag = 1
    'text1.text = no_orden
    'nvomen.Text1(0).Text = id_usuario
    sBuscar = "SELECT * FROM VsLicProd WHERE NO_CONTRATO = '" & RepVar & "' ORDER BY ID_PRODUCTO"
    'sBuscar = "SELECT CLIENTE.ID_CLIENTE, CLIENTE.NOMBRE, CLIENTE.DIRECCION, CLIENTE.COLONIA, CLIENTE.CIUDAD, LICITACIONES.NO_CONTRATO, LICITACIONES.NO_LICITACION, LICITACIONES.FECHA_INICIO, LICITACIONES.FECHA_FIN, LICITACIONES.ID_PRODUCTO, ALMACEN3.Descripcion, LICITACIONES.CANT_MIN, LICITACIONES.CANT_MAX, (SELECT COUNT(ID_PRODUCTO) From Ventas, VENTAS_DETALLE WHERE VENTAS.ID_VENTA = VENTAS_DETALLE.ID_VENTA AND VENTAS.FECHA BETWEEN LICITACIONES.FECHA_INICIO AND LICITACIONES.FECHA_FIN AND VENTAS_DETALLE.ID_PRODUCTO = LICITACIONES.ID_PRODUCTO) AS COMPRADO, LICITACIONES.Precio_Venta FROM LICITACIONES INNER JOIN CLIENTE ON LICITACIONES.ID_CLIENTE = CLIENTE.ID_CLIENTE INNER JOIN ALMACEN3 ON LICITACIONES.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO AND LICITACIONES.NO_CONTRATO = '" & RepVar & "'"
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\Licitaciones.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier_Bold, MacRomanEncoding
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        ' Encabezado del reporte
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 38, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 175, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 175, tRs4.Fields("DIRECCION"), "F3", 8, hLeft
        oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 70, 205, 100, 175, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 80, 205, 100, 175, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 90, 205, 100, 175, "REPORTE DE DETALLE DE LICITACION", "F2", 10, hCenter
        oDoc.WTextBox 50, 380, 20, 250, "Fecha :" & Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
        ' LLENADO DE LAS CAJAS
        'CAJA1
        oDoc.WTextBox 115, 20, 100, 175, "No. Cliente " & tRs1.Fields("ID_CLIENTE"), "F3", 8, hCenter
        oDoc.WTextBox 135, 20, 100, 375, tRs1.Fields("NOMBRE"), "F3", 8, hLeft
        oDoc.WTextBox 145, 20, 100, 375, tRs1.Fields("DIRECCION"), "F3", 8, hLeft
        oDoc.WTextBox 155, 20, 100, 375, tRs1.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 165, 20, 100, 375, tRs1.Fields("CIUDAD"), "F3", 8, hLeft
        oDoc.WTextBox 115, 405, 100, 105, "Contrato: " & tRs1.Fields("NO_CONTRATO"), "F3", 8, hLeft
        oDoc.WTextBox 125, 405, 100, 105, "Licitacion: " & tRs1.Fields("NO_LICITACION"), "F3", 8, hLeft
        oDoc.WTextBox 135, 405, 100, 105, "Inicio " & tRs1.Fields("FECHA_INICIO"), "F3", 8, hLeft
        oDoc.WTextBox 145, 405, 100, 105, "Fin " & tRs1.Fields("FECHA_FIN"), "F3", 8, hLeft
        'CAJA2
        'CAJA3
        Posi = 210
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 20, 90, "Producto", "F2", 8, hCenter
        oDoc.WTextBox Posi, 100, 20, 50, "DESCRIPCION", "F2", 8, hCenter
        oDoc.WTextBox Posi, 394, 20, 30, "Minimo", "F2", 8, hCenter
        oDoc.WTextBox Posi, 445, 20, 30, "Maximo", "F2", 8, hCenter
        oDoc.WTextBox Posi, 492, 20, 40, "Comprado", "F2", 8, hCenter
        oDoc.WTextBox Posi, 542, 20, 30, "Precio", "F2", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        'vsordenesrep
        If Not (tRs1.EOF And tRs1.BOF) Then
            Do While Not tRs1.EOF
                oDoc.WTextBox Posi, 20, 20, 90, tRs1.Fields("ID_PRODUCTO"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 110, 20, 360, tRs1.Fields("Descripcion"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 380, 20, 30, tRs1.Fields("CANT_MIN"), "F3", 7, hRight
                oDoc.WTextBox Posi, 435, 20, 30, tRs1.Fields("CANT_MAX"), "F3", 7, hRight
                oDoc.WTextBox Posi, 492, 20, 30, tRs1.Fields("TOTAL_COMPRADO"), "F3", 7, hRight
                oDoc.WTextBox Posi, 542, 20, 30, Format(tRs1.Fields("PRECIO_VENTA"), "###,###,##0.00"), "F3", 7, hRight
                Posi = Posi + 12
                tRs1.MoveNext
                If Posi >= 800 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    oDoc.NewPage A4_Vertical
                    Posi = 50
                    oDoc.WImage 50, 40, 43, 161, "Logo"
                    oDoc.WTextBox 40, 205, 100, 175, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
                    oDoc.WTextBox 60, 224, 100, 175, tRs4.Fields("DIRECCION"), "F3", 8, hLeft
                    oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hLeft
                    oDoc.WTextBox 70, 205, 100, 175, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
                    oDoc.WTextBox 80, 205, 100, 175, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
                    oDoc.WTextBox 90, 205, 100, 175, "REPORTE DE DETALLE DE LICITACION", "F2", 10, hCenter
                    oDoc.WTextBox 50, 380, 20, 250, "Fecha :" & Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
                    ' LLENADO DE LAS CAJAS
                    'CAJA1
                    oDoc.WTextBox 115, 20, 100, 175, "No. Cliente " & tRs1.Fields("ID_CLIENTE"), "F3", 8, hCenter
                    oDoc.WTextBox 135, 20, 100, 175, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
                    oDoc.WTextBox 145, 20, 100, 175, tRs1.Fields("DIRECCION"), "F3", 8, hCenter
                    oDoc.WTextBox 155, 20, 100, 175, tRs1.Fields("COLONIA"), "F3", 8, hCenter
                    oDoc.WTextBox 165, 20, 100, 175, tRs1.Fields("CIUDAD"), "F3", 8, hCenter
                    oDoc.WTextBox 115, 405, 100, 105, "Contrato: " & tRs1.Fields("NO_CONTRATO"), "F3", 8, hLeft
                    oDoc.WTextBox 125, 405, 100, 105, "Licitacion: " & tRs1.Fields("NO_LICITACION"), "F3", 8, hLeft
                    oDoc.WTextBox 135, 405, 100, 105, "Inicio " & tRs1.Fields("FECHA_INICIO"), "F3", 8, hLeft
                    oDoc.WTextBox 145, 405, 100, 105, "Fin " & tRs1.Fields("FECHA_FIN"), "F3", 8, hLeft
                    'CAJA2
                    'CAJA3
                    Posi = 210
                    ' ENCABEZADO DEL DETALLE
                    oDoc.WTextBox Posi, 5, 20, 90, "Producto", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 100, 20, 50, "DESCRIPCION", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 394, 20, 30, "Minimo", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 445, 20, 30, "Maximo", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 492, 20, 40, "Comprado", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 542, 20, 30, "Precio", "F2", 8, hCenter
                    Posi = Posi + 12
                End If
            Loop
        End If
        Posi = Posi + 10
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub ListView7_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBuscar As String
    Dim Total As Double
    Total = 0
    ListView8.ListItems.Clear
    sBuscar = "SELECT VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.DESCRIPCION, SUM(VENTAS_DETALLE.CANTIDAD) AS TOTAL, VENTAS_DETALLE.Precio_Venta , Ventas.ID_CLIENTE FROM VENTAS INNER JOIN VENTAS_DETALLE ON VENTAS.ID_VENTA = VENTAS_DETALLE.ID_VENTA INNER JOIN LICITACIONES ON VENTAS.ID_CLIENTE = LICITACIONES.ID_CLIENTE AND VENTAS.FECHA BETWEEN LICITACIONES.FECHA_INICIO AND LICITACIONES.FECHA_FIN AND VENTAS_DETALLE.ID_PRODUCTO = LICITACIONES.ID_PRODUCTO WHERE  LICITACIONES.NO_CONTRATO = '" & Item.SubItems(2) & "' AND VENTAS.ID_CLIENTE = " & Item & " GROUP BY VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.DESCRIPCION, VENTAS_DETALLE.PRECIO_VENTA, VENTAS.ID_CLIENTE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView8.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("DESCRIPCION")) Then tLi.SubItems(1) = tRs.Fields("DESCRIPCION")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(2) = tRs.Fields("TOTAL")
            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(3) = tRs.Fields("PRECIO_VENTA")
            Total = Total + (CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("TOTAL")))
            tRs.MoveNext
        Loop
    End If
    Text13.Text = Format(Total, "###,###,##0.00")
End Sub

