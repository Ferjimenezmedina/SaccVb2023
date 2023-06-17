VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAutOC 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorizar Orden de Compra"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9960
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   64
      Top             =   3480
      Width           =   975
      Begin VB.Image Image4 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmAutOC.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmAutOC.frx":030A
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EXCEL"
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
         TabIndex        =   65
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   5280
      TabIndex        =   16
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Generales"
      TabPicture(0)   =   "frmAutOC.frx":1E4C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblFolio"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblMoneda"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCantOrg"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtCant"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "opnIndirecta"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "opnInternacional"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "opnNacional"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDescuento"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtFlete"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtCargos"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtImpuesto"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtSubtotal"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtTotal"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Frame3"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Administracion"
      TabPicture(1)   =   "frmAutOC.frx":1E68
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   "Regresar Orden de Compra"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   49
         Top             =   3120
         Width           =   4215
         Begin VB.CommandButton Command4 
            Caption         =   "Regresar"
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
            Left            =   2280
            Picture         =   "frmAutOC.frx":1E84
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   1560
            Width           =   1215
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Indirecta"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   1560
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Internacional"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   1320
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Nacional"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1080
            TabIndex        =   56
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label18 
            Caption         =   "No. Orden :"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cancelar Orden de Compra"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   48
         Top             =   600
         Width           =   4215
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1080
            TabIndex        =   62
            Top             =   720
            Width           =   2775
         End
         Begin VB.CommandButton Command3 
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
            Height          =   375
            Left            =   2280
            Picture         =   "frmAutOC.frx":4856
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   1560
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Indirecta"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   1560
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Internacional"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   1320
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Nacional"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1080
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1080
            TabIndex        =   51
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label19 
            Caption         =   "Comentario:"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "No. Orden :"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1575
         Left            =   120
         TabIndex        =   29
         Top             =   2640
         Width           =   4215
         Begin VB.TextBox txtEnviara 
            Height          =   285
            Left            =   240
            TabIndex        =   31
            Top             =   480
            Width           =   3735
         End
         Begin VB.TextBox txtComentarios 
            Height          =   285
            Left            =   240
            TabIndex        =   30
            Top             =   1080
            Width           =   3735
         End
         Begin VB.Label Label7 
            Caption         =   "ENVIAR  A"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "COMENTARIOS"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.TextBox txtTotal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   28
         Text            =   "0"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtSubtotal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   27
         Text            =   "0"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtImpuesto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Text            =   "0"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtCargos 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   25
         Text            =   "0"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtFlete 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtDescuento 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   23
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton opnNacional 
         Caption         =   "Nacional"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opnInternacional 
         Caption         =   "Internacional"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton opnIndirecta 
         Caption         =   "Indirecta"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtCant 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cambiar"
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
         Left            =   2880
         Picture         =   "frmAutOC.frx":7228
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox txtCantOrg 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   4920
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblMoneda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   45
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label lblFolio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3000
         TabIndex        =   43
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SUBTOTAL"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "IMPUESTO"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "FLETE"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "OTROS CARGOS"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "DESCUENTO"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "FOLIO"
         Height          =   255
         Left            =   3000
         TabIndex        =   36
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "CANTIDAD:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Precios Expresados en:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   4320
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   14
      Top             =   5880
      Width           =   975
      Begin VB.Image Image6 
         Height          =   705
         Left            =   120
         MouseIcon       =   "frmAutOC.frx":9BFA
         MousePointer    =   99  'Custom
         Picture         =   "frmAutOC.frx":9F04
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rechazar"
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
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame21 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   12
      Top             =   4680
      Width           =   975
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Autorizar"
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
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Image cmdGuardar 
         Height          =   705
         Left            =   120
         MouseIcon       =   "frmAutOC.frx":B9B6
         MousePointer    =   99  'Custom
         Picture         =   "frmAutOC.frx":BCC0
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   10
      Top             =   7080
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmAutOC.frx":D772
         MousePointer    =   99  'Custom
         Picture         =   "frmAutOC.frx":DA7C
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   8520
      Picture         =   "frmAutOC.frx":FB5E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7920
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwOCInternacionales 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwOCNacionales 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwOCIndirectas 
      Height          =   135
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   238
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwCotizaciones 
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblidprod 
      Height          =   255
      Left            =   9840
      TabIndex        =   47
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   9840
      TabIndex        =   46
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblID 
      Height          =   255
      Left            =   9840
      TabIndex        =   44
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Nacionales 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Internacionales :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nacionales :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indirectas :"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblIndex 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   7920
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblSelec 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   7920
      Width           =   8055
   End
End
Attribute VB_Name = "frmAutOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim NumOrdenImprime As Integer
Dim IdProveedor As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub CmdGuardar_Click()
    Dim sBuscar As String
    sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'S', COMENTARIO = '" & txtComentarios.Text & "', ENVIARA = '" & txtEnviara.Text & "', ID_USUARIO_AUT = " & VarMen.Text1(0).Text & " WHERE ID_ORDEN_COMPRA = " & lblID.Caption
    cnn.Execute (sBuscar)
    If Hay_Ordenes_Compra Then
        Llenar_Lista_Compras "Internacionales"
        Llenar_Lista_Compras "Nacionales"
        Llenar_Lista_Compras "Indirectas"
    End If
    lvwCotizaciones.ListItems.Clear
    txtSubtotal.Text = "0"
    txtDescuento.Text = "0"
    txtFlete.Text = "0"
    txtCargos.Text = "0"
    txtEnviara.Text = ""
    txtComentarios.Text = ""
    Label10.Caption = ""
    lblSelec.Caption = ""
End Sub
Private Sub Command1_Click()
    If VarMen.Text1(47).Text = "S" Then
        Dim sBuscar As String
        Dim sBuscar2 As String
        Dim Total As Double
        Dim TAX As Double
        Dim Almacen As String
        If Command1.Caption = "Cambiar" Then
            txtCant.Enabled = True
            Command1.Caption = "Guardar"
            cmdGuardar.Enabled = False
            Command2.Enabled = False
            lvwCotizaciones.Enabled = False
            lvwOCIndirectas.Enabled = False
            lvwOCInternacionales.Enabled = False
            lvwOCNacionales.Enabled = False
            txtEnviara.Enabled = False
            txtComentarios.Enabled = False
            txtCantOrg.Text = txtCant.Text
        Else
            txtCant.Enabled = False
            Command1.Caption = "Cambiar"
            cmdGuardar.Enabled = True
            Command2.Enabled = True
            lvwCotizaciones.Enabled = True
            lvwOCIndirectas.Enabled = True
            lvwOCInternacionales.Enabled = True
            lvwOCNacionales.Enabled = True
            txtEnviara.Enabled = True
            txtComentarios.Enabled = True
            sBuscar = "UPDATE ORDEN_COMPRA_DETALLE SET CANTIDAD = " & txtCant.Text & " WHERE ID_ORDEN_COMPRA = " & lblID.Caption & " AND ID_PRODUCTO = '" & lblidprod.Caption & "'"
            cnn.Execute (sBuscar)
            TraeDatos lblID.Caption
            sqlQuery = "UPDATE ORDEN_COMPRA SET TOTAL = " & Replace(txtSubtotal.Text, ",", "") & ",  TAX = " & Replace(txtImpuesto.Text, ",", "") & " WHERE ID_ORDEN_COMPRA = " & lblID.Caption
            cnn.Execute (sqlQuery)
            CANTIDAD = CDbl(txtCantOrg.Text) - CDbl(txtCant.Text)
            sBuscar = "UPDATE COTIZA_REQUI SET CANTIDAD = " & Val(Replace(txtCant.Text, ",", "")) & " WHERE ID_PRODUCTO = '" & lblidprod.Caption & "' AND NUMOC = " & NumOrdenImprime & " AND ID_PROVEEDOR = " & IdProveedor
            cnn.Execute (sBuscar)
            If Val(Replace(txtCantOrg.Text, ",", "")) > Val(Replace(txtCant.Text, ",", "")) Then
                If MsgBox("DESEA GENERAR UNA REQUISISCION POR " & Val(Replace(txtCantOrg.Text, ",", "")) - Val(Replace(txtCant.Text, ",", "")) & " DE " & lblidprod.Caption, vbYesNo, "SACC") = vbYes Then
                    sBuscar = "SELECT ID_REQUISICION FROM REQUISICION WHERE ACTIVO = 0 AND ID_PRODUCTO = '" & lblidprod.Caption & "'"
                    Set tRs = cnn.Execute(sBuscar)
                    If tRs.BOF And tRs.EOF Then
                        sBuscar = "SELECT Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO = '" & lblidprod.Caption & "'"
                        Set tRs = cnn.Execute(sBuscar)
                        Almacen = "A3"
                        If tRs.BOF And tRs.EOF Then
                            sBuscar = "SELECT Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO = '" & lblidprod.Caption & "'"
                            Set tRs = cnn.Execute(sBuscar)
                            Almacen = "A2"
                            If tRs.BOF And tRs.EOF Then
                                sBuscar = "SELECT Descripcion FROM ALMACEN1 WHERE ID_PRODUCTO = '" & lblidprod.Caption & "'"
                                Set tRs = cnn.Execute(sBuscar)
                                Almacen = "A1"
                            End If
                        End If
                        sBuscar = "INSERT INTO REQUISICION (FECHA,ID_PRODUCTO,Descripcion,CANTIDAD,CONTADOR,ALMACEN) Values('" & Format(Date, "dd/mm/yyyy") & "', '" & lblidprod.Caption & "', '" & Label10.Caption & "', " & Val(Replace(txtCantOrg.Text, ",", "")) - Val(Replace(txtCant.Text, ",", "")) & ", 0, '" & Almacen & "')"
                        sBuscar2 = ""
                    Else
                        sBuscar = "UPDATE REQUISICION SET CANTIDAD = CANTIDAD + " & Val(Replace(txtCantOrg.Text, ",", "")) - Val(Replace(txtCant.Text, ",", "")) & "WHERE ID_REQUISICION = " & tRs.Fields("ID_REQUISICION") & " AND ID_PRODUCTO = '" & lblidprod.Caption & "'"
                        sBuscar2 = "INSERT INTO RASTREOREQUI (ID_REQUI, SOLICITO, CANTIDAD, COMENTARIO, FECHA) Values(" & tRs.Fields("ID_REQUISICION") & ", '" & VarMen.Text1(1).Text & "', " & CANTIDAD & ", 'REQUISICION HECHA DESDE LA AUTORIZACION DE LA PRE-ORDEN POR CAMBIO EN LA CANTIDAD DEL PRODUCTO', '" & Format(Date, "dd/mm/yyyy") & "')"
                    End If
                    tRs.Close
                    cnn.Execute (sBuscar)
                    If sBuscar2 = "" Then
                        sBuscar = "SELECT ID_REQUISICION FROM REQUISICION ORDER BY ID_REQUISICION DESC"
                        Set tRs = cnn.Execute(sBuscar)
                        sBuscar2 = "INSERT INTO RASTREOREQUI (ID_REQUI, SOLICITO, CANTIDAD, COMENTARIO, FECHA) Values(" & tRs.Fields("ID_REQUISICION") & ", '" & VarMen.Text1(1).Text & "', " & CANTIDAD & ", 'REQUISICION HECHA DESDE LA AUTORIZACION DE LA PRE-ORDEN POR CAMBIO EN LA CANTIDAD DEL PRODUCTO', '" & Format(Date, "dd/mm/yyyy") & "')"
                        tRs.Close
                    End If
                    cnn.Execute (sBuscar2)
                End If
            End If
            Label10.Caption = ""
        End If
    Else
        MsgBox "No cuenta con permisos para cambiar cantidades en las Ordenes de compra!", vbExclamation, "SACC"
    End If
End Sub
Private Sub Command2_Click()
    If MsgBox("SI ELIMINA EL PRODUCTO SE ENCIARA AL LISTADO DE PENDIENTES POR COTIZAR, ¿DESEA CONTINUAR?" & Chr(13) & "                             DESEA CONTINUAR", vbYesNo, "SACC") = vbYes Then
        Dim sBuscar As String
        sBuscar = "INSERT INTO REQUISICION (FECHA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, CONTADOR, ALMACEN) Values('" & Format(Date, "dd/mm/yyyy") & "', '" & lblidprod.Caption & "', '" & Label10.Caption & "', " & Val(Replace(txtCantOrg.Text, ",", "")) - Val(Replace(txtCant.Text, ",", "")) & ", 0, '" & Almacen & "')"
        cnn.Execute (sBuscar)
        If CANTIDAD <> "" Then
            sBuscar = "INSERT INTO RASTREOREQUI (ID_REQUI, SOLICITO, CANTIDAD, COMENTARIO, FECHA) Values(" & tRs.Fields("ID_REQUISICION") & ", '" & VarMen.Text1(1).Text & "', " & CANTIDAD & ", 'REQUISICION HECHA DESDE LA AUTORIZACION DE LA PRE-ORDEN POR CAMBIO EN LA CANTIDAD DEL PRODUCTO', '" & Format(Date, "dd/mm/yyyy") & "')"
            cnn.Execute (sBuscar)
        End If
        sBuscar = "DELETE ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & lblID.Caption & " AND ID_PRODUCTO = '" & lblidprod.Caption & "'"
        cnn.Execute (sBuscar)
        TraeDatos lblID.Caption
    End If
End Sub
Private Sub Command3_Click()
    If VarMen.Text1(47).Text = "S" Then
        Label19.Visible = True
        Text3.Visible = True
        If Text3 <> "" Then
            If Text1.Text <> "" Then
                If MsgBox("ESTA SEGURO QUE DESEA CANCELAR LA ORDEN DE COMPRA NO. " & Text1.Text, vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
                    Dim sBuscar As String
                    Dim VarTipo As String
                    Dim tRs As ADODB.Recordset
                    If Option1.Value = True Then
                        VarTipo = "N"
                    End If
                    If Option2.Value = True Then
                        VarTipo = "I"
                    End If
                    If Option3.Value = True Then
                        VarTipo = "X"
                    End If
                    sBuscar = "SELECT ID_ORDEN_COMPRA, CONFIRMADA FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text1.Text & " AND TIPO = '" & VarTipo & "'"
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        If tRs.Fields("CONFIRMADA") = "Y" Then
                            MsgBox "Esta orden ya tiene un pago registrado, imposible cancelarla!", vbExclamation, "SACC"
                            Text1.Text = ""
                            Text3.Text = ""
                        Else
                            sBuscar = "SELECT SUM(SURTIDO) AS ENTRADA, ID_ORDEN_COMPRA FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs.Fields("ID_ORDEN_COMPRA") & " GROUP BY ID_ORDEN_COMPRA"
                            Set tRs = cnn.Execute(sBuscar)
                            If Not (tRs.EOF And tRs.BOF) Then
                                If tRs.Fields("ENTRADA") = 0 Then
                                    sBuscar = "INSERT INTO ORDEN_CANCE (ORDEN, TIPO, COMENTARIO, FECHA, ID_USUARIO) VALUES (" & Text1.Text & ", '" & VarTipo & "','" & Text3.Text & "','" & Format(Date, "dd/mm/yyyy") & "','" & VarMen.Text1(0) & "');"
                                    cnn.Execute (sBuscar)
                                    sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'E' WHERE NUM_ORDEN = " & Text1.Text & " AND TIPO = '" & VarTipo & "'"
                                    cnn.Execute (sBuscar)
                                    'sBuscar = "DELETE FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text1.Text & " AND TIPO = '" & VarTipo & "'"
                                    'cnn.Execute (sBuscar)
                                    'sBuscar = "DELETE FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs.Fields("ID_ORDEN_COMPRA")
                                    'cnn.Execute (sBuscar)
                                    MsgBox "ORDEN DE COMPRA ELIMINADA", vbInformation, "SACC"
                                Else
                                    MsgBox "LA ORDEN TIENE ENTRADA, IMPOSIBLE DE CANCELAR!", vbExclamation, "SACC"
                                    Text1.Text = ""
                                    Text3.Text = ""
                                End If
                            End If
                        End If
                    Else
                        MsgBox "LA ORDEN NO EXISTE O FUE CANCELADA", vbInformation, "SACC"
                        Text1.Text = ""
                        Text3.Text = ""
                    End If
                End If
            End If
        Else
            MsgBox "INGRESAR  MOTIVO DE CANCELACION,COMENTARIO", vbInformation, "SACC"
        End If
    Else
        MsgBox "No cuenta con permisos para cancelar Ordenes de compra!", vbExclamation, "SACC"
    End If
End Sub
Private Sub Command4_Click()
    If VarMen.Text1(78).Text = "S" Then
        If Text2.Text <> "" Then
            If MsgBox("ESTA SEGURO QUE DESEA REGRESAR LA ORDEN DE COMPRA NO. " & Text2.Text, vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
                Dim sBuscar As String
                Dim VarTipo As String
                Dim tRs As ADODB.Recordset
                Dim sMsg As String
                If Option4.Value = True Then
                    VarTipo = "N"
                End If
                If Option5.Value = True Then
                    VarTipo = "I"
                End If
                If Option6.Value = True Then
                    VarTipo = "X"
                End If
                sBuscar = "SELECT CONFIRMADA FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = '" & VarTipo & "'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    If tRs.Fields("CONFIRMADA") <> "Y" Then
                        sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'P' WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = '" & VarTipo & "'"
                        cnn.Execute (sBuscar)
                        If Hay_Ordenes_Compra Then
                            Llenar_Lista_Compras "Internacionales"
                            Llenar_Lista_Compras "Nacionales"
                            Llenar_Lista_Compras "Indirectas"
                        End If
                    Else
                        sBuscar = "SELECT NUMCHEQUE, FECHA FROM ABONOS_PAGO_OC WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = '" & VarTipo & "'"
                        Set tRs = cnn.Execute(sBuscar)
                        If Not (tRs.EOF And tRs.BOF) Then
                            sMsg = "La orden fue pagada con el cheque " & tRs.Fields("NUMCHEQUE") & " el dia " & tRs.Fields("FECHA") & ", si la regresa se anulara este cheque, ¿Está seguro que desea regresarla?"
                        Else
                            sMsg = "La orden ya esta pagada, si la regresa se anulara el pago anterior, ¿Está seguro que desea regresarla?"
                        End If
                        If MsgBox(sMsg, vbYesNo, "SACC") = vbYes Then
                            sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'P' WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = '" & VarTipo & "'"
                            cnn.Execute (sBuscar)
                            If Hay_Ordenes_Compra Then
                                Llenar_Lista_Compras "Internacionales"
                                Llenar_Lista_Compras "Nacionales"
                                Llenar_Lista_Compras "Indirectas"
                            End If
                        End If
                    End If
                Else
                    MsgBox "Orden no encontrada, posiblemente no ha sido creada o fue eliminada!", vbExclamation, "SACC"
                End If
            End If
        End If
    Else
        MsgBox "No cuenta con permisos para reactivar ordenes de compra cerradas!", vbExclamation, "SACC"
    End If
End Sub
Private Sub Image4_Click()
On Error GoTo ManejaError
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Me.CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
    Me.CommonDialog1.ShowSave
    Ruta = Me.CommonDialog1.FileName
    If Ruta <> "" Then
        If opnNacional.Value = True Then
            NumColum = lvwOCNacionales.ColumnHeaders.Count
            For Con = 1 To lvwOCNacionales.ColumnHeaders.Count
                StrCopi = StrCopi & lvwOCNacionales.ColumnHeaders(Con).Text & Chr(9)
            Next
        Else
            If opnInternacional.Value = True Then
                NumColum = lvwOCInternacionales.ColumnHeaders.Count
                For Con = 1 To lvwOCInternacionales.ColumnHeaders.Count
                    StrCopi = StrCopi & lvwOCInternacionales.ColumnHeaders(Con).Text & Chr(9)
                Next
            Else
                If opnIndirecta.Value = True Then
                    NumColum = lvwOCIndirectas.ColumnHeaders.Count
                    For Con = 1 To lvwOCIndirectas.ColumnHeaders.Count
                        StrCopi = StrCopi & lvwOCIndirectas.ColumnHeaders(Con).Text & Chr(9)
                    Next
                End If
            End If
        End If
        StrCopi = StrCopi & Chr(13)
        If opnNacional.Value = True Then
            For Con = 1 To lvwOCNacionales.ListItems.Count
                StrCopi = StrCopi & lvwOCNacionales.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & lvwOCNacionales.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
        Else
            If opnInternacional.Value = True Then
                For Con = 1 To lvwOCInternacionales.ListItems.Count
                    StrCopi = StrCopi & lvwOCInternacionales.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & lvwOCInternacionales.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                Next
            Else
                If opnIndirecta.Value = True Then
                    For Con = 1 To lvwOCIndirectas.ListItems.Count
                        StrCopi = StrCopi & lvwOCIndirectas.ListItems.Item(Con) & Chr(9)
                        For Con2 = 1 To NumColum - 1
                            StrCopi = StrCopi & lvwOCIndirectas.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                        Next
                        StrCopi = StrCopi & Chr(13)
                    Next
                End If
            End If
        End If
        'archivo TXT
        Dim foo As Integer
        foo = FreeFile
        Open Ruta For Output As #foo
            Print #foo, StrCopi
        Close #foo
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub Image6_Click()
    If VarMen.Text1(8).Text = "S" Then
        If lblID.Caption <> "" Then
            Dim sBuscar As String
            Dim iAfectados As Long
            Dim tRs As ADODB.Recordset
            Dim IdProveedor As String
            Dim fecha As String
            Dim Moneda As String
            sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'N', COMENTARIO = '" & txtComentarios.Text & "', ENVIARA = '" & txtEnviara.Text & "' WHERE ID_ORDEN_COMPRA = " & lblID.Caption
            cnn.Execute (sBuscar)
            sBuscar = "UPDATE COTIZA_REQUI SET ESTADO_ACTUAL = 'X' WHERE NUMOC = " & lblID.Caption
            Set tRs = cnn.Execute(sBuscar, iAfectados, adCmdText)
            If iAfectados < 1 Then
                sBuscar = "SELECT ID_PROVEEDOR, FECHA, MONEDA FROM ORDEN_COMPRA WHERE ID_ORDEN_COMPRA = " & lblID.Caption
                Set tRs = cnn.Execute(sBuscar)
                IdProveedor = tRs.Fields("ID_PROVEEDOR")
                fecha = tRs.Fields("FECHA")
                Moneda = tRs.Fields("MONEDA")
                sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO, DIAS_ENTREGA FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & lblID.Caption
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    Do While Not tRs.EOF
                        sBuscar = "INSERT INTO COTIZA_REQUI (ID_REQUISICION, ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, CANTIDAD, DIAS_ENTREGA, PRECIO, FECHA, ESTADO_ACTUAL, NUMOC, MONEDA) VALUES (99999, " & IdProveedor & ", '" & tRs.Fields("ID_PRODUCTO") & "', '" & tRs.Fields("Descripcion") & "', '" & tRs.Fields("CANTIDAD") & "', '" & tRs.Fields("DIAS_ENTREGA") & "', '" & tRs.Fields("PRECIO") & "', '" & fecha & "', 'X', '" & lblID.Caption & "', '" & Moneda & "');"
                        cnn.Execute (sBuscar)
                        tRs.MoveNext
                    Loop
                End If
            End If
            If Hay_Ordenes_Compra Then
                Llenar_Lista_Compras "Internacionales"
                Llenar_Lista_Compras "Nacionales"
                Llenar_Lista_Compras "Indirectas"
            Else
                lvwOCInternacionales.ListItems.Clear
                lvwOCNacionales.ListItems.Clear
                lvwOCIndirectas.ListItems.Clear
            End If
            lvwCotizaciones.ListItems.Clear
            txtSubtotal.Text = "0"
            txtDescuento.Text = "0"
            txtFlete.Text = "0"
            txtCargos.Text = "0"
            txtEnviara.Text = ""
            txtImpuesto.Text = "0"
            txtComentarios.Text = ""
            Label10.Caption = ""
            lblSelec.Caption = ""
            txtTotal.Text = "0"
        End If
    Else
        MsgBox "No cuenta con permisos de rechazar Ordenes de compra!", vbExclamation, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwCotizaciones_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwCotizaciones.ListItems.Count > 0 Then
        lblIndex.Caption = Item.Index
        lblSelec.Caption = Item.SubItems(1) & " Cantidad: " & Item.SubItems(2)
        txtCant.Text = Item.SubItems(2)
        Label10.Caption = Item.SubItems(1)
        lblidprod.Caption = Item
    End If
End Sub
Private Sub lvwOCIndirectas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    NumOrdenImprime = Item
    IdProveedor = Item.SubItems(1)
End Sub
Private Sub lvwOCInternacionales_ItemClick(ByVal Item As MSComctlLib.ListItem)
    NumOrdenImprime = Item
    IdProveedor = Item.SubItems(1)
End Sub
Private Sub lvwOCNacionales_ItemClick(ByVal Item As MSComctlLib.ListItem)
    NumOrdenImprime = Item
    IdProveedor = Item.SubItems(1)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtCant_GotFocus()
    txtCant.BackColor = &HFFE1E1
End Sub
Private Sub txtCant_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCant.Text <> "" Then
        Command1.Value = True
    End If
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtCant_LostFocus()
    txtCant.BackColor = &H80000005
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With Me.lvwOCInternacionales
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID_ORDEN", 0
        .ColumnHeaders.Add , , "No. ORDEN", 1000
        .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
        .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
        .ColumnHeaders.Add , , "COMENTARIO", 1440
    End With
    With Me.lvwOCNacionales
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID_ORDEN", 0
        .ColumnHeaders.Add , , "No. ORDEN", 1000
        .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
        .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
        .ColumnHeaders.Add , , "COMENTARIO", 1440
    End With
    With Me.lvwOCIndirectas
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID_ORDEN", 0
        .ColumnHeaders.Add , , "No. ORDEN", 1000
        .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
        .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
        .ColumnHeaders.Add , , "COMENTARIO", 1440
    End With
    With Me.lvwCotizaciones
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripcion", 3500
        .ColumnHeaders.Add , , "CANTIDAD", 1000, 2
        .ColumnHeaders.Add , , "PRECIO", 1440, 2
        .ColumnHeaders.Add , , "SUBTOTAL", 1440, 2
    End With
    If Hay_Ordenes_Compra Then
        Llenar_Lista_Compras "Internacionales"
        Llenar_Lista_Compras "Nacionales"
        Llenar_Lista_Compras "Indirectas"
    End If
    If NvoMen.Text1(11).Text = "N" Then
        Command4.Enabled = False
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Public Function Hay_Ordenes_Compra() As Boolean
On Error GoTo ManejaError
    Dim sBuscar  As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT  COUNT(ID_ORDEN_COMPRA) AS CONTA FROM ORDEN_COMPRA WHERE CONFIRMADA = 'P' OR Confirmada = 'S'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Hay_Ordenes_Compra = True
    Else
        Hay_Ordenes_Compra = False
    End If
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Lista_Compras(Tipo As String)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT OC.Id_Orden_Compra,OC.NUM_ORDEN,OC.Id_Proveedor,P.Nombre,((OC.Total - OC.Discount) + OC.TAX + OC.Freight + OC.Otros_Cargos) AS Total_Pagar,OC.COMENTARIO FROM ORDEN_COMPRA AS OC JOIN PROVEEDOR AS P ON P.Id_Proveedor = OC.Id_Proveedor WHERE OC.Confirmada = 'P' AND OC.Tipo = '"
    Select Case Tipo
        Case "Internacionales":
            Me.lvwOCInternacionales.ListItems.Clear
            sBuscar = sBuscar & "I' ORDER BY NUM_ORDEN"
            Set tRs = cnn.Execute(sBuscar)
            With tRs
                While Not .EOF
                    Set ItMx = Me.lvwOCInternacionales.ListItems.Add(, , .Fields("ID_ORDEN_COMPRA"))
                    If Not IsNull(.Fields("NUM_ORDEN")) Then ItMx.SubItems(1) = Trim(.Fields("NUM_ORDEN"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(2) = Trim(.Fields("ID_PROVEEDOR"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(3) = Trim(.Fields("Nombre"))
                    If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(.Fields("Total_Pagar"))
                    If Not IsNull(.Fields("COMENTARIO")) Then ItMx.SubItems(5) = Trim(.Fields("COMENTARIO"))
                    .MoveNext
                Wend
            .Close
            End With
        Case "Nacionales":
            Me.lvwOCNacionales.ListItems.Clear
            sBuscar = sBuscar & "N' ORDER BY NUM_ORDEN"
            Set tRs = cnn.Execute(sBuscar)
            With tRs
                While Not .EOF
                    Set ItMx = Me.lvwOCNacionales.ListItems.Add(, , .Fields("ID_ORDEN_COMPRA"))
                    If Not IsNull(.Fields("NUM_ORDEN")) Then ItMx.SubItems(1) = Trim(.Fields("NUM_ORDEN"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(2) = Trim(.Fields("ID_PROVEEDOR"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(3) = Trim(.Fields("Nombre"))
                    If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(.Fields("Total_Pagar"))
                    If Not IsNull(.Fields("COMENTARIO")) Then ItMx.SubItems(5) = Trim(.Fields("COMENTARIO"))
                    .MoveNext
                Wend
            .Close
            End With
        Case "Indirectas":
            Me.lvwOCIndirectas.ListItems.Clear
            sBuscar = sBuscar & "X' ORDER BY NUM_ORDEN"
            Set tRs = cnn.Execute(sBuscar)
            With tRs
                While Not .EOF
                    Set ItMx = Me.lvwOCIndirectas.ListItems.Add(, , .Fields("ID_ORDEN_COMPRA"))
                    If Not IsNull(.Fields("NUM_ORDEN")) Then ItMx.SubItems(1) = Trim(.Fields("NUM_ORDEN"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(2) = Trim(.Fields("ID_PROVEEDOR"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(3) = Trim(.Fields("Nombre"))
                    If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(.Fields("Total_Pagar"))
                    If Not IsNull(.Fields("COMENTARIO")) Then ItMx.SubItems(5) = Trim(.Fields("COMENTARIO"))
                    .MoveNext
                Wend
            .Close
            End With
        Case Else:
            MsgBox "ERROR GRAVE. LA APLICACIÓN TERMINARA", vbCritical, "SACC"
            End
    End Select
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub lvwOCIndirectas_Click()
On Error GoTo ManejaError
    If lvwOCIndirectas.ListItems.Count > 0 Then
        TraeDatos lvwOCIndirectas.SelectedItem
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub lvwOCInternacionales_Click()
On Error GoTo ManejaError
    If lvwOCInternacionales.ListItems.Count > 0 Then
        TraeDatos lvwOCInternacionales.SelectedItem
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub lvwOCNacionales_Click()
On Error GoTo ManejaError
    If lvwOCNacionales.ListItems.Count > 0 Then
        TraeDatos lvwOCNacionales.SelectedItem
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub TraeDatos(NO As Integer)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tLi As ListItem
    Dim Subtotal As Double
    sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE ID_ORDEN_COMPRA = " & NO
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If .EOF And .BOF Then
            MsgBox "FALLA GRAVE DE INFORMACION, LLAME A SOPORTE", vbCritical, "SACC"
        Else
            lblID.Caption = .Fields("ID_ORDEN_COMPRA")
            lblFolio.Caption = .Fields("NUM_ORDEN")
            If .Fields("TIPO") = "I" Then
                opnInternacional.Value = True
                lblMoneda.Caption = "DOLARES"
            ElseIf .Fields("TIPO") = "N" Then
                opnIndirecta.Value = True
            Else
                opnIndirecta.Value = True
            End If
            lblID.Caption = NO
            lblMoneda.Caption = ""
            If Not IsNull(.Fields("DISCOUNT")) Then txtDescuento.Text = .Fields("DISCOUNT")
            If Not IsNull(.Fields("TAX")) Then txtImpuesto.Text = .Fields("TAX")
            If Not IsNull(.Fields("FREIGHT")) Then txtFlete.Text = .Fields("FREIGHT")
            If Not IsNull(.Fields("OTROS_CARGOS")) Then txtCargos.Text = .Fields("OTROS_CARGOS")
            If Not IsNull(.Fields("ENVIARA")) Then txtEnviara.Text = .Fields("ENVIARA")
            If Not IsNull(.Fields("COMENTARIO")) Then txtComentarios.Text = .Fields("COMENTARIO")
            If Not IsNull(.Fields("MONEDA")) Then lblMoneda.Caption = .Fields("MONEDA")
            Subtotal = 0
            lvwCotizaciones.ListItems.Clear
            sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & NO
            Set tRs2 = cnn.Execute(sBuscar)
            With tRs2
                If Not (.EOF And .BOF) Then
                    Do While Not .EOF
                        Set tLi = lvwCotizaciones.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                        If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = Trim(.Fields("Descripcion"))
                        If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = Trim(.Fields("CANTIDAD"))
                        If Not IsNull(.Fields("PRECIO")) Then tLi.SubItems(3) = Trim(.Fields("PRECIO"))
                        tLi.SubItems(4) = CDbl(.Fields("PRECIO")) * CDbl(.Fields("CANTIDAD"))
                        Subtotal = Subtotal + (CDbl(.Fields("PRECIO")) * CDbl(.Fields("CANTIDAD")))
                        txtTotal.Text = (CDbl(Replace(Subtotal, ",", "")) - CDbl(Replace(txtDescuento.Text, ",", ""))) + CDbl(Replace(txtImpuesto.Text, ",", "")) + CDbl(Replace(txtFlete.Text, ",", "")) + CDbl(Replace(txtCargos.Text, ",", ""))
                        .MoveNext
                    Loop
                End If
            End With
            If Not IsNull(.Fields("TOTAL")) Then txtSubtotal.Text = .Fields("TOTAL")
            txtSubtotal.Text = Subtotal
        End If
    End With
End Sub
Sub Llenar_Lista_Cotizaciones(NO As Integer)
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim CONT2 As Integer
    
    sqlQuery = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & NO
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        Me.lvwCotizaciones.ListItems.Clear
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvwCotizaciones.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = Trim(.Fields("Descripcion"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = Trim(.Fields("CANTIDAD"))
                If Not IsNull(.Fields("PRECIO")) Then tLi.SubItems(3) = Trim(.Fields("PRECIO"))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub txtSubtotal_Change()
On Error GoTo ManejaError
    'If opnInternacional.Value = False Then
    '    txtImpuesto.Text = ((Val(Replace(txtSubtotal.Text, ",", "")) - Val(Replace(txtDescuento.Text, ",", ""))) + txtFlete) * CDbl(CDbl(VarMen.Text4(7).Text) / 100)
    'Else
    '    txtImpuesto.Text = Val("###,###,##0.00")
    'End If
    'txtTotal.Text = (Val(Replace(txtSubtotal.Text, ",", "")) - Val(Replace(txtDescuento.Text, ",", ""))) + Val(Replace(txtImpuesto.Text, ",", "")) + Val(Replace(txtFlete.Text, ",", "")) + Val(Replace(txtCargos.Text, ",", ""))
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
