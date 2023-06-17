VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form VentasEsp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Venta  Especial"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12150
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   12150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text13 
      Height          =   1815
      Left            =   10680
      MultiLine       =   -1  'True
      TabIndex        =   60
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Forma de pago"
      Height          =   2055
      Left            =   10680
      TabIndex        =   44
      Top             =   3840
      Width           =   1335
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No Aplica"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "T. Debito"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Efectivo"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cheque"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "T. Credito"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "T. Electrón"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView ListView4 
      Height          =   255
      Left            =   11880
      TabIndex        =   43
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Comandas"
      Height          =   1455
      Left            =   10680
      TabIndex        =   40
      Top             =   2400
      Width           =   1335
      Begin VB.CommandButton Command2 
         Caption         =   "Extraer"
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
         Left            =   120
         Picture         =   "VentasEsp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10800
      TabIndex        =   38
      Top             =   7080
      Width           =   975
      Begin VB.Label Label34 
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
         TabIndex        =   39
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "VentasEsp.frx":29D2
         MousePointer    =   99  'Custom
         Picture         =   "VentasEsp.frx":2CDC
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10800
      TabIndex        =   36
      Top             =   5880
      Width           =   975
      Begin VB.Image Command1 
         Height          =   720
         Left            =   120
         MouseIcon       =   "VentasEsp.frx":4DBE
         MousePointer    =   99  'Custom
         Picture         =   "VentasEsp.frx":50C8
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label16 
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
         TabIndex        =   37
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "VentasEsp.frx":6A8A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label12"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label17"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label18"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label19"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ListView1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ListView2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ListView3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Check1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text12"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text11"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text2(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text3"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Option1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Option2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text2(11)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text6"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Command4"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Frame1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text7"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text8"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text9"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text10"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      Begin VB.TextBox Text10 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "0"
         Top             =   7320
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3840
         TabIndex        =   56
         Text            =   "0"
         Top             =   7680
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "0"
         Top             =   7320
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   9720
         TabIndex        =   51
         Top             =   240
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Producto Seleccionado"
         Height          =   1095
         Left            =   120
         TabIndex        =   18
         Top             =   4680
         Width           =   10215
         Begin VB.CommandButton Command3 
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
            Left            =   8640
            Picture         =   "VentasEsp.frx":6AA6
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   9
            Left            =   1440
            TabIndex        =   4
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   8
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   7
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   5895
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   5
            Top             =   600
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   5640
            TabIndex        =   21
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   49938433
            CurrentDate     =   38681
         End
         Begin VB.Label Label7 
            Caption         =   "Clave"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Descripcion"
            Height          =   255
            Left            =   3120
            TabIndex        =   25
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Precio de Venta"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha"
            Height          =   255
            Left            =   5040
            TabIndex        =   23
            Top             =   600
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   2880
            TabIndex        =   22
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.CommandButton Command4 
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
         Left            =   9240
         Picture         =   "VentasEsp.frx":9478
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7680
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   10200
         TabIndex        =   17
         Top             =   7320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   1200
         TabIndex        =   16
         Text            =   "0"
         Top             =   7320
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   6960
         TabIndex        =   15
         Top             =   2760
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   6960
         TabIndex        =   14
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   2880
         Width           =   5415
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   0
         Top             =   960
         Width           =   6735
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   7680
         Width           =   1455
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   6600
         TabIndex        =   11
         Text            =   "0"
         Top             =   7680
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CREDITO"
         Height          =   255
         Left            =   9240
         TabIndex        =   9
         Top             =   7320
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1335
         Left            =   120
         TabIndex        =   7
         Top             =   5880
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   2355
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   3240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   2355
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   2355
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
      Begin VB.Label Label19 
         Caption         =   "RETENCIÓN"
         Height          =   255
         Left            =   5400
         TabIndex        =   57
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "IMPUESTO 2"
         Height          =   255
         Left            =   2760
         TabIndex        =   55
         Top             =   7680
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "IMPUESTO 1"
         Height          =   255
         Left            =   2760
         TabIndex        =   54
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "% IVA"
         Height          =   255
         Left            =   9240
         TabIndex        =   52
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "SUBTOTAL"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Buscar producto"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Sucursal :"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label8 
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
         Left            =   960
         TabIndex        =   31
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label10 
         Caption         =   "Agente"
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label13 
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
         Left            =   5400
         TabIndex        =   29
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label Label14 
         Caption         =   "IVA"
         Height          =   255
         Left            =   720
         TabIndex        =   28
         Top             =   7680
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "TOTAL"
         Height          =   255
         Left            =   5880
         TabIndex        =   27
         Top             =   7680
         Width           =   615
      End
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   11880
      Top             =   6360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Comentarios:"
      Height          =   255
      Left            =   10680
      TabIndex        =   59
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "VentasEsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Elim As Double
Private elim2 As Double
Private tot As Double
Private Xind As Integer
Dim IdClien As String
Dim NomClien As String
Dim DesClien As Integer
Dim DiasCred As String
Dim LimCred As String
Dim CveVenta As String
Dim Porci As Double
Dim valor As Double
Dim IdProdEli As String
Dim CantCom As Integer
Dim ContComanda As Integer
Dim SinExis As Integer
Dim RFC As String
Dim VarAlmacen As String
Dim IVA As String
Dim IMPUESTO1 As String
Dim IMPUESTO2 As String
Dim RETENCION As String
Dim IVAEli As String
Dim Impu1Eli As String
Dim Impu2Eli As String
Dim RetEli As String
Private Sub Command1_Click()
On Error GoTo ManejaError
    If ListView3.ListItems.Count <> 0 And IdClien <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tRs1 As ADODB.Recordset
        Dim tRs8 As ADODB.Recordset
        Dim NumReg As Integer
        Dim Cont As Integer
        Dim AbonClien As String
        Dim DeudClien As String
        Dim CredDispClien As String
        Dim FormaPago As String
        Dim FormaPagoSAT As String
        sBuscar = "SELECT ID_CLIENTE, SUCURSAL, TOTAL FROM VENTAS WHERE ID_VENTA IN (SELECT TOP 1 ID_VENTA FROM VENTAS ORDER BY ID_VENTA DESC) AND ID_CLIENTE = " & IdClien & " AND SUCURSAL = '" & VarMen.Text4(0).Text & "' AND TOTAL = '" & Replace(Text12.Text, ",", "") & "' ORDER BY ID_VENTA DESC"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            If MsgBox("EL FOLIO DE VENTA ANTERIOR ES IGUAL A LA VENTA QUE SE INTENTA CAPTURAR, ¿DESEA REALIZAR UNA VENTA NUEVA?", vbYesNo, "SACC") = vbNo Then
                Exit Sub
            End If
        End If
        If Option3.Value = True Then
            FormaPago = "C"
            FormaPagoSAT = "001"
        End If
        If Option4.Value = True Then
            FormaPago = "H"
            FormaPagoSAT = "002"
        End If
        If Option5.Value = True Then
            FormaPago = "T"
            FormaPagoSAT = "004"
        End If
        If Option6.Value = True Then
            FormaPago = "E"
            FormaPagoSAT = "003"
        End If
        If Option12.Value = True Then
            FormaPago = "D"
            FormaPagoSAT = "028"
        End If
        If Option7.Value = True Then
            FormaPago = "N"
            FormaPagoSAT = "099"
        End If
        If Option3.Value = False And Option4.Value = False And Option5.Value = False And Option6.Value = False And Option12.Value = False And Option7.Value = False Then
            MsgBox "DEBE MARCAR UNA FORMA DE PAGO!", vbExclamation, "SACC"
            Exit Sub
        End If
        If Check1.Value = 1 Then
            sBuscar = "SELECT SUM(TOTAL_COMPRA) AS COMPRA FROM CUENTAS WHERE ID_CLIENTE = " & IdClien
            Set tRs = cnn.Execute(sBuscar)
            sBuscar = "SELECT SUM(CANT_ABONO) AS ABONO FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & IdClien
            Set tRs1 = cnn.Execute(sBuscar)
            If Not IsNull(tRs1.Fields("ABONO")) Then
                AbonClien = tRs1.Fields("ABONO")
            Else
                AbonClien = "0"
            End If
            If Not IsNull(tRs.Fields("COMPRA")) Then
                DeudClien = tRs.Fields("COMPRA")
            Else
                DeudClien = "0"
            End If
            CredDispClien = CDbl(LimCred) - CDbl(DeudClien) + CDbl(AbonClien)
        Else
            If MsgBox("Desea hacer la venta a credito?", vbYesNo, "SACC") = vbYes Then
                sBuscar = "SELECT SUM(TOTAL_COMPRA) AS COMPRA FROM CUENTAS WHERE ID_CLIENTE = " & IdClien
                Set tRs = cnn.Execute(sBuscar)
                sBuscar = "SELECT SUM(CANT_ABONO) AS ABONO FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & IdClien
                Set tRs1 = cnn.Execute(sBuscar)
                If Not IsNull(tRs1.Fields("ABONO")) Then
                    AbonClien = tRs1.Fields("ABONO")
                Else
                    AbonClien = "0"
                End If
                If Not IsNull(tRs.Fields("COMPRA")) Then
                    DeudClien = tRs.Fields("COMPRA")
                Else
                    DeudClien = "0"
                End If
                CredDispClien = CDbl(LimCred) - CDbl(DeudClien) + CDbl(AbonClien)
            Else
                CredDispClien = 0
            End If
        End If
        If CDbl(CredDispClien) > CDbl(Text12.Text) Or Check1.Value = 0 Then
            If DiasCred = "" Then
                DiasCred = "0"
            End If
            If Check1.Value = 1 Then
                sBuscar = "INSERT INTO VENTAS (ID_CLIENTE, NOMBRE, IVA, SUBTOTAL, TOTAL, DESCUENTO, ID_USUARIO, FECHA, SUCURSAL, DIAS_CREDITO, UNA_EXIBICION, TIPO_PAGO, FORMA_PAGO, IMPUESTO1, IMPUESTO2, RETENCION, COMENTARIO, FormaPagoSAT) VALUES (" & IdClien & ", '" & NomClien & "', " & Replace(Text11.Text, ",", "") & ", " & Replace(Text2(11).Text, ",", "") & ", " & Replace(Text12.Text, ",", "") & ", " & Replace(DesClien, ",", "") & ", '" & VarMen.Text1(0).Text & "',  DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), '" & VarMen.Text4(0).Text & "', " & DiasCred & ", 'N', '" & FormaPago & "', 'PAGO EN PARCIALIDADES', '" & Text8.Text & "', '" & Text9.Text & "', '" & Text10.Text & "', '" & Text13.Text & "', '" & FormaPagoSAT & "');"
            Else
                sBuscar = "INSERT INTO VENTAS (ID_CLIENTE, NOMBRE, IVA, SUBTOTAL, TOTAL, DESCUENTO, ID_USUARIO, FECHA, SUCURSAL, DIAS_CREDITO, UNA_EXIBICION, TIPO_PAGO, FORMA_PAGO, IMPUESTO1, IMPUESTO2, RETENCION, COMENTARIO, FormaPagoSAT) VALUES (" & IdClien & ", '" & NomClien & "', " & Replace(Text11.Text, ",", "") & ", " & Replace(Text2(11).Text, ",", "") & ", " & Replace(Text12.Text, ",", "") & ", " & Replace(DesClien, ",", "") & ", '" & VarMen.Text1(0).Text & "',  DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), '" & VarMen.Text4(0).Text & "', " & DiasCred & ", 'S', '" & FormaPago & "', 'PAGO EN UNA EXHIBICION', '" & Text8.Text & "', '" & Text9.Text & "', '" & Text10.Text & "', '" & Text13.Text & "', '" & FormaPagoSAT & "');"
            End If
            cnn.Execute (sBuscar)
            sBuscar = "SELECT TOP 1 ID_VENTA FROM VENTAS ORDER BY ID_VENTA DESC"
            Set tRs = cnn.Execute(sBuscar)
            CveVenta = tRs.Fields("ID_VENTA")
            NumReg = ListView3.ListItems.Count
            If Check1.Value = 1 Then
                sBuscar = "INSERT INTO CUENTAS (PAGADA, ID_CLIENTE, ID_USUARIO, FECHA, DIAS_CREDITO, FECHA_VENCE, DESCUENTO, SUCURSAL, TOTAL_COMPRA, DEUDA, ID_VENTA) VALUES ( 'N', " & IdClien & ", '" & VarMen.Text1(0).Text & "',  DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), " & DiasCred & ", '" & Format(Date + DiasCred, "dd/mm/yyyy") & "', " & DesClien & ", '" & VarMen.Text4(0).Text & "', " & Replace(Text12.Text, ",", "") & ", " & Replace(Text12.Text, ",", "") & ", " & tRs.Fields("ID_VENTA") & ");"
                cnn.Execute (sBuscar)
                sBuscar = "SELECT TOP 1 ID_CUENTA FROM CUENTAS ORDER BY ID_CUENTA DESC"
                Set tRs1 = cnn.Execute(sBuscar)
                sBuscar = "INSERT INTO CUENTA_VENTA (ID_VENTA, ID_CUENTA) VALUES (" & tRs.Fields("ID_VENTA") & ", " & tRs1.Fields("ID_CUENTA") & ");"
                cnn.Execute (sBuscar)
            End If
            For Cont = 1 To NumReg
                sBuscar = "INSERT INTO VENTAS_DETALLE (ID_VENTA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_VENTA, PRECIO_COSTO, GANANCIA, IMPORTE, IVA, IMPUESTO1, IMPUESTO2, RETENCION) VALUES (" & tRs.Fields("ID_VENTA") & ", '" & ListView3.ListItems(Cont) & "', '" & ListView3.ListItems(Cont).SubItems(1) & "', " & Replace(ListView3.ListItems(Cont).SubItems(5), ",", "") & ", " & Replace(ListView3.ListItems(Cont).SubItems(3), ",", "") & ", " & Replace(ListView3.ListItems(Cont).SubItems(4), ",", "") & ", " & Replace(ListView3.ListItems(Cont).SubItems(2), ",", "") & ", " & CDbl(ListView3.ListItems(Cont).SubItems(5)) * CDbl(ListView3.ListItems(Cont).SubItems(3)) & ", " & ListView3.ListItems(Cont).SubItems(6) & ", " & ListView3.ListItems(Cont).SubItems(7) & ", " & ListView3.ListItems(Cont).SubItems(8) & ", " & ListView3.ListItems(Cont).SubItems(9) & ");"
                cnn.Execute (sBuscar)
                If Check1.Value = 1 Then
                    sBuscar = "INSERT INTO CUENTA_DETALLE (ID_CUENTA, CANTIDAD, ID_PRODUCTO, PRECIO_VENTA) VALUES (" & tRs1.Fields("ID_CUENTA") & ", " & Replace(ListView3.ListItems(Cont).SubItems(5), ",", "") & ", '" & ListView3.ListItems(Cont) & "', " & Replace(ListView3.ListItems(Cont).SubItems(3), ",", "") & ");"
                    cnn.Execute (sBuscar)
                End If
            Next Cont
            Frame8.Visible = False
            ImprTICKET
            FunRemision
            Unload Me
        Else
            MsgBox "LA COMPRA SUPERA EL LIMITE DE CREDITO DISPONIBLE!($" & CredDispClien & " DISPONIBLE)", vbInformation, "SACC"
        End If
    Else
        MsgBox "FALTA INFORMACION NECESARIA PARA GUARDA LA VENTA!", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView4.ListItems.Clear
    sBuscar = "SELECT ID_PRODUCTO, CANTIDAD, CANTIDAD_NO_SIRVIO FROM COMANDAS_DETALLES_2 WHERE ID_COMANDA = " & Text4.Text & " AND ESTADO_ACTUAL IN ('L','N') AND (CANTIDAD - CANTIDAD_NO_SIRVIO) > 0"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView4.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            tLi.SubItems(1) = tRs.Fields("CANTIDAD") - tRs.Fields("CANTIDAD_NO_SIRVIO")
            tLi.SubItems(2) = Text4.Text
            tRs.MoveNext
        Loop
    End If
    ContComanda = 0
    CantCom = ListView4.ListItems.Count
    ExtraerComanda
End Sub
Private Sub ExtraerComanda()
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    ContComanda = ContComanda + 1
    If CantCom >= ContComanda Then
        SinExis = 1
        Text2(8).Text = ListView4.ListItems(ContComanda)
        Text5.Text = ListView4.ListItems(ContComanda).SubItems(1)
        sBuscar = "SELECT Descripcion, PRECIO_COSTO, GANANCIA, P_RETENCION, IVA, IMPUESTO1, IMPUESTO2 FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView4.ListItems(ContComanda) & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Text2(7).Text = tRs.Fields("Descripcion")
            Text2(9).Text = Format(CDbl(tRs.Fields("PRECIO_COSTO")) * CDbl(tRs.Fields("GANANCIA") + 1), "0.00")
            IVA = tRs.Fields("IVA")
            IMPUESTO1 = tRs.Fields("IMPUESTO1")
            IMPUESTO2 = tRs.Fields("IMPUESTO2")
            RETENCION = tRs.Fields("P_RETENCION")
        End If
    End If
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    If Text2(8).Text <> "" And Text5.Text <> "" And Text2(9).Text <> "" Then
        Text2(9).Text = Replace(Text2(9).Text, ",", "")
        Text3.SetFocus
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim cant As String
        Dim tLi As ListItem
        Dim ClvProd As String
        ' checa si es producto de Almacen 1, 2 o 3 para descontar existencia
        If VarAlmacen <> "Activo Fijo" Then
            ' ---------------------------- PRODUCTOS EQUIVALENTES ----------------------------
            ' ------------------------------ 25/05/2021 H VALDEZ -----------------------------
            cant = CDbl(Text5.Text)
            ClvProd = Text2(8).Text
            sBuscar = "SELECT JUEGO_REPARACION.ID_PRODUCTO, JUEGO_REPARACION.CANTIDAD FROM ALMACEN3 INNER JOIN JUEGO_REPARACION ON ALMACEN3.ID_PRODUCTO = JUEGO_REPARACION.ID_REPARACION WHERE (ALMACEN3.TIPO = 'EQUIVALE') AND  ALMACEN3.ID_PRODUCTO = '" & ClvProd & "'"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                ClvProd = tRs.Fields("ID_PRODUCTO")
                cant = CDbl(tRs.Fields("CANTIDAD")) * cant
            End If
            ' ++++++++++++++++++++++++++++ PRODUCTOS EQUIVALENTES ++++++++++++++++++++++++++++
            sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ClvProd & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "' AND CANTIDAD >= " & cant
            Set tRs = cnn.Execute(sBuscar)
            If (tRs.BOF And tRs.EOF) And SinExis = 0 Then
                MsgBox "NO CUENTA CON EXISTENCIA SUFICIENTE!", vbInformation, "SACC"
            Else
                Me.Command1.Enabled = True
                If SinExis = 0 Then
                    cant = tRs.Fields("CANTIDAD") - CDbl(cant)
                    cant = Replace(cant, ",", "")
                    sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & cant & " WHERE ID_PRODUCTO = '" & ClvProd & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                    Set tRs = cnn.Execute(sBuscar)
                End If
                Set tLi = ListView3.ListItems.Add(, , Text2(8).Text)
                If Not IsNull(Text2(7).Text) Then tLi.SubItems(1) = Text2(7).Text
                If Not IsNull(Porci) Then tLi.SubItems(2) = Porci
                If Not IsNull(Text2(9).Text) Then tLi.SubItems(3) = Text2(9).Text
                If Porci <> "0" Then
                    If Not IsNull(Text2(9).Text) Then tLi.SubItems(4) = CDbl(Text2(9).Text) - (CDbl(Text2(9).Text) / CDbl(Porci))
                Else
                    If Not IsNull(Text2(9).Text) Then tLi.SubItems(4) = CDbl(Text2(9).Text) - (CDbl(Text2(9).Text))
                End If
                If Not IsNull(Text5.Text) Then tLi.SubItems(5) = Text5.Text
                tLi.SubItems(6) = Format((CDbl(Text2(9).Text) * CDbl(Text5.Text)) * CDbl(IVA), "0.00")
                tLi.SubItems(7) = Format((CDbl(Text2(9).Text) * CDbl(Text5.Text)) * CDbl(IMPUESTO1), "0.00")
                tLi.SubItems(8) = Format((CDbl(Text2(9).Text) * CDbl(Text5.Text)) * CDbl(IMPUESTO2), "0.00")
                tLi.SubItems(9) = Format((CDbl(Text2(9).Text) * CDbl(Text5.Text)) * CDbl(RETENCION), "0.00")
                Text2(11).Text = Format(CDbl(Text2(11).Text) + (CDbl(Text5.Text) * CDbl(Text2(9).Text)), "0.00")
                If RFC = "XEXX010101000" Then
                    Text11.Text = "0.00"
                    Text8.Text = "0.00"
                    Text9.Text = "0.00"
                    Text10.Text = "0.00"
                    Text12.Text = CDbl(Text2(11).Text)
                Else
                    Text8.Text = Format(CDbl(Text8.Text) + ((CDbl(Text2(9).Text) * CDbl(Text5.Text)) * CDbl(IMPUESTO1)), "0.00")
                    Text9.Text = Format(CDbl(Text9.Text) + ((CDbl(Text2(9).Text) * CDbl(Text5.Text)) * CDbl(IMPUESTO2)), "0.00")
                    Text10.Text = Format(CDbl(Text10.Text) + ((CDbl(Text2(9).Text) * CDbl(Text5.Text)) * CDbl(RETENCION)), "0.00")
                    Text11.Text = Format(CDbl(Text11.Text) + ((CDbl(Text2(9).Text) * CDbl(Text5.Text)) * CDbl(IVA)), "0.00")
                    Text12.Text = Format(CDbl(Text2(11).Text) + CDbl(Text11.Text) + CDbl(Text8.Text) + CDbl(Text9.Text) - CDbl(Text10.Text), "0.00")
                End If
                Text2(7).Text = ""
                Text2(8).Text = ""
                Text2(9).Text = ""
                Text5.Text = ""
            End If
        Else
            sBuscar = "SELECT SUM(CANTIDAD) AS CANTIDAD FROM EXISTENCIA_FIJA WHERE ID_PRODUCTO = '" & Text2(8).Text & "'"
            Set tRs = cnn.Execute(sBuscar)
            If CDbl(Text5.Text) <= tRs.Fields("CANTIDAD") Then
                Me.Command1.Enabled = True
                sBuscar = "INSERT INTO EXISTENCIA_FIJA (ID_PRODUCTO, CANTIDAD) VALUES ('" & Text2(8).Text & "', -" & Text5.Text & ");"
                cnn.Execute (sBuscar)
                Set tLi = ListView3.ListItems.Add(, , Text2(8).Text)
                If Not IsNull(Text2(7).Text) Then tLi.SubItems(1) = Text2(7).Text
                If Not IsNull(Porci) Then tLi.SubItems(2) = Porci
                If Not IsNull(Text2(9).Text) Then tLi.SubItems(3) = Text2(9).Text
                If Porci <> "0" Then
                    If Not IsNull(Text2(9).Text) Then tLi.SubItems(4) = CDbl(Text2(9).Text) - (CDbl(Text2(9).Text) / CDbl(Porci))
                Else
                    If Not IsNull(Text2(9).Text) Then tLi.SubItems(4) = CDbl(Text2(9).Text) - (CDbl(Text2(9).Text))
                End If
                If Not IsNull(Text5.Text) Then tLi.SubItems(5) = Text5.Text
                tLi.SubItems(6) = Format(CDbl(Text2(9).Text) * CDbl(VarMen.Text4(7).Text), "0.00")
                tLi.SubItems(7) = "0.00"
                tLi.SubItems(8) = "0.00"
                tLi.SubItems(9) = "0.00"
                Text2(11).Text = Format(CDbl(Text2(11).Text) + (CDbl(Text5.Text) * CDbl(Text2(9).Text)), "0.00")
                If RFC = "XEXX010101000" Then
                    Text11.Text = "0.00"
                    Text8.Text = "0.00"
                    Text9.Text = "0.00"
                    Text10.Text = "0.00"
                    Text12.Text = CDbl(Text2(11).Text)
                Else
                    Text11.Text = CDbl(Text11.Text) + (CDbl(Text2(9).Text) * CDbl(IVA))
                    Text8.Text = "0.00"
                    Text9.Text = "0.00"
                    Text10.Text = "0.00"
                    Text12.Text = CDbl(Text2(11).Text) + CDbl(Text11.Text) + CDbl(Text8.Text) + CDbl(Text9.Text) - CDbl(Text10.Text)
                End If
                Text2(7).Text = ""
                Text2(8).Text = ""
                Text2(9).Text = ""
                Text5.Text = ""
            Else
                MsgBox "NO CUENTA CON EXISTENCIA SUFICIENTE!", vbInformation, "SACC"
            End If
        End If
    Else
        MsgBox "FATLA INFORMACIÓN NECESARIA PARA AGREGAR EL ARTICULO!", vbInformation, "SACC"
    End If
    SinExis = 0
    If CantCom > ContComanda Then
        ExtraerComanda
    Else
        If CantCom <> 0 Then
            sBuscar = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'I' WHERE ID_COMANDA = " & Text4.Text
            cnn.Execute (sBuscar)
            ListView4.ListItems.Clear
            Text4.Text = ""
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command4_Click()
On Error GoTo ManejaError
    With ListView3
        If .ListItems.Count > 0 Then
            .ListItems(.SelectedItem.Index).Selected = True
            .ListItems.Remove .SelectedItem.Index
        End If
    End With
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim cant As String
    ' ---------------------------- PRODUCTOS EQUIVALENTES ----------------------------
    ' ------------------------------ 25/05/2021 H VALDEZ -----------------------------
    sBuscar = "SELECT JUEGO_REPARACION.ID_PRODUCTO, JUEGO_REPARACION.CANTIDAD FROM ALMACEN3 INNER JOIN JUEGO_REPARACION ON ALMACEN3.ID_PRODUCTO = JUEGO_REPARACION.ID_REPARACION WHERE (ALMACEN3.TIPO = 'EQUIVALE') AND  ALMACEN3.ID_PRODUCTO = '" & IdProdEli & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        IdProdEli = tRs.Fields("ID_PRODUCTO")
        Elim = CDbl(tRs.Fields("CANTIDAD")) * Elim
    End If
    ' ++++++++++++++++++++++++++++ PRODUCTOS EQUIVALENTES ++++++++++++++++++++++++++++
    sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & IdProdEli & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "' AND CANTIDAD > 0"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        cant = tRs.Fields("CANTIDAD") + Elim
        cant = Replace(cant, ",", "")
        sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & cant & " WHERE ID_PRODUCTO = '" & IdProdEli & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
        cnn.Execute (sBuscar)
        Text2(11).Text = Format(Val(Replace(Text2(11).Text, ",", "")) - (elim2), "0.00")
    Else
        ' modificacion pasra regresar Activo Fijo 12/01/2012
        sBuscar = "SELECT ID_EXIS_FIJA FROM EXISTENCIA_FIJA WHERE ID_PRODUCTO = '" & IdProdEli & "' AND CANTIDAD = -" & Elim
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.BOF And tRs.EOF) Then
            sBuscar = "DELETE FROM EXISTENCIA_FIJA WHERE ID_EXIS_FIJA = " & tRs.Fields("ID_EXIS_FIJA")
            cnn.Execute (sBuscar)
        Else
            sBuscar = "INSERT INTO EXISTENCIAS(ID_PRODUCTO, SUCURSAL, CANTIDAD) VALUES('" & IdProdEli & "','" & VarMen.Text4(0).Text & "', " & Elim & ");"
            cnn.Execute (sBuscar)
        End If
        Text2(11).Text = Format(Val(Replace(Text2(11).Text, ",", "")) - (elim2), "0.00")
    End If
    If ListView3.ListItems.Count = 0 Then
        Text2(11).Text = "0.00"
    End If
    Text11.Text = Format(CDbl(Text11.Text) - CDbl(IVAEli), "###,###,##0.00")
    Text8.Text = Format(CDbl(Text8.Text) - CDbl(Impu1Eli), "###,###,##0.00")
    Text9.Text = Format(CDbl(Text9.Text) - CDbl(Impu2Eli), "###,###,##0.00")
    Text10.Text = Format(CDbl(Text10.Text) - CDbl(RetEli), "###,###,##0.00")
    Text12.Text = Format(CDbl(Text2(11).Text) + CDbl(Text11.Text) + CDbl(Text8.Text) + CDbl(Text9.Text) - CDbl(Text10.Text), "###,###,##0.00")
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Command4.Enabled = False
    Label13.Caption = VarMen.Text1(1).Text
    Label8.Caption = VarMen.Text4(0).Text
    Me.Command1.Enabled = False
    Me.DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    CantCom = 0
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "# DEL CLIENTE", 2400
        .ColumnHeaders.Add , , "NOMBRE", 6100
        .ColumnHeaders.Add , , "DESCUENTO", 2300
        .ColumnHeaders.Add , , "DIAS DE CREDITO", 0
        .ColumnHeaders.Add , , "LIMITE DE CREDITO", 2300
        .ColumnHeaders.Add , , "RFC", 0
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE", 2400
        .ColumnHeaders.Add , , "DESCRIPCIÓN", 6800
        .ColumnHeaders.Add , , "% de GANANCIA", 0
        .ColumnHeaders.Add , , "PRECIO COSTO", 0
        .ColumnHeaders.Add , , "ALMACEN", 1100
        .ColumnHeaders.Add , , "IVA", 1100
        .ColumnHeaders.Add , , "IMPUESTO1", 1100
        .ColumnHeaders.Add , , "IMPUESTO2", 1100
        .ColumnHeaders.Add , , "RETENCIÓN", 1100
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE", 2100
        .ColumnHeaders.Add , , "DESCRIPCIÓN", 6000
        .ColumnHeaders.Add , , "% de GANANCIA", 0
        .ColumnHeaders.Add , , "PRECIO DE VENTA", 1800
        .ColumnHeaders.Add , , "PRECIO DE COSTO", 0
        .ColumnHeaders.Add , , "CANTIDAD", 1100
        .ColumnHeaders.Add , , "IVA", 1100
        .ColumnHeaders.Add , , "IMPUESTO1", 1100
        .ColumnHeaders.Add , , "IMPUESTO2", 1100
        .ColumnHeaders.Add , , "RETENCIÓN", 1100
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "PRODUCTO", 1000
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "COMANDA", 1000
    End With
    Text2(0).Text = ""
    Text2(7).Text = ""
    Text2(8).Text = ""
    Text2(9).Text = ""
    Text7.Text = VarMen.Text4(7).Text
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    If ListView3.ListItems.Count = 0 Then
        Unload Me
    Else
        MsgBox "DEBE ELIMINAR LOS ARTICULOS DEL LISTADO DE LA VENTA PARA PODER SALIR!", vbInformation, "SACC"
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdClien = Item
    NomClien = Item.SubItems(1)
    Text1.Text = Item.SubItems(1)
    Text2(0).Text = Item
    DesClien = Item.SubItems(2)
    DiasCred = Item.SubItems(3)
    LimCred = Item.SubItems(4)
    RFC = Item.SubItems(5)
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text3.SetFocus
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text2(8).Text = Item
    Text2(7).Text = Item.SubItems(1)
    Dim cod As String
    cod = Text2(8).Text
    If Item.SubItems(2) <> "" Or Val(Item.SubItems(2)) <> 0 Then
        Porci = CDbl(Item.SubItems(2))
        valor = Item.SubItems(3)
        If Porci <> 0 And valor <> 0 Then
            valor = Val(valor)
            Porci = CDbl(Item.SubItems(2)) + 1
            valor = valor * Porci
            Text2(9).Text = Format(valor, "0.00")
        Else
            Text2(9).Text = ""
        End If
    Else
        Text2(9).Text = ""
    End If
    IVA = Item.SubItems(5)
    IMPUESTO1 = Item.SubItems(6)
    IMPUESTO2 = Item.SubItems(7)
    RETENCION = Item.SubItems(8)
    VarAlmacen = Item.SubItems(4)
End Sub
Private Sub ListView2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text5.SetFocus
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Elim = CDbl(Item.SubItems(5))
    elim2 = CDbl(Item.SubItems(3)) * CDbl(Item.SubItems(5))
    Xind = Item.Index
    Command4.Enabled = True
    IdProdEli = Item
    IVAEli = Item.SubItems(6)
    Impu1Eli = Item.SubItems(7)
    Impu2Eli = Item.SubItems(8)
    RetEli = Item.SubItems(9)
End Sub
Private Sub ListView3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command4.SetFocus
    End If
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.ListView1.SetFocus
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        sBuscar = Text1.Text
        sBuscar = Replace(sBuscar, "*", "%")
        sBuscar = Replace(sBuscar, "?", "_")
        Text1.Text = sBuscar
        If IsNumeric(sBuscar) Then
            sBuscar = "SELECT ID_CLIENTE, NOMBRE, DESCUENTO, DIAS_CREDITO, LIMITE_CREDITO, RFC FROM CLIENTE WHERE NOMBRE LIKE '%" & sBuscar & "%' AND VALORACION = 'A' OR NOMBRE_COMERCIAL LIKE '%" & sBuscar & "%' AND VALORACION = 'A' OR ID_CLIENTE = '" & sBuscar & "' AND VALORACION = 'A' ORDER BY NOMBRE"
        Else
            sBuscar = "SELECT ID_CLIENTE, NOMBRE, DESCUENTO, DIAS_CREDITO, LIMITE_CREDITO, RFC FROM CLIENTE WHERE NOMBRE LIKE '%" & sBuscar & "%' AND VALORACION = 'A' OR NOMBRE_COMERCIAL LIKE '%" & sBuscar & "%' AND VALORACION = 'A' ORDER BY NOMBRE"
        End If
        Set tRs = cnn.Execute(sBuscar)
        With tRs
                ListView1.ListItems.Clear
                Do While Not .EOF
                    Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE")
                    If Not IsNull(.Fields("DESCUENTO")) Then
                        tLi.SubItems(2) = .Fields("DESCUENTO")
                    Else
                        tLi.SubItems(2) = "0"
                    End If
                    If Not IsNull(.Fields("DIAS_CREDITO")) Then
                        tLi.SubItems(3) = .Fields("DIAS_CREDITO")
                    Else
                        tLi.SubItems(3) = "0"
                    End If
                    If Not IsNull(.Fields("LIMITE_CREDITO")) Then tLi.SubItems(4) = .Fields("LIMITE_CREDITO")
                    If Not IsNull(.Fields("RFC")) Then tLi.SubItems(5) = .Fields("RFC")
                    .MoveNext
                Loop
        End With
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
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
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
Dim Valido As String
    Valido = "1234567890.-"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.-"
    If Index = 11 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub Text3_GotFocus()
    Text3.BackColor = &HFFE1E1
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
If KeyAscii = 13 Then
        ListView2.SetFocus
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        Dim Query As String
        Dim bus As String
        Dim sBus As String
        Query = Text3.Text
        ListView2.ListItems.Clear
        If Option1.Value = True Then
            sBus = "SELECT * FROM ALMACEN2 WHERE Descripcion LIKE '%" & Query & "%'"
        Else
            sBus = "SELECT * FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Query & "%'"
        End If
        Set tRs = cnn.Execute(sBus)
        With tRs
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = .Fields("Descripcion")
                tLi.SubItems(2) = .Fields("GANANCIA")
                tLi.SubItems(3) = .Fields("PRECIO_COSTO")
                tLi.SubItems(4) = "2"
                tLi.SubItems(5) = .Fields("IVA")
                tLi.SubItems(6) = .Fields("IMPUESTO1")
                tLi.SubItems(7) = .Fields("IMPUESTO2")
                tLi.SubItems(8) = .Fields("RETENCION")
                .MoveNext
            Loop
        End With
        If Option1.Value = True Then
            sBus = "SELECT * FROM ALMACEN1 WHERE Descripcion LIKE '%" & Query & "%'"
        Else
            sBus = "SELECT * FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & Query & "%'"
        End If
        Set tRs = cnn.Execute(sBus)
        With tRs
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = .Fields("Descripcion")
                tLi.SubItems(2) = .Fields("GANANCIA")
                tLi.SubItems(3) = .Fields("PRECIO_COSTO")
                tLi.SubItems(4) = "1"
                tLi.SubItems(5) = .Fields("IVA")
                tLi.SubItems(6) = .Fields("IMPUESTO1")
                tLi.SubItems(7) = .Fields("IMPUESTO2")
                tLi.SubItems(8) = .Fields("RETENCION")
                .MoveNext
            Loop
        End With
        If Option1.Value = True Then
            sBus = "SELECT * FROM ALMACEN3 WHERE Descripcion LIKE '%" & Query & "%'"
        Else
            sBus = "SELECT * FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Query & "%'"
        End If
        Set tRs = cnn.Execute(sBus)
        With tRs
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = .Fields("Descripcion")
                tLi.SubItems(2) = .Fields("GANANCIA")
                tLi.SubItems(3) = .Fields("PRECIO_COSTO")
                tLi.SubItems(4) = "3"
                tLi.SubItems(5) = .Fields("IVA")
                tLi.SubItems(6) = .Fields("IMPUESTO1")
                tLi.SubItems(7) = .Fields("IMPUESTO2")
                tLi.SubItems(8) = .Fields("P_RETENCION")
                .MoveNext
            Loop
        End With
        ' Venta de productos de activo fijo
        ' AGREGADO POR H VALDEZ 29 DE DIC DE 2011
        'If Option1.Value = True Then
        '    sBus = "SELECT PRODUCTOS_CONSUMIBLES.ID_PRODUCTO, PRODUCTOS_CONSUMIBLES.Descripcion, SUM(EXISTENCIA_FIJA.CANTIDAD) AS CANTIDAD, PRODUCTOS_CONSUMIBLES.PRECIO FROM EXISTENCIA_FIJA, PRODUCTOS_CONSUMIBLES WHERE PRODUCTOS_CONSUMIBLES.ID_PRODUCTO = EXISTENCIA_FIJA.ID_PRODUCTO AND PRODUCTOS_CONSUMIBLES.Descripcion LIKE '%" & Query & "%' GROUP BY PRODUCTOS_CONSUMIBLES.ID_PRODUCTO, PRODUCTOS_CONSUMIBLES.Descripcion, PRODUCTOS_CONSUMIBLES.PRECIO"
        'Else
        '    sBus = "SELECT PRODUCTOS_CONSUMIBLES.ID_PRODUCTO, PRODUCTOS_CONSUMIBLES.Descripcion, SUM(EXISTENCIA_FIJA.CANTIDAD) AS CANTIDAD, PRODUCTOS_CONSUMIBLES.PRECIO FROM EXISTENCIA_FIJA, PRODUCTOS_CONSUMIBLES WHERE PRODUCTOS_CONSUMIBLES.ID_PRODUCTO = EXISTENCIA_FIJA.ID_PRODUCTO AND PRODUCTOS_CONSUMIBLES.ID_PRODUCTO LIKE '%" & Query & "%' GROUP BY PRODUCTOS_CONSUMIBLES.ID_PRODUCTO, PRODUCTOS_CONSUMIBLES.Descripcion, PRODUCTOS_CONSUMIBLES.PRECIO"
        'End If
        'Set tRs = cnn.Execute(sBus)
        'With tRs
        '    Do While Not .EOF
        '        Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO"))
        '        tLi.SubItems(1) = .Fields("Descripcion")
        '        tLi.SubItems(2) = "0"
        '        tLi.SubItems(3) = .Fields("PRECIO")
        '        tLi.SubItems(4) = "Activo Fijo"
        '        .MoveNext
        '    Loop
        'End With
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
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
Private Sub Text3_LostFocus()
    Text3.BackColor = &H80000005
End Sub
Private Sub Text5_GotFocus()
    Text5.BackColor = &HFFE1E1
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command3.SetFocus
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
Private Sub Text5_LostFocus()
    Text5.BackColor = &H80000005
End Sub
Private Sub ImprTICKET() 'CVEVENTA
    Dim sBuscar As String
    Dim Acum As String
    Dim tRs As ADODB.Recordset
    Dim tRs8 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim POSY As Integer
    Dim Usuario As String
    Dim Usu As String
    Dim Cliente As String
    Dim Sucu As String
    Dim sExibicion As String
    Dim sSubtotal As String
    Dim sIVA As String
    Dim fech As String
    Dim sTotal As String
    Dim NOM As String
    Dim sTipoVenta As String
    Dim sRetencion As String
    Dim sImpuesto1 As String
    Dim sImpuesto2 As String
    If VarMen.Text4(0).Text = "BODEGA" Then
        If CveVenta <> 0 Then
            sBuscar = "SELECT ID_USUARIO, NOMBRE, SUCURSAL, FECHA, UNA_EXIBICION, SUBTOTAL, IVA, TOTAL, TIPO_PAGO, IMPUESTO1, IMPUESTO2, RETENCION FROM VENTAS WHERE ID_VENTA = " & CveVenta
            Set tRs = cnn.Execute(sBuscar)
            sSubtotal = tRs.Fields("SUBTOTAL")
            sTotal = tRs.Fields("TOTAL")
            sIVA = tRs.Fields("IVA")
            sExibicion = tRs.Fields("UNA_EXIBICION")
            sTipoVenta = tRs.Fields("TIPO_PAGO")
            sImpuesto1 = tRs.Fields("IMPUESTO1")
            sImpuesto2 = tRs.Fields("IMPUESTO2")
            sRetencion = tRs.Fields("RETENCION")
            If Not (tRs.EOF And tRs.BOF) Then
                Usuario = tRs.Fields("ID_USUARIO")
                Cliente = tRs.Fields("NOMBRE")
                Sucu = tRs.Fields("SUCURSAL")
                fech = tRs.Fields("FECHA")
                tRs.Close
                sBuscar = "SELECT NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & Usuario
                Set tRs = cnn.Execute(sBuscar)
                If tRs.EOF And tRs.BOF Then
                    Usuario = VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
                Else
                    Usuario = tRs.Fields("NOMBRE") & " " & tRs.Fields("APELLIDOS")
                End If
                tRs.Close
                Acum = "0"
                sBuscar = "SELECT * FROM SUCURSALES WHERE NOMBRE = '" & VarMen.Text4(0).Text & "' AND ELIMINADO = 'N' "
                Set tRs8 = cnn.Execute(sBuscar)
                sBuscar = "SELECT * FROM EMPRESA"
                Set tRs3 = cnn.Execute(sBuscar)
                '********************************IMPRIMIR TICKET********************************************
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                Printer.Print VarMen.Text5(0).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(tRs8.Fields("CALLE") & " COL. " & tRs8.Fields("COLONIA"))) / 2
                Printer.Print tRs8.Fields("CALLE") & " COL. " & tRs8.Fields("COLONIA")
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & tRs8.Fields("CP"))) / 2
                Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & tRs8.Fields("CP")
                Printer.Print "FECHA : " & fech
                Printer.Print "SUCURSAL : " & Sucu
                Printer.Print "TELEFONO SUCURSAL : " & tRs8.Fields("TELEFONO")
                Printer.Print "No. DE VENTA : " & CveVenta
                If sTipoVenta = "C" Then
                    Printer.Print "FORMA DE PAGO : EFECTIVO"
                Else
                    If sTipoVenta = "H" Then
                        Printer.Print "FORMA DE PAGO : CHEQUE"
                    Else
                        If sTipoVenta = "T" Then
                            Printer.Print "FORMA DE PAGO : TARJETA DE CREDITO"
                        Else
                            If sTipoVenta = "E" Then
                                Printer.Print "FORMA DE PAGO : TRANSFERENCIA ELECTRONICA"
                            Else
                                Printer.Print "FORMA DE PAGO : NO INDICADO"
                            End If
                        End If
                    End If
                End If
                Printer.Print "ATENDIDO POR : " & Usuario
                Printer.Print "CLIENTE : " & Cliente
                If sExibicion = "N" Then
                    Printer.Print "VENTA A CREDITO"
                Else
                    Printer.Print "VENTA A CONTADO"
                End If
                Printer.Print "--------------------------------------------------------------------------------"
                Printer.Print "                          NOTA DE FACTURA"
                Printer.Print "--------------------------------------------------------------------------------"
                POSY = 2900
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print "Producto"
                Printer.CurrentY = POSY
                Printer.CurrentX = 1900
                Printer.Print "Cant."
                Printer.CurrentY = POSY
                Printer.CurrentX = 3000
                Printer.Print "Precio unitario"
                sBuscar = "SELECT VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.PRECIO_VENTA, VENTAS_DETALLE.CANTIDAD, VENTAS.SUBTOTAL, VENTAS.IVA, VENTAS.TOTAL FROM VENTAS_DETALLE, VENTAS WHERE VENTAS_DETALLE.ID_VENTA = VENTAS.ID_VENTA AND VENTAS.ID_VENTA = " & CveVenta
                Set tRs = cnn.Execute(sBuscar)
                sSubtotal = tRs.Fields("SUBTOTAL")
                sTotal = tRs.Fields("TOTAL")
                sIVA = tRs.Fields("IVA")
                If Not (tRs.EOF And tRs.BOF) Then
                    Do While Not tRs.EOF
                        POSY = POSY + 200
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 100
                        Printer.Print Mid(tRs.Fields("ID_PRODUCTO"), 1, 25)
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 1900
                        Printer.Print tRs.Fields("CANTIDAD")
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 2900
                        Printer.Print Format(CDbl(tRs.Fields("PRECIO_VENTA")), "###,###,##0.00")
                        Acum = CDbl(Acum) + CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("CANTIDAD"))
                        tRs.MoveNext
                    Loop
                End If
                Printer.Print ""
                Printer.Print "SUBTOTAL : " & Format(CDbl(sSubtotal), "###,###,##0.00")
                Printer.Print "IVA              : " & Format(CDbl(sIVA), "###,###,##0.00")
                If CDbl(sImpuesto1) <> 0 Then
                    Printer.Print "IMPUESTO 1 : " & Format(CDbl(sImpuesto1), "###,###,##0.00")
                End If
                If CDbl(sImpuesto2) <> 0 Then
                    Printer.Print "IMPUESTO 2 : " & Format(CDbl(sImpuesto2), "###,###,##0.00")
                End If
                If CDbl(sRetencion) <> 0 Then
                    Printer.Print "RETENCION : " & Format(CDbl(sRetencion), "###,###,##0.00")
                End If
                Printer.Print "TOTAL        : " & Format(CDbl(sTotal), "###,###,##0.00")
                Printer.Print ""
                Printer.Print "--------------------------------------------------------------------------------"
                Printer.Print "               GRACIAS POR SU COMPRA"
                Printer.Print "LA GARANTÍA SERÁ VÁLIDA HASTA 15 DIAS"
                Printer.Print "     DESPUES DE HABER EFECTUADO SU "
                Printer.Print "                                COMPRA"
                Printer.Print "--------------------------------------------------------------------------------"
                Printer.EndDoc
            Else
                MsgBox "LA VENTA NO EXISTE!", vbInformation, "SACC"
            End If
        Else
            MsgBox "DEBE DAR EL NUMERO DE VENTA!", vbInformation, "SACC"
        End If
    Else
        If CveVenta <> 0 Then
            sBuscar = "SELECT ID_USUARIO, NOMBRE, SUCURSAL,FECHA FROM VENTAS WHERE ID_VENTA = " & CveVenta
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Usuario = tRs.Fields("ID_USUARIO")
                Cliente = tRs.Fields("NOMBRE")
                NOM = tRs.Fields("NOMBRE")
                Sucu = tRs.Fields("SUCURSAL")
                fech = tRs.Fields("FECHA")
                tRs.Close
                sBuscar = "SELECT NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & Usuario
                Set tRs = cnn.Execute(sBuscar)
                If tRs.EOF And tRs.BOF Then
                    Usuario = VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
                Else
                    Usuario = tRs.Fields("NOMBRE") & " " & tRs.Fields("APELLIDOS")
                End If
                tRs.Close
                Acum = "0"
                sBuscar = "SELECT * FROM SUCURSALES WHERE NOMBRE = '" & VarMen.Text4(0).Text & "' AND ELIMINADO = 'N' "
                Set tRs8 = cnn.Execute(sBuscar)
                sBuscar = "SELECT * FROM EMPRESA"
                Set tRs3 = cnn.Execute(sBuscar)
                '********************************IMPRIMIR TICKET********************************************
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                Printer.Print VarMen.Text5(0).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(tRs8.Fields("CALLE") & " COL. " & tRs8.Fields("COLONIA"))) / 2
                Printer.Print tRs8.Fields("CALLE") & " COL. " & tRs8.Fields("COLONIA")
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & tRs8.Fields("CP"))) / 2
                Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & tRs8.Fields("CP")
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("MATRIZ" & (tRs3.Fields("DIRECCION")) & " COL. " & tRs8.Fields("COLONIA"))) / 2
                Printer.Print "MATRIZ : " & tRs3.Fields("DIRECCION") & " COL. " & tRs3.Fields("COLONIA")
                Printer.Print "FECHA : " & fech
                Printer.Print "SUCURSAL : " & VarMen.Text4(0).Text
                Printer.Print "TELEFONO SUCURSAL : " & tRs8.Fields("TELEFONO")
                Printer.Print "No. DE VENTA : " & CveVenta
                Printer.Print "ATENDIDO POR : " & VarMen.Text1(1).Text & " "; VarMen.Text1(2).Text
                Printer.Print "CLIENTE : " & NOM
                Printer.Print "--------------------------------------------------------------------------------"
                Printer.Print "                          NOTA DE FACTURA"
                Printer.Print "--------------------------------------------------------------------------------"
                POSY = 2400
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print "Producto"
                Printer.CurrentY = POSY
                Printer.CurrentX = 1900
                Printer.Print "Cant."
                Printer.CurrentY = POSY
                Printer.CurrentX = 3000
                Printer.Print "Precio unitario"
                sBuscar = "SELECT ID_PRODUCTO, PRECIO_VENTA, CANTIDAD FROM VENTAS_DETALLE WHERE ID_VENTA = " & CveVenta
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    Do While Not tRs.EOF
                        POSY = POSY + 200
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 100
                        Printer.Print tRs.Fields("ID_PRODUCTO")
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 1900
                        Printer.Print tRs.Fields("CANTIDAD")
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 2900
                        Printer.Print Format(CDbl(tRs.Fields("PRECIO_VENTA")), "###,###,##0.00")
                        Acum = CDbl(Acum) + CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("CANTIDAD"))
                        tRs.MoveNext
                    Loop
                End If
                Printer.Print ""
                If sSubtotal = "" Then
                    sSubtotal = Format(CDbl(Text2(11).Text), "0.00")
                End If
                If sImpuesto1 <> "" Then
                    If CDbl(sImpuesto1) <> 0 Then
                        Printer.Print "IMPUESTO 1 : " & Format(CDbl(sImpuesto1), "###,###,##0.00")
                    End If
                End If
                If sImpuesto2 <> "" Then
                    If CDbl(sImpuesto2) <> 0 Then
                        Printer.Print "IMPUESTO 2 : " & Format(CDbl(sImpuesto2), "###,###,##0.00")
                    End If
                End If
                If sRetencion <> "" Then
                    If CDbl(sRetencion) <> 0 Then
                        Printer.Print "RETENCION : " & Format(CDbl(sRetencion), "###,###,##0.00")
                    End If
                End If
                Printer.Print "SUBTOTAL : " & Format(CDbl(sSubtotal), "###,###,##0.00")
                If sIVA = "" Then
                    sIVA = Format(CDbl(Text11.Text), "0.00")
                End If
                Printer.Print "IVA              : " & Format(CDbl(sIVA), "###,###,##0.00")
                If sTotal = "" Then
                    sTotal = Format(CDbl(Text12.Text), "0.00")
                End If
                Printer.Print "TOTAL        : " & Format(CDbl(sTotal), "###,###,##0.00")
                Printer.Print ""
                Printer.Print "--------------------------------------------------------------------------------"
                Printer.Print "               GRACIAS POR SU COMPRA"
                Printer.Print "           PRODUCTO 100% GARANTIZADO"
                Printer.Print "LA GARANTÍA SERÁ VÁLIDA HASTA 15 DIAS"
                Printer.Print "     DESPUES DE HABER EFECTUADO SU "
                Printer.Print "                                COMPRA"
                Printer.Print "                APLICA RESTRICCIONES"
                Printer.Print "--------------------------------------------------------------------------------"
                Printer.EndDoc
            Else
                MsgBox "LA VENTA NO EXISTE!", vbInformation, "SACC"
            End If
        Else
            MsgBox "DEBE DAR EL NUMERO DE VENTA!", vbInformation, "SACC"
        End If
     End If
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command3.SetFocus
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
Private Sub FunRemision()
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
    Dim tRs5 As ADODB.Recordset
    Dim tRs6 As ADODB.Recordset
    Dim ConPag As Integer
    ConPag = 1
    Dim sBuscar As String
    sBuscar = "SELECT * FROM VENTAS WHERE ID_VENTA = " & CveVenta & ""
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\Remision.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        Image4.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image4, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 20, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 205, 20, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "Remision : " & tRs1.Fields("ID_VENTA"), "F3", 8, hCenter
        oDoc.WTextBox 70, 340, 20, 250, "Fecha :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        
        
        'CAJA1
        sBuscar = "SELECT * FROM CLIENTE WHERE ID_CLIENTE = " & tRs1.Fields("ID_CLIENTE")
        Set tRs2 = cnn.Execute(sBuscar)
        oDoc.WTextBox 110, 20, 100, 400, "CLIENTE:", "F3", 8, hLeft
        oDoc.WTextBox 120, 20, 100, 400, "DOMICILIO", "F3", 8, hLeft
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 110, 20, 100, 400, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 120, 20, 100, 400, tRs2.Fields("DIRECCION") & "Col. " & tRs2.Fields("COLONIA"), "F3", 8, hCenter
        End If
        Posi = 150
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 20, 50, "CANTIDAD", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 55, 20, 90, "CLAVE", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 145, 20, 280, "DESCRIPCION", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 425, 20, 60, "PRESENTACION", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 485, 20, 50, "PRECIO UNITARIO", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 535, 20, 50, "TOTAL", "F2", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 20

        ' DETALLE
        sBuscar = "SELECT VENTAS_DETALLE.CANTIDAD, VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.DESCRIPCION, ALMACEN3.PRESENTACION, VENTAS_DETALLE.PRECIO_VENTA, VENTAS_DETALLE.PRECIO_VENTA * VENTAS_DETALLE.CANTIDAD AS TOTAL FROM ALMACEN3 INNER JOIN VENTAS_DETALLE ON ALMACEN3.ID_PRODUCTO = VENTAS_DETALLE.ID_PRODUCTO WHERE VENTAS_DETALLE.ID_VENTA = " & tRs1.Fields("ID_VENTA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 5, 15, 50, Format(tRs3.Fields("CANTIDAD"), "###,###,##0.00"), "F3", 7, hCenter, , , 1, vbBlack
                
                'oDoc.WTextBox Posi, 55, 15, 90, " " & tRs3.Fields("ID_PRODUCTO"), "F3", 7, hLeft, , , 1, vbBlack
                oDoc.WTextBox Posi, 55, 15, 30, " " & Mid(tRs3.Fields("ID_PRODUCTO"), 1, 3), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 85, 15, 30, " " & Mid(tRs3.Fields("ID_PRODUCTO"), 4, 3), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 115, 15, 30, " " & Mid(tRs3.Fields("ID_PRODUCTO"), 7, 3), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 145, 15, 280, " " & tRs3.Fields("DESCRIPCION"), "F3", 7, hLeft, , , 1, vbBlack
                'oDoc.WTextBox Posi, 425, 15, 60, tRs3.Fields("PRESENTACION"), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 425, 15, 60, Mid(tRs3.Fields("ID_PRODUCTO"), 11, 7), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 485, 15, 50, Format(CDbl(tRs3.Fields("PRECIO_VENTA")), "###,###,##0.00") & " ", "F3", 7, hRight, , , 1, vbBlack
                oDoc.WTextBox Posi, 535, 15, 50, Format(CDbl(tRs3.Fields("TOTAL")), "###,###,##0.00") & " ", "F3", 7, hRight, , , 1, vbBlack
                Posi = Posi + 15
                tRs3.MoveNext
                If Posi >= 600 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    sBuscar = "SELECT * FROM VENTAS WHERE ID_VENTA = " & tRs1.Fields("ID_VENTA")
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
                        oDoc.NewPage A4_Vertical
                        oDoc.WImage 70, 40, 43, 161, "Logo"
                        oDoc.WTextBox 40, 205, 20, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
                        oDoc.WTextBox 60, 205, 20, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
                
                        oDoc.WTextBox 60, 340, 20, 250, "Remision : " & tRs1.Fields("ID_VENTA"), "F3", 8, hCenter
                        oDoc.WTextBox 70, 340, 20, 250, "Fecha :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
                        Posi = Posi + 15
                        oDoc.WTextBox 110, 20, 100, 400, "CLIENTE:", "F3", 8, hLeft
                        oDoc.WTextBox 120, 20, 100, 400, "DOMICILIO", "F3", 8, hLeft
                        If Not (tRs2.EOF And tRs2.BOF) Then
                            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 110, 20, 100, 400, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
                            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 120, 20, 100, 400, tRs2.Fields("DIRECCION") & "Col. " & tRs2.Fields("COLONIA"), "F3", 8, hCenter
                        End If
                        Posi = 210
                    End If
                End If
            Loop
        End If
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 600
        oDoc.WLineTo 580, 600
        oDoc.LineStroke
        Posi = Posi + 6
        ' TEXTO ABAJO
        'oDoc.WTextBox 620, 20, 100, 275, "Envíe por favor los artículos a la direccion de envío, si usted tiene alguna duda o pregunta sobre esta orden de compra, por favor contacte a  Dept Compras, al Tel. " & VarMen.TxtEmp(2).Text, "F3", 8, hLeft, , , 0, vbBlack
        oDoc.WTextBox 620, 400, 20, 70, "SUBTOTAL:", "F2", 8, hRight
        oDoc.WTextBox 640, 400, 20, 70, "Impuesto 1:", "F2", 8, hRight
        oDoc.WTextBox 660, 400, 20, 70, "Impuesto 2:", "F2", 8, hRight
        oDoc.WTextBox 680, 400, 20, 70, "Retencion:", "F2", 8, hRight
        oDoc.WTextBox 700, 400, 20, 70, "I.V.A:", "F2", 8, hRight
        oDoc.WTextBox 720, 400, 20, 70, "TOTAL:", "F2", 8, hRight
        'totales
        oDoc.WTextBox 620, 488, 20, 50, Format(tRs1.Fields("TOTAL"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("IMPUESTO1")) Then oDoc.WTextBox 640, 488, 20, 50, Format(tRs1.Fields("IMPUESTO1"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("IMPUESTO2")) Then oDoc.WTextBox 660, 488, 20, 50, Format(tRs1.Fields("IMPUESTO2"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("RETENCION")) Then oDoc.WTextBox 680, 488, 20, 50, Format(tRs1.Fields("RETENCION"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("IVA")) Then oDoc.WTextBox 700, 488, 20, 50, Format(tRs1.Fields("IVA"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("TOTAL")) Then oDoc.WTextBox 720, 488, 20, 50, Format(tRs1.Fields("TOTAL"), "###,###,##0.00"), "F3", 8, hRight
        'oDoc.WTextBox 720, 200, 20, 250, "Firma de autorizado", "F3", 10, hCenter
        'If tRs1.Fields("CONFIRMADA") = "E" Then
        '    oDoc.WTextBox 750, 200, 25, 250, "COPIA CANCELADA", "F3", 25, hCenter, , vbBlue
        'End If
        'oDoc.WTextBox 620, 15, 20, 250, "Precios expresados en " & tRs1.Fields("MONEDA"), "F3", 10, hCenter
        oDoc.WTextBox 620, 20, 100, 275, "OBSERVACIONES:", "F3", 8, hLeft, , , 0, vbBlack
        If Not IsNull(tRs1.Fields("COMENTARIO")) Then oDoc.WTextBox 640, 60, 100, 300, tRs1.Fields("COMENTARIO"), "F3", 8, hLeft, , , 0, vbBlack
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs6 = cnn.Execute(sBuscar)
            If Not (tRs6.EOF And tRs6.BOF) Then
                oDoc.WTextBox 700, 20, 100, 275, "RESPONSABLE : ", "F3", 8, hLeft
                oDoc.WTextBox 720, 20, 100, 275, tRs6.Fields("NOMBRE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hLeft
            End If
        End If
        ' cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "No se enc ontro la orden de compra solicitada, puede ser que este cancelda o aun no se genere el folio", vbExclamation, "SACC"
    End If
End Sub
