VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form MENU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SISTEMA AP TONER"
   ClientHeight    =   7890
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11055
   ClipControls    =   0   'False
   Icon            =   "MENU.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   5
      Left            =   4560
      TabIndex        =   19
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   18
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   3
      Left            =   4320
      TabIndex        =   17
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   2
      Left            =   4200
      TabIndex        =   16
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1920
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SUCURSALES"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   13
      Top             =   7560
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "21/02/2006"
            Object.ToolTipText     =   "FECHA"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "16:07"
            Object.ToolTipText     =   "HORA"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4940
            MinWidth        =   4940
            Text            =   "SISTEMA AP TONER 1.1"
            TextSave        =   "SISTEMA AP TONER 1.1"
            Object.ToolTipText     =   "Sistema Ap Toner 1.0"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4680
      MaskColor       =   &H80000013&
      TabIndex        =   11
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del usuario"
      Height          =   2415
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   7695
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1200
         Width           =   5175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "PASSWORD"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "NOMBRE"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   7575
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   11055
      _cx             =   19500
      _cy             =   13361
      FlashVars       =   ""
      Movie           =   "C:/indexAP.SWF"
      Src             =   "C:/indexAP.SWF"
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   3600
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   3480
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   3360
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "USUARIOS"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu Ventillas 
      Caption         =   "Ventas"
      Enabled         =   0   'False
      Begin VB.Menu Comandas 
         Caption         =   "Comandas"
      End
      Begin VB.Menu Venta 
         Caption         =   "Nota de Venta"
      End
      Begin VB.Menu Facturar 
         Caption         =   "Factura"
      End
      Begin VB.Menu NotaCred 
         Caption         =   "Nota Credito"
      End
      Begin VB.Menu Garantias 
         Caption         =   "Garantias"
      End
      Begin VB.Menu Cotizacion2 
         Caption         =   "Cotizacion"
      End
      Begin VB.Menu AsTec 
         Caption         =   "Asistencia Tecnica"
      End
      Begin VB.Menu Corte 
         Caption         =   "Corte de Caja"
      End
      Begin VB.Menu Busca 
         Caption         =   "Buscar"
         Begin VB.Menu BusPro 
            Caption         =   "Producto"
         End
         Begin VB.Menu BusExi 
            Caption         =   "Existencia"
         End
      End
   End
   Begin VB.Menu Pedido 
      Caption         =   "Pedidos"
      Enabled         =   0   'False
      Begin VB.Menu PedidosSuc 
         Caption         =   "Pedidos"
      End
      Begin VB.Menu Requisicion 
         Caption         =   "Requisicion"
      End
   End
   Begin VB.Menu Almacen 
      Caption         =   "Almacen"
      Enabled         =   0   'False
      Begin VB.Menu Entradas 
         Caption         =   "Entradas"
         Begin VB.Menu EntAlm1 
            Caption         =   "Entrada Almacen 1"
         End
         Begin VB.Menu EntAl2 
            Caption         =   "Entrada Almacen 2"
         End
         Begin VB.Menu EntAl3 
            Caption         =   "Entrada Almacen 3"
         End
      End
      Begin VB.Menu RegProd 
         Caption         =   "Registrar Producto"
         Begin VB.Menu Alm2 
            Caption         =   "Almacen 1 y 2"
         End
         Begin VB.Menu Alm3 
            Caption         =   "Almacen 3"
         End
      End
      Begin VB.Menu TrasInvent 
         Caption         =   "Traspasos de Inventario"
      End
      Begin VB.Menu BusEntr 
         Caption         =   "Buscar Entrada"
      End
      Begin VB.Menu VerPerSuc 
         Caption         =   "Ver Pedidos de Sucursales"
      End
      Begin VB.Menu Inventarios 
         Caption         =   "Inventarios"
      End
   End
   Begin VB.Menu Compras 
      Caption         =   "Compras"
      Enabled         =   0   'False
      Begin VB.Menu OrCompr 
         Caption         =   "Orden de Compra"
      End
   End
   Begin VB.Menu AsTecMen 
      Caption         =   "Asistencia Tecnica"
      Enabled         =   0   'False
      Begin VB.Menu VerAsTec 
         Caption         =   "Ver"
      End
   End
   Begin VB.Menu Recursos 
      Caption         =   "Administrador"
      Enabled         =   0   'False
      Begin VB.Menu Agregar 
         Caption         =   "Agregar"
         Begin VB.Menu Agente2 
            Caption         =   "Agente"
            Enabled         =   0   'False
         End
         Begin VB.Menu Cliente 
            Caption         =   "Cliente"
         End
         Begin VB.Menu Sucursal 
            Caption         =   "Sucursal"
            Enabled         =   0   'False
         End
         Begin VB.Menu ProveedorMen 
            Caption         =   "Proveedor"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu Eliminar 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Begin VB.Menu ElAgente 
            Caption         =   "Agente"
         End
         Begin VB.Menu ElCliente 
            Caption         =   "Cliente"
         End
         Begin VB.Menu ElProv 
            Caption         =   "Proveedor"
         End
      End
      Begin VB.Menu Modificar 
         Caption         =   "Modificar"
         Enabled         =   0   'False
         Begin VB.Menu ModPrecio 
            Caption         =   "Precio"
         End
      End
      Begin VB.Menu Excel 
         Caption         =   "Pasar a Excel"
      End
   End
   Begin VB.Menu Produccion 
      Caption         =   "Produccion"
      Enabled         =   0   'False
      Begin VB.Menu JuegRep 
         Caption         =   "Nuevo Juego de Reparación"
      End
      Begin VB.Menu Ver 
         Caption         =   "Ver pedidos"
      End
      Begin VB.Menu VerCl 
         Caption         =   "Ver pedidos de Clientes"
      End
      Begin VB.Menu VerSuc 
         Caption         =   "Ver pedidos de Sucursales"
      End
      Begin VB.Menu VJR 
         Caption         =   "Ver Juegos de Reparación"
      End
   End
   Begin VB.Menu Utilerias 
      Caption         =   "Utilerias"
      Enabled         =   0   'False
      Begin VB.Menu Promociones 
         Caption         =   "Promociones"
      End
      Begin VB.Menu Marcas 
         Caption         =   "Marcas"
      End
      Begin VB.Menu dolar2 
         Caption         =   "Dolar"
      End
   End
   Begin VB.Menu Exit 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Agente2_Click()
    Agente.Show vbModal
End Sub
Private Sub Alm2_Click()
    frmAlmacen2.Show vbModal
End Sub
Private Sub Alm3_Click()
    frmAlmacen3.Show vbModal
End Sub
Private Sub AsTec_Click()
    AsisTec.Show vbModal
End Sub
Private Sub BusEntr_Click()
    BuscaEntrada.Show vbModal
End Sub
Private Sub BusExi_Click()
    BuscaExist.Show vbModal
End Sub
Private Sub BusPro_Click()
    BuscaProd.Show vbModal
End Sub
Private Sub Cliente_Click()
    AltaClien.Show vbModal
End Sub
Private Sub Comandas_Click()
    frmComandas.Show vbModal
End Sub
Private Sub Command1_Click()
    checar
End Sub
Private Sub Corte_Click()
    CorteCaja.Show vbModal
End Sub
Private Sub Cotizacion2_Click()
    Cotiza.Show vbModal
End Sub
Private Sub dolar2_Click()
    Dolar.Show vbModal
End Sub
Private Sub ElAgente_Click()
    EliAgente.Show vbModal
End Sub
Private Sub ElCliente_Click()
    EliCliente.Show vbModal
End Sub
Private Sub ElProv_Click()
    EliProveedor.Show vbModal
End Sub
Private Sub EntAl2_Click()
    EntradaProd2.Show vbModal
End Sub
Private Sub EntAl3_Click()
    EntradaProd3.Show vbModal
End Sub
Private Sub EntAlm1_Click()
    EntradaProd.Show vbModal
End Sub
Private Sub Excel_Click()
    BajaExcel.Show vbModal
End Sub
Private Sub Exit_Click()
    Unload Me
End Sub
Private Sub Facturar_Click()
    frmFactura.Show vbModal
End Sub
Private Sub Form_Load()
    Dim I As Long
    For I = 0 To 5
        Set Text1(I).DataSource = Adodc1
    Next
    Text1(0).DataField = "ID_USUARIO"
    Text1(1).DataField = "NOMBRE"
    Text1(2).DataField = "APELLIDOS"
    Text1(3).DataField = "PUESTO"
    Text1(4).DataField = "PASSWORD"
    Text1(5).DataField = "ID_SUCURSAL"
    Dim X As Long
    For X = 0 To 5
        Set Text4(X).DataSource = Adodc2
    Next
    Text4(0).DataField = "NOMBRE"
    Text4(1).DataField = "CALLE"
    Text4(2).DataField = "COLONIA"
    Text4(3).DataField = "CIUDAD"
    Text4(4).DataField = "ESTADO"
    Text4(5).DataField = "TELEFONO"
End Sub
Private Sub Garantias_Click()
    frmGarantias.Show vbModal
End Sub
Private Sub Inventarios_Click()
    frmSucInv.Show vbModal
End Sub
Private Sub JuegRep_Click()
    JuegoRep.Show vbModal
End Sub
Private Sub Marcas_Click()
    Marca.Show vbModal
End Sub
Private Sub ModPrecio_Click()
    CambioPRe.Show vbModal
End Sub
Private Sub NotaCred_Click()
    NotaCrd.Show vbModal
End Sub
Private Sub OrCompr_Click()
    Orden.Show vbModal
End Sub
Private Sub PedidosSuc_Click()
    Pedidos.Show vbModal
End Sub
Private Sub Promociones_Click()
    Promos.Show vbModal
End Sub
Private Sub ProveedorMen_Click()
    Proveedor.Show vbModal
End Sub
Private Sub Requisicion_Click()
    frmRequi.Show vbModal
End Sub
Private Sub Sucursal_Click()
    AltaSucu.Show vbModal
End Sub
Private Sub Text1_Change(Index As Integer)
    If Index = 5 Then
        Dim nReg2 As String
        Dim vBookmark2 As Variant
        Dim sADOBuscar2 As String
        On Error Resume Next
        nReg2 = Text1(5).Text
        sADOBuscar2 = "ID_SUCURSAL = " & nReg2 '& "'"
        vBookmark2 = Adodc2.Recordset.Bookmark
        Adodc2.Recordset.MoveFirst
        Adodc2.Recordset.Find sADOBuscar2
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text2.Text <> "" And Text3.Text <> "" Then
        checar
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text2.Text <> "" And Text3.Text <> "" Then
        checar
    End If
End Sub
Private Sub TrasInvent_Click()
    Transfe.Show vbModal
End Sub
Private Sub Venta_Click()
    Ventas.Show vbModal
End Sub
Private Sub Ver_Click()
    frmProd.Show vbModal
End Sub
Private Sub VerAsTec_Click()
    frmAT.Show vbModal
End Sub
Private Sub VerCl_Click()
    frmRevCom.Show vbModal
End Sub
Private Sub VerPerSuc_Click()
    frmRevPed.Show vbModal
End Sub
Private Sub VerSuc_Click()
    frmRevComSuc.Show vbModal
End Sub
Private Sub VJR_Click()
    VerJuegoRep.Show vbModal
End Sub
Private Sub checar()
    If Text2.Text <> "" And Text3.Text <> "" Then
        Dim nReg As String
        Dim vBookmark As Variant
        Dim sADOBuscar As String
        On Error Resume Next
        nReg = Text2.Text
        sADOBuscar = "NOMBRE LIKE '" & nReg & "'"
        vBookmark = Adodc1.Recordset.Bookmark
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset.Find sADOBuscar
        Text2.Text = Replace(Text2.Text, " ", "")
        Text1(1).Text = Replace(Text1(1).Text, " ", "")
        Text1(4).Text = Replace(Text1(4).Text, " ", "")
        Text3.Text = Replace(Text3.Text, " ", "")
        If Text2.Text = Text1(1).Text And Text3.Text = Text1(4).Text Then
            Me.ShockwaveFlash1.Visible = True
            Me.Frame1.Visible = False
            Me.Command1.Visible = False
            Me.Ventillas.Enabled = True
            'Me.Utilerias.Enabled = True
            'Me.Produccion.Enabled = True
            Me.Recursos.Enabled = True
            'Me.Almacen.Enabled = True
            'Me.Compras.Enabled = True
            Me.Pedido.Enabled = True
            'Me.AsTecMen.Enabled = True
        Else
            If Adodc1.Recordset.EOF Or Text2.Text <> Text1(1).Text Or Text3.Text <> Text1(4).Text Then
                Err.Clear
                MsgBox "El Nombre o el Password son incorrectos."
                Adodc1.Recordset.Bookmark = vBookmark
            End If
        End If
    End If
End Sub
