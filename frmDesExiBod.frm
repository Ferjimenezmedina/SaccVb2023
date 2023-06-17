VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDesExiBod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DESCUENTO DE INVENTARIOS"
   ClientHeight    =   6780
   ClientLeft      =   3165
   ClientTop       =   7935
   ClientWidth     =   11175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   9720
      TabIndex        =   1
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtDescontado 
      Height          =   285
      Left            =   9480
      TabIndex        =   23
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtArticulo 
      Height          =   285
      Left            =   2400
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtComanda 
      Height          =   285
      Left            =   1320
      TabIndex        =   21
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtProDes 
      Height          =   285
      Left            =   3480
      TabIndex        =   20
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCanExi 
      Height          =   285
      Left            =   4560
      TabIndex        =   19
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCanDes 
      Height          =   285
      Left            =   5640
      TabIndex        =   18
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCantidadJr2 
      Height          =   285
      Left            =   7800
      TabIndex        =   16
      Top             =   6480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtId_ProductoJr2 
      Height          =   285
      Left            =   6720
      TabIndex        =   15
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtSucursal 
      Height          =   285
      Left            =   8880
      TabIndex        =   14
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtCantidadJr 
      Height          =   285
      Left            =   2400
      TabIndex        =   13
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtId_ProductoJr 
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCantidad2 
      Height          =   285
      Left            =   8880
      TabIndex        =   9
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtId_Reparacion 
      Height          =   285
      Left            =   6720
      TabIndex        =   8
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtId_Producto2 
      Height          =   285
      Left            =   7800
      TabIndex        =   7
      Top             =   6240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtId_Producto 
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCantidad_No 
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   5640
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc adcJuegos_Reparacion 
      Height          =   330
      Left            =   0
      Top             =   6480
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
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc adcExistencias 
      Height          =   330
      Left            =   120
      Top             =   6120
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin MSAdodcLib.Adodc adcComandas_Detalles 
      Height          =   330
      Left            =   120
      Top             =   6360
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
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin MSComctlLib.ListView lvwComandas_Detalles 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PEDIDO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "NO."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "PRODUCTO"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "CANTIDAD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "¿LLEGO COMPLRETO?"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "NO LLEGO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "REVISADO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "PROCESADO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "DESCONTADO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "NO_SIRVIO"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvwExistencias 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NUMERO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "PRODUCTO"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "CANTIDAD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "SUCURSAL"
         Object.Width           =   5080
      EndProperty
   End
   Begin MSComctlLib.ListView lvwJuegos_Reparacion 
      Height          =   615
      Left            =   6720
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1085
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID_REPARACION"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID_PRODUCTO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CANTIDAD"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwJR 
      Height          =   615
      Left            =   6720
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1085
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NUM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID_PRODUCTO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CANTIDAD"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwE 
      Height          =   615
      Left            =   6720
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1085
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NUM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID_PRODUCTO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CANTIDAD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CANTIDAD A DESCONTAR"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwNo_Hay 
      Height          =   5895
      Left            =   6720
      TabIndex        =   17
      Top             =   240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   10398
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NUM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID_PRODUCTO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CANTIDAD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CANTIDAD FALTA"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmDesExiBod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim sqlComDet As String
Dim sqlJueRep As String
Dim sqlExc As String
Dim tRs As Recordset
Dim tLi As ListItem
Dim NoRe As Integer
Dim NoRe3 As Integer
Dim NoRe5 As Integer
Dim NoRe6 As Integer
Dim Cont As Integer
Dim ConT2 As Integer
Dim ConT3 As Integer
Dim ConT4 As Integer
Dim ConT5 As Integer
Dim ConT6 As Integer
Dim CantidadTotal As Double
Dim CantidadTotalJr As Double
Dim CanPed As Double
Dim BanNoHay As Boolean
Dim sqlComanda As String
Dim BanNoHayLlenar As Boolean
Dim sqlDescontado As String
Sub Descuenta()
    Me.adcComandas_Detalles.Refresh
    Me.adcExistencias.Refresh
    Me.adcJuegos_Reparacion.Refresh
    COMANDAS_DETALLES
    JUEGO_REPARACION
    EXISTENCIAS
    BanNoHay = False
    Me.lvwNo_Hay.ListItems.Clear
    NoRe = Me.lvwComandas_Detalles.ListItems.Count
    For Cont = 1 To NoRe
        Me.txtComanda.Text = Me.lvwComandas_Detalles.ListItems(Cont)
        Me.txtArticulo.Text = Me.lvwComandas_Detalles.ListItems(Cont).SubItems(1)
        Me.txtId_Producto.Text = Me.lvwComandas_Detalles.ListItems(Cont).SubItems(2)
        Me.txtCantidad.Text = Me.lvwComandas_Detalles.ListItems(Cont).SubItems(3)
        Me.txtCantidad_No.Text = Me.lvwComandas_Detalles.ListItems(Cont).SubItems(5)
        Me.txtDescontado.Text = Me.lvwComandas_Detalles.ListItems(Cont).SubItems(8)
        CantidadTotal = Me.txtCantidad.Text - Me.txtCantidad_No.Text
        If CantidadTotal > 0 Then
            Me.adcJuegos_Reparacion.Recordset.MoveFirst
            Me.lvwJR.ListItems.Clear
            Me.lvwE.ListItems.Clear
                Do While Not Me.adcJuegos_Reparacion.Recordset.EOF
                    If RTrim(Me.txtId_Producto.Text) = RTrim(Me.txtId_Reparacion.Text) Then
                        ConT2 = Me.lvwJR.ListItems.Count + 1
                        Set tLi = Me.lvwJR.ListItems.Add(, , ConT2)
                        tLi.SubItems(1) = Trim(Me.txtId_Producto2.Text)
                        tLi.SubItems(2) = Trim(Me.txtCantidad2.Text)
                        Me.adcJuegos_Reparacion.Recordset.MoveNext
                    Else
                        Me.adcJuegos_Reparacion.Recordset.MoveNext
                    End If
                Loop
            NoRe3 = Me.lvwJR.ListItems.Count
            For ConT3 = 1 To NoRe3
                Me.txtId_ProductoJr.Text = Me.lvwJR.ListItems.Item(ConT3).SubItems(1)
                Me.txtCantidadJr.Text = Me.lvwJR.ListItems.Item(ConT3).SubItems(2)
                    Me.adcExistencias.Recordset.MoveFirst
                    Do While Not Me.adcExistencias.Recordset.EOF
                        If RTrim(Me.txtSucursal.Text) = "BODEGA" Then
                            If RTrim(Me.txtId_ProductoJr.Text) = RTrim(Me.txtId_ProductoJr2.Text) Then
                                ConT3 = Me.lvwE.ListItems.Count + 1
                                Set tLi = Me.lvwE.ListItems.Add(, , ConT3)
                                tLi.SubItems(1) = Trim(Me.txtId_ProductoJr2.Text)
                                tLi.SubItems(2) = Trim(Me.txtCantidadJr2.Text)
                                CantidadTotalJr = Trim(Me.txtCantidadJr.Text) * CantidadTotal
                                tLi.SubItems(3) = CantidadTotalJr
                                If Trim(Me.txtCantidadJr2.Text) < CantidadTotalJr Then
                                    BanNoHay = True
                                        NoRe6 = Me.lvwNo_Hay.ListItems.Count
                                        For ConT6 = 1 To NoRe6
                                            If Me.lvwNo_Hay.ListItems.Item(ConT6).SubItems(1) = Trim(Me.txtId_ProductoJr2.Text) Then
                                                BanNoHayLlenar = True
                                                Exit For
                                            End If
                                        Next ConT6
                                            If BanNoHayLlenar = True Then
                                                NoRe6 = Me.lvwNo_Hay.ListItems.Count
                                                For ConT6 = 1 To NoRe6
                                                    If Me.lvwNo_Hay.ListItems.Item(ConT6).SubItems(1) = Trim(Me.txtId_ProductoJr2.Text) Then
                                                        CanPed = CantidadTotalJr - Val(Me.txtCantidadJr2.Text)
                                                        Me.lvwNo_Hay.ListItems.Item(ConT6).SubItems(3) = Me.lvwNo_Hay.ListItems.Item(ConT6).SubItems(3) + CanPed
                                                    End If
                                                Next ConT6
                                            Else
                                                ConT4 = Me.lvwNo_Hay.ListItems.Count + 1
                                                Set tLi = Me.lvwNo_Hay.ListItems.Add(, , ConT4)
                                                tLi.SubItems(1) = Trim(Me.txtId_ProductoJr2.Text)
                                                tLi.SubItems(2) = Trim(Me.txtCantidadJr2.Text)
                                                CanPed = CantidadTotalJr - Val(Me.txtCantidadJr2.Text)
                                                tLi.SubItems(3) = CanPed
                                            End If
                                End If
                                Me.adcExistencias.Recordset.MoveNext
                            Else
                                Me.adcExistencias.Recordset.MoveNext
                            End If
                        Else
                            Me.adcExistencias.Recordset.MoveNext
                        End If
                    Loop
            Next ConT3
            If BanNoHay = False Then
                NoRe5 = Me.lvwE.ListItems.Count
                For ConT5 = 1 To NoRe5
                    Me.txtProDes.Text = Me.lvwE.ListItems.Item(ConT5).SubItems(1)
                    Me.txtCanExi.Text = Me.lvwE.ListItems.Item(ConT5).SubItems(2)
                    Me.txtCanDes.Text = Me.lvwE.ListItems.Item(ConT5).SubItems(3)
                        Me.adcExistencias.Recordset.MoveFirst
                        Do While Not Me.adcExistencias.Recordset.EOF
                            If RTrim(Me.txtSucursal.Text) = "BODEGA" Then
                                If RTrim(Me.txtProDes.Text) = RTrim(Me.txtId_ProductoJr2.Text) Then
                                    Me.txtCantidadJr2.Text = Val(Me.txtCantidadJr2.Text) - Val(Me.txtCanDes.Text)
                                    Me.adcExistencias.Recordset.MoveNext
                                Else
                                    Me.adcExistencias.Recordset.MoveNext
                                End If
                            Else
                                Me.adcExistencias.Recordset.MoveNext
                            End If
                        Loop
                Next ConT5
                sqlComanda = "UPDATE COMANDAS_DETALLES SET PROCESADO = 1 WHERE COMANDA = " & Me.txtComanda.Text & " AND ARTICULO = " & Me.txtArticulo.Text
                Set tRs = cnn.Execute(sqlComanda)
            End If
        Else
            sqlComanda = "UPDATE COMANDAS_DETALLES SET PROCESADO = 1 WHERE COMANDA = " & Me.txtComanda.Text & " AND ARTICULO = " & Me.txtArticulo.Text
            Set tRs = cnn.Execute(sqlComanda)
        End If
    Next Cont
    COMANDAS_DETALLES
    JUEGO_REPARACION
    EXISTENCIAS
End Sub
Private Sub cmdProcesar_Click()
    Descuenta
End Sub
Private Sub Form_Activate()
    COMANDAS_DETALLES
    JUEGO_REPARACION
    EXISTENCIAS
End Sub
Private Sub Form_Load()
    Const sPathBase As String = "LINUX"
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    With Me.adcComandas_Detalles
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "COMANDAS_DETALLES"
    End With
    With Me.adcJuegos_Reparacion
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "JUEGO_REPARACION"
    End With
    With Me.adcExistencias
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "EXISTENCIAS"
    End With
    Set Me.txtId_Reparacion.DataSource = Me.adcJuegos_Reparacion
    Me.txtId_Reparacion.DataField = "ID_REPARACION"
    Set Me.txtId_Producto2.DataSource = Me.adcJuegos_Reparacion
    Me.txtId_Producto2.DataField = "ID_PRODUCTO"
    Set Me.txtCantidad2.DataSource = Me.adcJuegos_Reparacion
    Me.txtCantidad2.DataField = "CANTIDAD"
    Set Me.txtId_ProductoJr2.DataSource = Me.adcExistencias
    Me.txtId_ProductoJr2.DataField = "ID_PRODUCTO"
    Set Me.txtCantidadJr2.DataSource = Me.adcExistencias
    Me.txtCantidadJr2.DataField = "CANTIDAD"
    Set Me.txtSucursal.DataSource = Me.adcExistencias
    Me.txtSucursal.DataField = "SUCURSAL"
End Sub
Sub COMANDAS_DETALLES()
    sqlComDet = "SELECT * FROM COMANDAS_DETALLES WHERE REVISADO = 1 AND PROCESADO = 0"
    Set tRs = cnn.Execute(sqlComDet)
    With tRs
            Me.lvwComandas_Detalles.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwComandas_Detalles.ListItems.Add(, , .Fields("COMANDA"))
                tLi.SubItems(1) = .Fields("ARTICULO")
                tLi.SubItems(2) = RTrim(.Fields("ID_PRODUCTO"))
                tLi.SubItems(3) = .Fields("CANTIDAD")
                tLi.SubItems(4) = .Fields("LLEGO")
                tLi.SubItems(5) = .Fields("CANTIDAD_NO")
                tLi.SubItems(6) = .Fields("REVISADO")
                tLi.SubItems(7) = .Fields("PROCESADO")
                tLi.SubItems(8) = .Fields("DESCONTADO")
                tLi.SubItems(9) = .Fields("NO_SIRVIO")
                .MoveNext
            Loop
    End With
End Sub
Sub JUEGO_REPARACION()
    sqlJueRep = "SELECT * FROM JUEGO_REPARACION"
    Set tRs = cnn.Execute(sqlJueRep)
    With tRs
            Me.lvwJuegos_Reparacion.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwJuegos_Reparacion.ListItems.Add(, , .Fields("ID_REPARACION"))
                tLi.SubItems(1) = .Fields("ID_PRODUCTO")
                tLi.SubItems(2) = .Fields("CANTIDAD")
                .MoveNext
            Loop
    End With
End Sub
Sub EXISTENCIAS()
    sqlExc = "SELECT * FROM EXISTENCIAS"
    Set tRs = cnn.Execute(sqlExc)
    With tRs
            Me.lvwExistencias.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwExistencias.ListItems.Add(, , .Fields("ID_EXISTENCIA"))
                tLi.SubItems(1) = RTrim(.Fields("ID_PRODUCTO"))
                tLi.SubItems(2) = .Fields("CANTIDAD")
                tLi.SubItems(3) = RTrim(.Fields("SUCURSAL"))
                .MoveNext
            Loop
    End With
End Sub

