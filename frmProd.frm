VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PRODUCCÓN"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox txtNoSIr 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      TabIndex        =   3
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdNoSir 
      Caption         =   "NO SIRVIO"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   5040
      TabIndex        =   13
      Top             =   7080
      Width           =   3015
   End
   Begin VB.TextBox txtProducto 
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Top             =   6960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdPedidoListo 
      Caption         =   "Quitar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdListo 
      Caption         =   "Listo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox txtNo 
      Height          =   285
      Left            =   6120
      TabIndex        =   12
      Top             =   7080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtComanda 
      Height          =   285
      Left            =   5040
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwComCli 
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "COMANDA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "FECHA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CLIENTE"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "USUARIO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ACTIVO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "REVISADO"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvwComSuc 
      Height          =   3135
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "COMANDA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "FECHA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CLIENTE"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "USUARIO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ACTIVO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "REVISADO"
         Object.Width           =   0
      EndProperty
   End
   Begin MSAdodcLib.Adodc adcComDet 
      Height          =   330
      Left            =   6240
      Top             =   7560
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
   Begin MSAdodcLib.Adodc adcCom 
      Height          =   330
      Left            =   5040
      Top             =   7560
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
   Begin MSComctlLib.ListView lvwComandas_Detalles_Cliente 
      Height          =   3135
      Left            =   3600
      TabIndex        =   9
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5530
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PEDIDO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "NO."
         Object.Width           =   1764
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
         Object.Width           =   0
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
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "CANTIDAD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "REVISADO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "PROCESADO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "DESCONTADO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "NO SIRVIO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "LISTO"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvwComandas_Detalles_Sucursal 
      Height          =   3135
      Left            =   3600
      TabIndex        =   10
      Top             =   3960
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5530
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PEDIDO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "NO."
         Object.Width           =   1764
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
         Object.Width           =   0
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
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "CANTIDAD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "REVISADO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "PROCESADO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "DESCONTADO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "NO_SIRVIO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "LISTO"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "--------------------PEDIDOS DE SUCURSALES--------------------"
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
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   10335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "--------------------PEDIDOS DE CLIENTES--------------------"
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
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frmProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim tRs As Recordset
Dim tLi As ListItem
Dim sqlComClie As String
Dim sqlComSuc As String
Dim SQL As String
Dim BanListaPed As Byte
Private Sub cmdListo_Click()
    SQL = "UPDATE COMANDAS_DETALLES SET DESCONTADO = 1 WHERE COMANDA = " & Me.txtComanda.Text & " AND ARTICULO = " & Me.txtNo.Text
    Set tRs = cnn.Execute(SQL)
    If BanLista = 1 Then
        Llena_Comanda_Sucursal
    Else
        Llena_Comanda_Cliente
    End If
    Me.cmdListo.Enabled = False
End Sub
Private Sub cmdNoSir_Click()
    If Me.lvwComandas_Detalles_Cliente.SelectedItem.SubItems(11) = "1" Then
            MsgBox "YA FUE AUMENTADO", vbExclamation, "AVISO DEL SISTEMA"
    Else
        If Me.txtNoSIr.Text = "" Then
            MsgBox "POR FAVOR, ESPECIFIQUE LA CANTIDAD QUE NO SIRVIO", vbExclamation, "AVISO DEL SISTEMA"
            Me.txtNoSIr.SetFocus
        Else
            If Val(Me.txtNoSIr.Text) > Val(Me.lvwComandas_Detalles_Cliente.SelectedItem.SubItems(6)) Then
                MsgBox "LA CANTIDAD QUE NO SIRVIO DEBE SER MENOR", vbExclamation, "AVISO DEL SISTEMA"
            Else
                SQL = "UPDATE COMANDAS_DETALLES SET NO_SIRVIO = " & Val(Me.txtNoSIr.Text) & " WHERE COMANDA = " & Me.txtComanda.Text & " AND ARTICULO = " & Me.txtNo.Text
                Set tRs = cnn.Execute(SQL)
                Llena_Comanda_Cliente
                Me.lvwComandas_Detalles_Cliente.SelectedItem.SubItems(11) = "1"
                Me.cmdNoSir.Enabled = False
                Me.txtNoSIr.Enabled = False
                frmRegAlm.Show vbModal
            End If
        End If
    End If
End Sub
Private Sub cmdPedidoListo_Click()
    SQL = "UPDATE COMANDAS SET ACTIVO = 0 WHERE COMANDA = " & Me.txtComanda.Text
    Set tRs = cnn.Execute(SQL)
    SQL = "UPDATE COMANDAS SET FECHA_FIN = " & Date & " WHERE COMANDA = " & Me.txtComanda.Text
    Set tRs = cnn.Execute(SQL)
    If BanListaPed = 0 Then
        Llena_Comanda_Cli
        Me.lvwComandas_Detalles_Cliente.ListItems.Clear
    Else
        Llena_Comanda_Suc
        Me.lvwComandas_Detalles_Sucursal.ListItems.Clear
    End If
    Me.cmdPedidoListo.Enabled = False
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    Const sPathBase As String = "LINUX"
    With Me.adcCom
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "COMANDAS"
    End With
    With Me.adcComDet
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "COMANDAS_DETALLES"
    End With
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    Llena_Comanda_Cli
    Llena_Comanda_Suc
End Sub
Private Sub lvwComandas_Detalles_Cliente_Click()
    On Error GoTo MANEJAERROR
    Me.cmdPedidoListo.Enabled = False
    Me.txtNo.Text = Me.lvwComandas_Detalles_Cliente.SelectedItem.SubItems(1)
    Me.txtComanda.Text = Me.lvwComandas_Detalles_Cliente.SelectedItem
    Me.txtProducto.Text = Me.lvwComandas_Detalles_Cliente.SelectedItem.SubItems(2)
    BanLista = 0
    Me.cmdNoSir.Enabled = True
    Me.cmdListo.Enabled = True
    Me.txtNoSIr.Enabled = True
    Exit Sub
MANEJAERROR:
    Err = 0
End Sub
Private Sub lvwComandas_Detalles_Cliente_DblClick()
    If Me.txtNoSIr.Enabled = True Then
        Me.txtNoSIr.SetFocus
    End If
End Sub
Private Sub lvwComandas_Detalles_Sucursal_Click()
    On Error GoTo MANEJAERROR
    Me.cmdPedidoListo.Enabled = False
    Me.txtNo.Text = Me.lvwComandas_Detalles_Sucursal.SelectedItem.SubItems(1)
    Me.txtComanda.Text = Me.lvwComandas_Detalles_Sucursal.SelectedItem
    Me.txtProducto.Text = Me.lvwComandas_Detalles_Sucursal.SelectedItem.SubItems(2)
    BanLista = 1
    Me.cmdNoSir.Enabled = False
    Me.cmdListo.Enabled = True
    Exit Sub
MANEJAERROR:
    Err.Clear
End Sub
Private Sub lvwComCli_Click()
    On Error GoTo MANEJAERROR
    Me.txtComanda.Text = Me.lvwComCli.SelectedItem
    Llena_Comanda_Cliente
    Me.cmdPedidoListo.Enabled = True
    Me.cmdListo.Enabled = False
    BanListaPed = 0
    Exit Sub
MANEJAERROR:
    Err.Clear
End Sub

Sub Llena_Comanda_Cliente()
    SQL = "SELECT * FROM COMANDAS_DETALLES WHERE REVISADO = 1  AND PROCESADO = 1 AND DESCONTADO = 0 AND COMANDA = " & Me.txtComanda.Text & " ORDER BY COMANDA"
    Set tRs = cnn.Execute(SQL)
    With tRs
        Me.lvwComandas_Detalles_Cliente.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwComandas_Detalles_Cliente.ListItems.Add(, , .Fields("COMANDA"))
                tLi.SubItems(1) = .Fields("ARTICULO")
                tLi.SubItems(2) = RTrim(.Fields("ID_PRODUCTO"))
                tLi.SubItems(3) = .Fields("CANTIDAD")
                tLi.SubItems(4) = .Fields("LLEGO")
                tLi.SubItems(5) = .Fields("CANTIDAD_NO")
                tLi.SubItems(6) = (.Fields("CANTIDAD") - .Fields("CANTIDAD_NO"))
                tLi.SubItems(7) = .Fields("REVISADO")
                tLi.SubItems(8) = .Fields("PROCESADO")
                tLi.SubItems(9) = .Fields("DESCONTADO")
                tLi.SubItems(10) = .Fields("NO_SIRVIO")
                .MoveNext
            Loop
    End With
End Sub
Sub Llena_Comanda_Sucursal()
    SQL = "SELECT * FROM COMANDAS_DETALLES WHERE REVISADO = 1  AND PROCESADO = 1 AND DESCONTADO = 0 AND COMANDA = " & Me.txtComanda.Text & " ORDER BY COMANDA"
    Set tRs = cnn.Execute(SQL)
    With tRs
        Me.lvwComandas_Detalles_Sucursal.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwComandas_Detalles_Sucursal.ListItems.Add(, , .Fields("COMANDA"))
                tLi.SubItems(1) = .Fields("ARTICULO")
                tLi.SubItems(2) = RTrim(.Fields("ID_PRODUCTO"))
                tLi.SubItems(3) = .Fields("CANTIDAD")
                tLi.SubItems(4) = .Fields("LLEGO")
                tLi.SubItems(5) = .Fields("CANTIDAD_NO")
                tLi.SubItems(6) = (.Fields("CANTIDAD") - .Fields("CANTIDAD_NO"))
                tLi.SubItems(7) = .Fields("REVISADO")
                tLi.SubItems(8) = .Fields("PROCESADO")
                tLi.SubItems(9) = .Fields("DESCONTADO")
                tLi.SubItems(10) = .Fields("NO_SIRVIO")
                .MoveNext
            Loop
    End With
End Sub
Private Sub lvwComSuc_Click()
    On Error GoTo MANEJAERROR
    Me.txtComanda.Text = Me.lvwComSuc.SelectedItem
    Llena_Comanda_Sucursal
    Me.cmdPedidoListo.Enabled = True
    Me.cmdListo.Enabled = False
    BanListaPed = 1
    Exit Sub
MANEJAERROR:
    Err = 0
End Sub
Private Sub txtNoSIr_GotFocus()
    Me.txtNoSIr.SelStart = 0
    Me.txtNoSIr.SelLength = Len(Me.txtNoSIr.Text)
End Sub
Private Sub txtNoSIr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdNoSir.Value = True
    Else
        Dim Valido As String
        Valido = "1234567890."
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii > 26 Then
                If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
                End If
            End If
    End If
End Sub
Sub Llena_Comanda_Cli()
    sqlComClie = "SELECT * FROM COMANDAS WHERE CLIENTE <> 0 AND REVISADO='Si' AND ACTIVO = 1 ORDER BY COMANDA"
    Set tRs = cnn.Execute(sqlComClie)
    With tRs
            Me.lvwComCli.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwComCli.ListItems.Add(, , .Fields("COMANDA"))
                tLi.SubItems(1) = .Fields("FECHA")
                tLi.SubItems(2) = .Fields("CLIENTE")
                tLi.SubItems(3) = .Fields("USUARIO")
                tLi.SubItems(4) = .Fields("ACTIVO")
                tLi.SubItems(5) = .Fields("REVISADO")
                .MoveNext
            Loop
    End With
End Sub
Sub Llena_Comanda_Suc()
    sqlComSuc = "SELECT * FROM COMANDAS WHERE CLIENTE = 0 AND REVISADO='Si' AND ACTIVO = 1 ORDER BY COMANDA"
    Set tRs = cnn.Execute(sqlComSuc)
    With tRs
        Me.lvwComSuc.ListItems.Clear
        Do While Not .EOF
            Set tLi = Me.lvwComSuc.ListItems.Add(, , .Fields("COMANDA"))
                tLi.SubItems(1) = .Fields("FECHA")
                tLi.SubItems(2) = .Fields("CLIENTE")
                tLi.SubItems(3) = .Fields("USUARIO")
                tLi.SubItems(4) = .Fields("ACTIVO")
                tLi.SubItems(5) = .Fields("REVISADO")
                .MoveNext
        Loop
    End With
End Sub
