VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComandas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "COMANDAS"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtID_User 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3600
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   375
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.CheckBox chkPedSuc 
      Caption         =   "Pedido Sucursal"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6600
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc adcComDet 
      Height          =   330
      Left            =   10440
      Top             =   5880
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
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
   Begin MSAdodcLib.Adodc adcComandas 
      Height          =   330
      Left            =   240
      Top             =   6840
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
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
   Begin MSAdodcLib.Adodc adcCliente 
      Height          =   330
      Left            =   5400
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
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
   Begin VB.TextBox txtCliente 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   8880
      TabIndex        =   7
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox txtID_Producto 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   6375
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      MaxLength       =   8
      TabIndex        =   10
      Text            =   "1"
      Top             =   6120
      Width           =   1215
   End
   Begin VB.VScrollBar vsbCantidad 
      Height          =   495
      Left            =   1440
      Min             =   1
      TabIndex        =   11
      Top             =   6600
      Value           =   1
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox txtDescripcion 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   6120
      Width           =   4815
   End
   Begin MSAdodcLib.Adodc adcAlmacen3 
      Height          =   330
      Left            =   5400
      Top             =   2880
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
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
   Begin MSComctlLib.ListView lvwComandas 
      Height          =   4935
      Left            =   6960
      TabIndex        =   12
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8705
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
      Enabled         =   0   'False
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NO"
         Object.Width           =   1270
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "CLAVE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "DESCRIPCIÓN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "CANTIDAD"
         Object.Width           =   1270
      EndProperty
   End
   Begin MSComctlLib.ListView lvwClientes 
      Height          =   2055
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3625
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
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "NOMBRE"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "NOMBRE COMERCIAL"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ListView lvwProducto 
      Height          =   2055
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3625
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
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "CLAVE"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "DESCRIPCION"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.TextBox txtFecha 
      Height          =   285
      Left            =   240
      TabIndex        =   16
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtID_Cliente 
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtComanda 
      Height          =   285
      Left            =   240
      TabIndex        =   19
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtProd 
      Height          =   285
      Left            =   240
      TabIndex        =   20
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCom 
      Height          =   285
      Left            =   240
      TabIndex        =   21
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCant 
      Height          =   285
      Left            =   240
      TabIndex        =   22
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtID_Prod 
      Height          =   285
      Left            =   240
      TabIndex        =   23
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adcUsuarios 
      Height          =   330
      Left            =   4080
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
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
   Begin VB.Label Label4 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Seleccione aquí el cliente..."
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Descripción del producto..."
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione aquí el producto..."
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   2535
   End
End
Attribute VB_Name = "frmComandas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim Cont As Integer
Dim xA As ListItem
Dim sqlCliente As String
Dim sqlProducto As String
Dim tRs As Recordset
Dim tLi As ListItem
Dim bBCli As Boolean, bBAgr As Boolean
Private Sub chkPedSuc_Click()
On Error GoTo ManejaError
    If Me.chkPedSuc.Value = 1 Then
        Me.txtId_Cliente.Text = 0
        Me.txtCliente.Text = ""
        Me.txtCliente.Enabled = False
        Me.lvwClientes.Enabled = False
        Me.cmdNuevo.Enabled = False
    Else
        Me.txtCliente.Enabled = True
        Me.lvwClientes.Enabled = True
        Me.txtCliente.SetFocus
        Me.cmdNuevo.Enabled = True
        bBCli = False
    End If
End Sub
Private Sub cmdAgregar_Click()
On Error GoTo ManejaError
    If Trim(Me.txtId_Producto.Text) = "" Then
        MsgBox "Seleccione el articulo que desea pedir.", vbInformation, "MENSAJE DEL SISTEMA"
        Me.txtId_Producto.SetFocus
    Else
        Cont = Me.lvwComandas.ListItems.Count + 1
        Set xA = lvwComandas.ListItems.Add(, , Cont)
        xA.SubItems(1) = Trim(Me.txtId_Producto.Text)
        xA.SubItems(2) = Trim(Me.txtDescripcion.Text)
        xA.SubItems(3) = Val(Me.txtCantidad.Text)
        Me.lvwComandas.Enabled = True
        Me.cmdBorrar.Enabled = False
        bBAgr = True
        Me.txtId_Producto.Text = ""
        Me.txtDescripcion.Text = ""
        Me.txtId_Producto.SetFocus
    End If
End Sub
Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub cmdBorrar_Click()
On Error GoTo ManejaError
    If Me.lvwComandas.ListItems.Count <= 0 Then
        MsgBox "No hay campos que borrar.", vbInformation, "MENSAJE DEL SISTEMA"
        bBCli = False
        bBAgr = False
        Exit Sub
    End If
    If Me.lvwComandas.SelectedItem.Selected = True Then
        lvwComandas.ListItems.Remove (lvwComandas.SelectedItem.Index)
        Dim NR As Integer
        Dim Co As Integer
        NR = Me.lvwComandas.ListItems.Count
        For Co = 1 To NR
             Me.lvwComandas.ListItems.Item(Co) = Co
        Next Co
    Else
        MsgBox "Seleccione el campo que desea borrar.", vbInformation, "MENSAJE DEL SISTEMA"
    End If
End Sub
Private Sub cmdBorrar_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub cmdEnviar_Click()
On Error GoTo ManejaError
    Dim cID_USER As String
    Dim sqlUser As String
    If lvwComandas.ListItems.Count = 0 Then
        MsgBox "            No ha dado articulos a esta comanda            "
        Exit Sub
    End If
    If Me.chkPedSuc.Value = 1 Then
        Set Me.txtID_User.DataSource = Me.adcUsuarios
        sqlUser = Trim(Me.txtID_User.Text)
        sqlUser = "SELECT ID_USUARIO FROM USUARIOS WHERE ID_USUARIO LIKE '%" & sqlUser & "%'"
        Set tRs = cnn.Execute(sqlUser)
        With tRs
                Do While Not .EOF
                    cID_USER = .Fields("ID_USUARIO") & ""
                    .MoveNext
                Loop
        End With
        If Trim(Me.txtID_User.Text) = cID_USER Then
            GUARDAR_PEDIDO
        Else
            MsgBox "¡USUARIO INCORRECTO!", vbCritical, "MENSAJE DEL SISTEMA"
            Me.txtID_User.SetFocus
        End If
    Else
        If Val(Me.txtId_Cliente.Text) = 0 Then
            MsgBox "SELECCIONE EL CLIENTE DE LA LISTA, DELO DE ALTA DE SER NECESARIO. SI ES UN PEDIDO DE LA SUCURSAL SELECCIONE LA CASILLA DE 'PEDIDOS DE LA SUCURSAL'", vbExclamation, "AVISO DEL SISTEMA"
        Else
            Set Me.txtID_User.DataSource = Me.adcUsuarios
            sqlUser = Trim(Me.txtID_User.Text)
            sqlUser = "SELECT ID_USUARIO FROM USUARIOS WHERE ID_USUARIO LIKE '%" & sqlUser & "%'"
            Set tRs = cnn.Execute(sqlUser)
            With tRs
                    Do While Not .EOF
                        cID_USER = .Fields("ID_USUARIO") & ""
                        .MoveNext
                    Loop
            End With
            If Trim(Me.txtID_User.Text) = cID_USER Then
                GUARDAR_PEDIDO
                Me.txtCliente.SetFocus
            Else
                MsgBox "¡USUARIO INCORRECTO!", vbCritical, "MENSAJE DEL SISTEMA"
                Me.txtID_User.SetFocus
            End If
        End If
    End If
End Sub
Private Sub cmdEnviar_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub cmdNuevo_Click()
On Error GoTo ManejaError
    AltaClien.Show vbModal
    
End Sub

Private Sub cmdSalir_Click()
On Error GoTo ManejaError
    Unload Me
    
End Sub

Private Sub Form_Load()
On Error GoTo ManejaError
    txtID_User.Text = Menu.Text1(0).Text
    txtID_User.Enabled = False
    Const sPathBase As String = "LINUX"
    With Me.adcAlmacen3
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "ALMACEN3"
    End With
    With Me.adcCliente
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "CLIENTE"
    End With
    With Me.adcUsuarios
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "USUARIOS"
    End With
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
End Sub

Private Sub lvwClientes_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    If bBCli = True And bBAgr = True Then
        MsgBox "SOLO PUEDE HACER UNA COMANDA POR CLIENTE", vbInformation, "MENSAJE DEL SISTEMA"
    Else
        Me.txtId_Cliente.Text = Item
        Me.txtCliente.Text = Item.SubItems(1)
        bBCli = True
    End If
End Sub
Private Sub lvwClientes_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.txtId_Cliente.Text = Me.lvwClientes.SelectedItem
        Me.txtCliente.Text = Me.lvwClientes.SelectedItem.SubItems(1)
        bBCli = True
        Me.txtId_Producto.SetFocus
    End If
End Sub
Private Sub lvwComandas_Click()
On Error GoTo ManejaError
        Me.cmdBorrar.Enabled = True
End Sub
Private Sub lvwComandas_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub lvwProducto_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Me.txtId_Producto.Text = Item.SubItems(1)
    Me.txtDescripcion.Text = Item.SubItems(2)
End Sub
Private Sub lvwProducto_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.txtId_Producto.Text = Me.lvwProducto.SelectedItem.SubItems(1)
        Me.txtDescripcion.Text = Me.lvwProducto.SelectedItem.SubItems(2)
        Me.txtCantidad.SetFocus
    End If
End Sub
Private Sub TxtCantidad_GotFocus()
On Error GoTo ManejaError
    Me.txtCantidad.SelStart = 0
    Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = vbKeyReturn Then
        Me.cmdAgregar.Value = True
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

Private Sub txtCliente_GotFocus()
On Error GoTo ManejaError
    Me.txtCliente.SelStart = 0
    Me.txtCliente.SelLength = Len(Me.txtCliente.Text)
End Sub
Private Sub txtCliente_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Traer_Clientes
        Me.lvwClientes.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
    If bBCli = True And bBAgr = True Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub txtId_Producto_GotFocus()
On Error GoTo ManejaError
        Me.txtId_Producto.SelStart = 0
        Me.txtId_Producto.SelLength = Len(Me.txtId_Producto.Text)
End Sub
Private Sub txtId_Producto_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Traer_Productos
        Me.lvwProducto.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtID_User_GotFocus()
On Error GoTo ManejaError
    Me.txtID_User.SelStart = 0
    Me.txtID_User.SelLength = Len(Me.txtID_User.Text)
End Sub
Private Sub txtID_User_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
Private Sub vsbCantidad_Change()
On Error GoTo ManejaError
    Me.txtCantidad.Text = Me.vsbCantidad.Value
End Sub
Private Sub vsbCantidad_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
Sub GUARDAR_PEDIDO()
On Error GoTo ManejaError
    Dim nCliente As Integer
    Dim nUsuario As Integer
    Dim nComanda As Integer
        If MsgBox("¿SEGURO QUE DESEA ENVIAR EL PEDIDO?", vbYesNo + vbDefaultButton1 + vbQuestion, "MENSAJE DEL SISTEMA") = vbYes Then
            nCliente = Val(Me.txtId_Cliente.Text)
            nUsuario = Val(Me.txtID_User.Text)
            With Me.adcComandas
                .ConnectionString = _
                "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
                "Data Source=LINUX;"
                .RecordSource = "COMANDAS"
            End With
            Set Me.txtComanda.DataSource = Me.adcComandas
            Set Me.txtFecha.DataSource = Me.adcComandas
            Set Me.txtId_Cliente.DataSource = Me.adcComandas
            Set Me.txtUser.DataSource = Me.adcComandas
            Me.txtComanda.DataField = "COMANDA"
            Me.txtFecha.DataField = "FECHA"
            Me.txtId_Cliente.DataField = "CLIENTE"
            Me.txtUser.DataField = "USUARIO"
            Me.adcComandas.Recordset.AddNew
            Me.txtFecha.Text = Date
            Me.txtId_Cliente.Text = nCliente
            Me.txtUser.Text = nUsuario
            Me.adcComandas.Recordset.Update
            Me.adcComandas.Recordset.MoveNext
            Me.adcComandas.Recordset.MovePrevious
            nComanda = Me.txtComanda.Text
            With Me.adcComDet
                .ConnectionString = _
                "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
                "Data Source=LINUX;"
                .RecordSource = "COMANDAS_DETALLES"
            End With
            Set Me.txtCom.DataSource = Me.adcComDet
            Set Me.txtProd.DataSource = Me.adcComDet
            Set Me.txtID_Prod.DataSource = Me.adcComDet
            Set Me.txtCant.DataSource = Me.adcComDet
            Me.txtCom.DataField = "COMANDA"
            Me.txtProd.DataField = "ARTICULO"
            Me.txtID_Prod.DataField = "ID_PRODUCTO"
            Me.txtCant.DataField = "CANTIDAD"
            Dim NoRe As Integer
            Dim C As Integer
            NoRe = Me.lvwComandas.ListItems.Count
            For C = 1 To NoRe
                Me.adcComDet.Recordset.AddNew
                Me.txtCom.Text = nComanda
                Me.txtProd.Text = Me.lvwComandas.ListItems.Item(C)
                Me.txtID_Prod.Text = Me.lvwComandas.ListItems.Item(C).SubItems(1)
                Me.txtCant.Text = Me.lvwComandas.ListItems.Item(C).SubItems(3)
                Me.adcComDet.Recordset.Update
            Next C
            If Me.chkPedSuc.Value = 1 Then
                MsgBox "EL NUMERO DE SU PEDIDO ES: " & nComanda, vbDefaultButton1 + vbExclamation + vbCritical, "ENVIANDO PEDIDO..."
                Imprimir_Ticket
            Else
                MsgBox "IDENTIFIQUE TODOS LOS CARTUCHOS DEL CLIENTE CON EL NUMERO: " & nComanda, vbDefaultButton1 + vbExclamation + vbCritical, "ENVIANDO PEDIDO..."
                Imprimir_Ticket
            End If
            Me.lvwComandas.ListItems.Clear
        End If
End Sub

Sub Traer_Clientes()
On Error GoTo ManejaError
    sqlCliente = Trim(Me.txtCliente.Text)
    sqlCliente = "SELECT ID_CLIENTE, NOMBRE, NOMBRE_COMERCIAL FROM CLIENTE WHERE NOMBRE LIKE '%" & sqlCliente & "%' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sqlCliente)
    With tRs
            Me.lvwClientes.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwClientes.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                tLi.SubItems(1) = .Fields("NOMBRE") & ""
               tLi.SubItems(2) = .Fields("NOMBRE_COMERCIAL") & ""
                .MoveNext
            Loop
    End With

End Sub

Sub Traer_Productos()
On Error GoTo ManejaError
    sqlProducto = Trim(Me.txtId_Producto.Text)
    sqlProducto = "SELECT ID_PRODUCTO, DESCRIPCION FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & sqlProducto & "%' ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sqlProducto)
    With tRs
            Me.lvwProducto.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwProducto.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                tLi.SubItems(1) = .Fields("ID_PRODUCTO") & ""
                tLi.SubItems(2) = .Fields("DESCRIPCION") & ""
                .MoveNext
            Loop
    End With

End Sub
Sub Imprimir_Ticket()
On Error GoTo ManejaError
    Printer.Print "   ACTITUD POSITIVA EN TONER S DE RL MI"
    Printer.Print "                    R.F.C. APT- 040201-KA5"
    Printer.Print "ORTIZ DE CAMPOS No. 1308 COL. SAN FELIPE"
    Printer.Print "      CHIHUAHUA, CHIHUAHUA C.P. 31203"
    Printer.Print "FECHA : " & Date
    Printer.Print "SUCURSAL : " & Menu.Text4(0).Text
    Printer.Print "No. DE COMANDA : " & Me.txtCom.Text
    Printer.Print "ATENDIDO POR : " & Menu.Text1(1).Text
    Printer.Print "CLIENTE : " & txtCliente.Text
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                          Comanda de Recarga"
    Printer.Print "--------------------------------------------------------------------------------"
    Dim NRegistros As Integer
    NRegistros = Me.lvwComandas.ListItems.Count
    Dim CoN As Integer
    Dim POSY As Integer
    POSY = 2200
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3000
    Printer.Print "Cant."
    For CoN = 1 To NRegistros
        POSY = POSY + 200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print Me.lvwComandas.ListItems(CoN).SubItems(1)
        Printer.CurrentY = POSY
        Printer.CurrentX = 2900
        Printer.Print Me.lvwComandas.ListItems(CoN).SubItems(3)
    Next CoN
    Printer.Print ""
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "NOTA: No debe haber ningun cobro hasta la entrega del"
    Printer.Print "       cartucho lleno"
    Printer.Print "               GRACIAS POR SU COMPRA"
    Printer.Print "           PRODUCTO 100% GARANTIZADO"
    Printer.Print ""
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.EndDoc
End Sub

