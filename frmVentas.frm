VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VENTAS"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtId_Cliente 
      Height          =   285
      Left            =   9000
      TabIndex        =   10
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "CLIENTE"
      Height          =   2175
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   11295
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar"
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
         Left            =   9720
         Picture         =   "frmVentas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuscar 
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
         Left            =   9720
         Picture         =   "frmVentas.frx":29D2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton opnNombre 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   6720
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opnClave 
         Caption         =   "Clave"
         Height          =   255
         Left            =   5760
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwClientes 
         Height          =   1215
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2143
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CLAVE"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NOMBRE"
            Object.Width           =   9878
         EndProperty
      End
      Begin VB.TextBox txtNombreCliente 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   9960
      Picture         =   "frmVentas.frx":53A4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   1335
   End
   Begin TabDlg.SSTab sstVentas 
      Height          =   4335
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "COMANDA"
      TabPicture(0)   =   "frmVentas.frx":7D76
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lvwNuevaComanda"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lvwProductosComanda"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAceptarComanda"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdBuscarComanda"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "opnCodigoComanda"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "opnClaveComanda"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "opnDescripcionComanda"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdQuitarComanda"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtProductoComanda"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCantidadComanda"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdAgregarComanda"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      Begin VB.CommandButton cmdAgregarComanda 
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
         Left            =   9720
         Picture         =   "frmVentas.frx":7D92
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtCantidadComanda 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   9720
         TabIndex        =   18
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtProductoComanda 
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   5295
      End
      Begin VB.CommandButton cmdQuitarComanda 
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
         Left            =   9720
         Picture         =   "frmVentas.frx":A764
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2880
         Width           =   1335
      End
      Begin VB.OptionButton opnDescripcionComanda 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   6720
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton opnClaveComanda 
         Caption         =   "Clave"
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   720
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opnCodigoComanda 
         Caption         =   "Codigo barras"
         Height          =   255
         Left            =   8040
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscarComanda 
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
         Left            =   9720
         Picture         =   "frmVentas.frx":D136
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdAceptarComanda 
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
         Height          =   375
         Left            =   9720
         Picture         =   "frmVentas.frx":FB08
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3600
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvwProductosComanda 
         Height          =   1215
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2143
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CLAVE"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPCIÓN"
            Object.Width           =   8114
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "GANANCIA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "PRECIO COSTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "PRECIO"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lvwNuevaComanda 
         Height          =   1215
         Left            =   240
         TabIndex        =   21
         Top             =   2760
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2143
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
            Text            =   "CLAVE"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPCIÓN"
            Object.Width           =   8114
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "CANTIDAD"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   9720
         TabIndex        =   24
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Producto"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Comanda"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2520
         Width           =   1335
      End
   End
   Begin VB.Label lblEstado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   7080
      Width           =   9255
   End
End
Attribute VB_Name = "frmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As Recordset

Private Sub cmdAceptarComanda_Click()

On Error GoTo ManejaError

    If Puede_Guardar Then
    
        Dim NoRe As Integer
        Dim Cont As Integer
        'Dim dFecha As String
        Dim nComanda As Integer
        Dim cTipo As String
        
        'Hora del sistema.
        'sqlQuery = "SELECT getdate() as FechaHoraSistema"
        'Set tRs = cnn.Execute(sqlQuery)
        'dFecha = tRs.Fields("FechaHoraSistema")
    
        sqlQuery = "INSERT INTO COMANDAS_2 (FECHA_INICIO, ID_CLIENTE, ID_AGENTE, ID_SUCURSAL) VALUES ('" & Date & "', " & Me.txtId_Cliente.Text & ", " & Menu.Text1(0).Text & ", " & Menu.Text1(5).Text & ")"
        cnn.Execute (sqlQuery)
        
        Me.lblEstado.Caption = "Enviando"
        Me.lblEstado.ForeColor = vbBlack
        DoEvents
        
        sqlQuery = "SELECT TOP 1 ID_COMANDA FROM COMANDAS_2 ORDER BY ID_COMANDA DESC"
        Set tRs = cnn.Execute(sqlQuery)
        nComanda = tRs.Fields("ID_COMANDA")
    
        Me.lblEstado.Caption = Me.lblEstado.Caption & " comanda " & nComanda
        Me.lblEstado.ForeColor = vbBlack
        DoEvents
                
        NoRe = Me.lvwNuevaComanda.ListItems.Count
        
        For Cont = 1 To NoRe
            If Mid(Me.lvwNuevaComanda.ListItems.Item(Cont), 3, 1) = "T" Then
                cTipo = "T" 'Toner
            ElseIf Mid(Me.lvwNuevaComanda.ListItems.Item(Cont), 3, 1) = "I" Then
                cTipo = "I" 'Tinta
            Else
                cTipo = "X" 'Error
            End If
            sqlQuery = "INSERT INTO COMANDAS_DETALLES_2 (ID_COMANDA, ARTICULO, ID_PRODUCTO, CANTIDAD, TIPO) VALUES (" & nComanda & ", " & Cont & ", '" & Me.lvwNuevaComanda.ListItems.Item(Cont) & "', " & Me.lvwNuevaComanda.ListItems.Item(Cont).SubItems(2) & ", '" & cTipo & "')"
            cnn.Execute (sqlQuery)
            
            Me.lblEstado.Caption = Me.lblEstado.Caption & ", producto " & Cont & " de " & NoRe
            Me.lblEstado.ForeColor = vbBlack
            DoEvents
            
        Next Cont
        
        Imprimir_Ticket (nComanda)
        Imprimir_Ticket (nComanda)
        
        Borrar_Campos
    
    End If

Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If

End Sub

Private Sub cmdAgregarComanda_Click()

    If Puede_Agregar_Comanda Then
        Set tLi = Me.lvwNuevaComanda.ListItems.Add(, , Me.lvwProductosComanda.SelectedItem)
        tLi.SubItems(1) = Me.lvwProductosComanda.SelectedItem.SubItems(1)
        tLi.SubItems(2) = Me.txtCantidadComanda.Text
        Me.lblEstado.Caption = ""
        Me.txtProductoComanda.SetFocus
    End If
    
End Sub

Private Sub cmdBuscar_Click()

    If Puede_Buscar Then
        If Hay_Clientes(Trim(Me.txtNombreCliente.Text)) Then
            Llenar_Lista_Clientes Trim(Me.txtNombreCliente.Text)
        End If
    End If
    
End Sub

Private Sub cmdBuscarComanda_Click()

    If Puede_Buscar_Producto Then
        'If Hay_Productos(Trim(Me.txtProductoComanda.Text)) Then
            Me.lblEstado.Caption = "Buscando"
            Me.lblEstado.ForeColor = vbBlack
            DoEvents
            Llenar_Lista_Productos Trim(Me.txtProductoComanda.Text)
        'End If
    End If
                
End Sub

Private Sub cmdQuitarComanda_Click()

    If Me.lvwNuevaComanda.ListItems.Count <> 0 Then
        If Me.lvwNuevaComanda.SelectedItem.Selected Then
            Me.lvwNuevaComanda.ListItems.Remove (Me.lvwNuevaComanda.SelectedItem.Index)
        End If
    End If

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub cmdSeleccionar_Click()

    If Puede_Seleccionar_Cliente Then
        Me.txtNombreCliente.Text = Me.lvwClientes.SelectedItem.SubItems(1)
        Me.txtId_Cliente.Text = Me.lvwClientes.SelectedItem
        Me.txtProductoComanda.SetFocus
    End If
    
End Sub

Private Sub Form_Load()

On Error GoTo ManejaError

    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Dim sPathBase As String
    sPathBase = "LINUX"
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
    
End Sub

Sub Reconexion()
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Const sPathBase As String = "LINUX"
    With cnn
        'If BanCnn = False Then
            '.Close
            'BanCnn = True
        'End If
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
        'BanCnn = False
        MsgBox "LA CONEXION SE RESABLECIO CON EXITO. PUEDE CONTINUAR CON SU TRABAJO.", vbInformation, "MENSAJE DEL SISTEMA"
    End With
Exit Sub
ManejaError:
    MsgBox Err.Number & Err.Description
    Err.Clear
    If MsgBox("NO PUDIMOS RESTABLECER LA CONEXIÓN, ¿DESEA REINTENTARLO?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
End Sub

Function Hay_Clientes(cCliente As String) As Boolean
    
On Error GoTo ManejaError
    
    Me.lblEstado.Caption = "Buscando"
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    
    If Me.opnNombre.Value = True Then
        sqlQuery = "SELECT COUNT(ID_CLIENTE)ID_CLIENTE FROM CLIENTE WHERE NOMBRE LIKE '%" & cCliente & "%'"
    Else
        sqlQuery = "SELECT COUNT(ID_CLIENTE)ID_CLIENTE FROM CLIENTE WHERE ID_CLIENTE = " & Val(cCliente)
    End If
    
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_CLIENTE") <> 0 Then
            Hay_Clientes = True
            Me.lblEstado.Caption = "Se encontraron " & .Fields("ID_CLIENTE") & " clientes"
            Me.lblEstado.ForeColor = vbBlue
            DoEvents
        Else
            Hay_Clientes = False
            Me.lblEstado.Caption = "No se encontraron clientes"
            Me.lblEstado.ForeColor = vbRed
            Me.txtNombreCliente.SetFocus
        End If
    End With

Exit Function
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Function

Function Llenar_Lista_Clientes(cCliente As String)

On Error GoTo ManejaError

        If Me.opnNombre.Value = True Then
            sqlQuery = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE NOMBRE LIKE '%" & cCliente & "%'"
        Else
            sqlQuery = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE ID_CLIENTE = " & Val(cCliente)
        End If
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            Me.lvwClientes.ListItems.Clear
            Do While Not .EOF
                Set tLi = lvwClientes.ListItems.Add(, , .Fields("ID_CLIENTE"))
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE")
                    .MoveNext
            Loop
        End With

Exit Function
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
    
End Function

Private Sub lvwClientes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Me.cmdSeleccionar.Value = True
    End If
    
End Sub

Private Sub lvwProductosComanda_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Me.lvwProductosComanda.SelectedItem.Selected Then
            Me.txtProductoComanda.Text = Me.lvwProductosComanda.SelectedItem.SubItems(1)
            Me.txtCantidadComanda.SetFocus
        End If
    End If

End Sub

Private Sub txtCantidadComanda_GotFocus()

    Me.txtCantidadComanda.SelStart = 0
    Me.txtCantidadComanda.SelLength = Len(Me.txtCantidadComanda.Text)

End Sub

Private Sub txtCantidadComanda_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Me.cmdAgregarComanda.Value = True
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

Private Sub txtNombreCliente_GotFocus()

    Me.txtNombreCliente.SelStart = 0
    Me.txtNombreCliente.SelLength = Len(Me.txtNombreCliente.Text)

End Sub

Private Sub txtNombreCliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Me.lvwClientes.SetFocus
        Me.cmdBuscar.Value = True
    End If
    
End Sub

Function Puede_Buscar() As Boolean
    
    If Trim(Me.txtNombreCliente.Text) = "" Then
        Puede_Buscar = False
        Me.lblEstado.Caption = "Introsusca el cliente"
        Me.lblEstado.ForeColor = vbRed
        Me.txtNombreCliente.SetFocus
        Exit Function
    End If
    
    Puede_Buscar = True
    
End Function

Function Puede_Buscar_Producto() As Boolean
    
    If Trim(Me.txtProductoComanda.Text) = "" Then
        Puede_Buscar_Producto = False
        Me.lblEstado.Caption = "Introsusca el producto"
        Me.lblEstado.ForeColor = vbRed
        Me.txtProductoComanda.SetFocus
        Exit Function
    End If
    
    Puede_Buscar_Producto = True
    
End Function

Function Hay_Productos(cProducto As String) As Boolean
    
On Error GoTo ManejaError

    Me.lblEstado.Caption = "Buscando"
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    
    If Me.opnClaveComanda.Value = True Then
        'sqlQuery = "SELECT COUNT(ID_PRODUCTO)ID_PRODUCTO FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & cProducto & "%' AND (ID_PRODUCTO LIKE '__T%' OR ID_PRODUCTO LIKE '__I%')"
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE J.ID_REPARACION LIKE '%" & cProducto & "%' ORDER BY J.ID_REPARACION"
        'InputBox "", "", sqlQuery
    ElseIf Me.opnDescripcionComanda.Value = True Then
        'sqlQuery = "SELECT COUNT(ID_PRODUCTO)ID_PRODUCTO FROM ALMACEN3 WHERE DESCRIPCION LIKE '%" & cProducto & "%'AND (ID_PRODUCTO LIKE '__T%' OR ID_PRODUCTO LIKE '__I%')"
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE A.DESCRIPCION LIKE '%" & cProducto & "%' ORDER BY J.ID_REPARACION"
    Else
        'sqlQuery = "SELECT COUNT(ID_PRODUCTO)ID_PRODUCTO FROM ALMACEN3 WHERE ID_PRODUCTO = " & cProducto & " AND (ID_PRODUCTO LIKE '__T%' OR ID_PRODUCTO LIKE '__I%')"
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE J.ID_REPARACION = '" & cProducto & "' ORDER BY J.ID_REPARACION"
    End If
    
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not .EOF Then
            Hay_Productos = True
            Me.lblEstado.Caption = ""
            Me.lblEstado.ForeColor = vbBlue
            DoEvents
        Else
            Hay_Productos = False
            Me.lblEstado.Caption = "No se encontraron productos"
            Me.lblEstado.ForeColor = vbRed
            Me.txtProductoComanda.SetFocus
        End If
    End With
    
Exit Function
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Function

Private Sub txtProductoComanda_GotFocus()

    Me.txtProductoComanda.SelStart = 0
    Me.txtProductoComanda.SelLength = Len(Me.txtProductoComanda.Text)

End Sub

Private Sub txtProductoComanda_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Me.lvwProductosComanda.SetFocus
        Me.cmdBuscarComanda.Value = True
    End If

End Sub

Function Llenar_Lista_Productos(cProducto As String)

On Error GoTo ManejaError

    If Me.opnClaveComanda.Value = True Then
        'sqlQuery = "SELECT ID_PRODUCTO, DESCRIPCION, GANANCIA, PRECIO_COSTO  FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & cProducto & "%' AND (ID_PRODUCTO LIKE '__T%' OR ID_PRODUCTO LIKE '__I%')"
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION, A.DESCRIPCION, A.GANANCIA, A.PRECIO_COSTO FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE J.ID_REPARACION LIKE '%" & cProducto & "%' ORDER BY J.ID_REPARACION"
    ElseIf Me.opnDescripcionComanda.Value = True Then
        'sqlQuery = "SELECT ID_PRODUCTO, DESCRIPCION, GANANCIA, PRECIO_COSTO  FROM ALMACEN3 WHERE DESCRIPCION LIKE '%" & cProducto & "%' AND (ID_PRODUCTO LIKE '__T%' OR ID_PRODUCTO LIKE '__I%')"
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION, A.DESCRIPCION, A.GANANCIA, A.PRECIO_COSTO FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE A.DESCRIPCION LIKE '%" & cProducto & "%' ORDER BY J.ID_REPARACION"
    Else
        'sqlQuery = "SELECT ID_PRODUCTO, DESCRIPCION, GANANCIA, PRECIO_COSTO  FROM ALMACEN3 WHERE ID_PRODUCTO = " & cProducto & " AND (ID_PRODUCTO LIKE '__T%' OR ID_PRODUCTO LIKE '__I%')"
        sqlQuery = "SELECT DISTINCT J.ID_REPARACION, A.DESCRIPCION, A.GANANCIA, A.PRECIO_COSTO FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE J.ID_REPARACION = '" & cProducto & "' ORDER BY J.ID_REPARACION"
    End If
        'InputBox "", "", sqlQuery
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            If Not (.EOF Or .BOF) Then
                Me.lblEstado.Caption = ""
                Me.lvwProductosComanda.ListItems.Clear
                Do While Not .EOF
                    Set tLi = lvwProductosComanda.ListItems.Add(, , .Fields("ID_REPARACION"))
                        If Not IsNull(.Fields("DESCRIPCION")) Then tLi.SubItems(1) = .Fields("DESCRIPCION")
                        If Not IsNull(.Fields("GANANCIA")) Then tLi.SubItems(2) = .Fields("GANANCIA")
                        If Not IsNull(.Fields("PRECIO_COSTO")) Then tLi.SubItems(3) = .Fields("PRECIO_COSTO")
                        If Not IsNull(.Fields("PRECIO_COSTO")) And Not IsNull(.Fields("GANANCIA")) Then
                            tLi.SubItems(4) = Format((1 + CDbl(.Fields("GANANCIA"))) * CDbl(.Fields("PRECIO_COSTO")), "0.00")
                        Else
                            MsgBox "BASE DE DATOS CORRUPTA", vbCritical, "ERROR GRAVE"
                        End If
                        .MoveNext
                Loop
            Else
            Me.lblEstado.Caption = "No se encontraron productos"
            Me.lblEstado.ForeColor = vbRed
            Me.txtProductoComanda.SetFocus
            End If
        End With

Exit Function
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
    
End Function

Function Puede_Agregar_Comanda() As Boolean

    If Me.lvwProductosComanda.ListItems.Count = 0 Then
        Puede_Agregar_Comanda = False
        Me.lblEstado.Caption = "Introsusca el producto"
        Me.lblEstado.ForeColor = vbRed
        Me.txtProductoComanda.SetFocus
        Exit Function
    End If
    
    If Trim(Val(Me.txtCantidadComanda.Text)) = 0 Then
        Puede_Agregar_Comanda = False
        Me.lblEstado.Caption = "Introsusca la cantidad"
        Me.lblEstado.ForeColor = vbRed
        Me.txtCantidadComanda.SetFocus
        Exit Function
    End If
    
    Puede_Agregar_Comanda = True

End Function

Function Puede_Seleccionar_Cliente() As Boolean

    If Me.lvwClientes.ListItems.Count = 0 Then
        Puede_Seleccionar_Cliente = False
        Me.lblEstado.Caption = "Introsusca el cliente"
        Me.lblEstado.ForeColor = vbRed
        Me.txtNombreCliente.SetFocus
        Exit Function
    End If
    
    Puede_Seleccionar_Cliente = True

End Function

Function Puede_Guardar() As Boolean

    If Me.txtId_Cliente.Text = "" Then
        Puede_Guardar = False
        Me.lblEstado.Caption = "Introsusca el cliente"
        Me.lblEstado.ForeColor = vbRed
        Me.txtNombreCliente.SetFocus
        Exit Function
    End If
    
    If Me.lvwNuevaComanda.ListItems.Count = 0 Then
        Puede_Guardar = False
        Me.lblEstado.Caption = "Introsusca el producto"
        Me.lblEstado.ForeColor = vbRed
        Me.txtProductoComanda.SetFocus
        Exit Function
    End If
    
    Puede_Guardar = True
    
End Function

Function Borrar_Campos()

    Me.txtCantidadComanda.Text = "1"
    Me.txtProductoComanda.Text = ""
    Me.lvwProductosComanda.ListItems.Clear
    Me.lvwNuevaComanda.ListItems.Clear
    Me.lblEstado.Caption = ""
    
End Function

Function Imprimir_Ticket(cNoCom As Integer)

On Error GoTo ManejaError

    Printer.Print "   ACTITUD POSITIVA EN TONER S DE RL MI"
    Printer.Print "                    R.F.C. APT- 040201-KA5"
    Printer.Print "ORTIZ DE CAMPOS No. 1308 COL. SAN FELIPE"
    Printer.Print "      CHIHUAHUA, CHIHUAHUA C.P. 31203"
    Printer.Print "FECHA : " & Now
    Printer.Print "SUCURSAL : " & Menu.Text4(0).Text
    Printer.Print "No. DE COMANDA : " & cNoCom
    Printer.Print "ATENDIDO POR : " & Menu.Text1(1).Text & " " & Menu.Text1(2).Text
    Printer.Print "CLIENTE : " & Me.txtNombreCliente.Text
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                           RECARGA DE TINTA"
    Dim NRegistros As Integer
    NRegistros = Me.lvwNuevaComanda.ListItems.Count
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
        If Mid(Me.lvwNuevaComanda.ListItems.Item(CoN), 3, 1) = "I" Then
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print Me.lvwNuevaComanda.ListItems(CoN)
            Printer.CurrentY = POSY
            Printer.CurrentX = 2900
            Printer.Print Me.lvwNuevaComanda.ListItems(CoN).SubItems(2)
        End If
    Next CoN
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                           RECARGA DE TONER"
    POSY = POSY + 600
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3000
    Printer.Print "Cant."
    For CoN = 1 To NRegistros
        If Mid(Me.lvwNuevaComanda.ListItems.Item(CoN), 3, 1) = "T" Then
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print Me.lvwNuevaComanda.ListItems(CoN)
            Printer.CurrentY = POSY
            Printer.CurrentX = 2900
            Printer.Print Me.lvwNuevaComanda.ListItems(CoN).SubItems(2)
        End If
    Next CoN
    Printer.Print ""
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "               GRACIAS POR SU COMPRA"
    Printer.Print "           PRODUCTO 100% GARANTIZADO"
    Printer.Print ""
    Printer.Print "Conserve su ticket"
    Printer.Print "El cobro se hará hasta la entrega del cartucho lleno"
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.EndDoc
    
Exit Function
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Function

