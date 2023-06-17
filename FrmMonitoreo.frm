VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMonitoreo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Monitoreo de Documentos (Más de un mes abiertos)"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11160
      TabIndex        =   5
      Top             =   5520
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmMonitoreo.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmMonitoreo.frx":030A
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
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Documentos Abiertos"
      TabPicture(0)   =   "FrmMonitoreo.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListView2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListView1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Catálogos"
      TabPicture(1)   =   "FrmMonitoreo.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView5"
      Tab(1).Control(1)=   "ListView6"
      Tab(1).Control(2)=   "ListView7"
      Tab(1).Control(3)=   "ListView8"
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(6)=   "Label6"
      Tab(1).Control(7)=   "Label5"
      Tab(1).ControlCount=   8
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   3840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2415
         Left            =   5520
         TabIndex        =   7
         Top             =   1080
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   2415
         Left            =   5520
         TabIndex        =   9
         Top             =   3840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   11
         Top             =   1080
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   13
         Top             =   3840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView7 
         Height          =   2415
         Left            =   -69480
         TabIndex        =   15
         Top             =   1080
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView8 
         Height          =   2415
         Left            =   -69480
         TabIndex        =   17
         Top             =   3840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label8 
         Caption         =   "Usuarios sin Actividad en un mes"
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
         Left            =   -69360
         TabIndex        =   18
         Top             =   3600
         Width           =   5175
      End
      Begin VB.Label Label7 
         Caption         =   "Proveedor sin compras en un año"
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
         Left            =   -69360
         TabIndex        =   16
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label Label6 
         Caption         =   "Productos sin compras en un año"
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
         Left            =   -74760
         TabIndex        =   14
         Top             =   3600
         Width           =   5175
      End
      Begin VB.Label Label5 
         Caption         =   "Clientes sin compras en un año"
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
         Left            =   -74760
         TabIndex        =   12
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label Label4 
         Caption         =   "Deudas de Clientes sin Abono"
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
         Left            =   5640
         TabIndex        =   10
         Top             =   3600
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "Ordenes de Compra sin Pago"
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
         Left            =   5640
         TabIndex        =   8
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "Ventas Programadas Abiertas"
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
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "Ordenes de Compra Abiertas"
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
         Left            =   240
         TabIndex        =   3
         Top             =   3600
         Width           =   5175
      End
   End
End
Attribute VB_Name = "FrmMonitoreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
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
        .ColumnHeaders.Add , , "No. Pedido", 1000
        .ColumnHeaders.Add , , "Capturó", 1500
        .ColumnHeaders.Add , , "Cliente", 6500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Fecha Captura", 1500
        .ColumnHeaders.Add , , "No. de Orden", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Numero", 900
        .ColumnHeaders.Add , , "Tipo", 1200
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Proveedor", 2500
        .ColumnHeaders.Add , , "Total", 1000
        .ColumnHeaders.Add , , "Comentarios", 2500
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Numero", 1000
        .ColumnHeaders.Add , , "Tipo", 1000
        .ColumnHeaders.Add , , "Fecha", 1000
        .ColumnHeaders.Add , , "Proveedor", 2500
        .ColumnHeaders.Add , , "Total", 1000
        .ColumnHeaders.Add , , "Comentarios", 2500
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Venta", 1000
        .ColumnHeaders.Add , , "Factura", 1000
        .ColumnHeaders.Add , , "Cliente", 2500
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Fecha Vence", 1200
        .ColumnHeaders.Add , , "T. Deuda", 1000
        .ColumnHeaders.Add , , "T. Compra", 1000
    End With
    With ListView5
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id", 800
        .ColumnHeaders.Add , , "Cliente", 3500
        .ColumnHeaders.Add , , "RFC", 1400
    End With
    With ListView6
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id", 1500
        .ColumnHeaders.Add , , "Descripción", 3500
    End With
    With ListView7
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id", 800
        .ColumnHeaders.Add , , "Cliente", 3500
        .ColumnHeaders.Add , , "RFC", 1400
    End With
    With ListView8
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id", 800
        .ColumnHeaders.Add , , "Nombre", 3500
        .ColumnHeaders.Add , , "Puesto", 1400
    End With
    VentasProgramadas
    OrdenesCompra
    Pagos
    DeudaCliente
    Clientes
    Productos
    Proveedor
    Usuarios
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub VentasProgramadas()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim BusClie As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT PED_CLIEN.NO_PEDIDO, USUARIOS.NOMBRE AS USUARIO, dbo.CLIENTE.NOMBRE AS CLIENTE, PED_CLIEN.FECHA, PED_CLIEN.FECHA_CAPTURA, PED_CLIEN.NO_ORDEN FROM PED_CLIEN INNER JOIN CLIENTE ON PED_CLIEN.ID_CLIENTE = CLIENTE.ID_CLIENTE INNER JOIN USUARIOS ON PED_CLIEN.USUARIO = USUARIOS.ID_USUARIO WHERE PED_CLIEN.ESTADO IN ('I', 'C') AND PED_CLIEN.FECHA <= GETDATE() - 30 ORDER BY PED_CLIEN.NO_PEDIDO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("NO_PEDIDO"))
                If Not IsNull(.Fields("USUARIO")) Then tLi.SubItems(1) = .Fields("USUARIO")
                If Not IsNull(.Fields("CLIENTE")) Then tLi.SubItems(2) = .Fields("CLIENTE")
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(3) = .Fields("FECHA")
                If Not IsNull(.Fields("FECHA_CAPTURA")) Then tLi.SubItems(4) = .Fields("FECHA_CAPTURA")
                If Not IsNull(.Fields("NO_ORDEN")) Then tLi.SubItems(5) = .Fields("NO_ORDEN")
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub OrdenesCompra()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim sTipo As String
    Dim tLi As ListItem
    Dim tRs As ADODB.Recordset
    ListView2.ListItems.Clear
    sBuscar = "SELECT OC.Id_Orden_Compra, OC.NUM_ORDEN, OC.Id_Proveedor, P.Nombre, ((OC.Total - OC.Discount) + OC.TAX + OC.Freight + OC.Otros_Cargos) AS Total_Pagar, OC.COMENTARIO, OC.TIPO, OC.FECHA FROM ORDEN_COMPRA AS OC JOIN PROVEEDOR AS P ON P.Id_Proveedor = OC.Id_Proveedor WHERE OC.Confirmada IN ('P', 'S') AND OC.FECHA <= GETDATE() - 30 ORDER BY OC.ID_ORDEN_COMPRA"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        While Not .EOF
            If .Fields("TIPO") = "I" Then
                sTipo = "INTERNACIONAL"
            Else
                If .Fields("TIPO") = "N" Then
                    sTipo = "NACIONAL"
                Else
                    If .Fields("TIPO") = "X" Then
                        sTipo = "INDIRECTA"
                    Else
                        sTipo = "INDEFINIDA"
                    End If
                End If
            End If
            Set tLi = Me.ListView2.ListItems.Add(, , .Fields("NUM_ORDEN"))
            If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(1) = Trim(.Fields("FECHA"))
            tLi.SubItems(2) = sTipo
            If Not IsNull(.Fields("Nombre")) Then tLi.SubItems(3) = Trim(.Fields("Nombre"))
            If Not IsNull(.Fields("Total_Pagar")) Then tLi.SubItems(4) = Trim(.Fields("Total_Pagar"))
            If Not IsNull(.Fields("COMENTARIO")) Then tLi.SubItems(5) = Trim(.Fields("COMENTARIO"))
            .MoveNext
        Wend
    .Close
    End With
    sBuscar = "SELECT ORDEN_RAPIDA.ID_ORDEN_RAPIDA, ORDEN_RAPIDA.FECHA, PROVEEDOR_CONSUMO.NOMBRE, (SELECT SUM(TOTAL) AS TOTAL From ORDEN_RAPIDA_DETALLE WHERE (ID_ORDEN_RAPIDA = ORDEN_RAPIDA.ID_ORDEN_RAPIDA)) AS TOTAL, ORDEN_RAPIDA.COMENTARIO FROM ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR WHERE (ORDEN_RAPIDA.ESTADO = 'M') AND ORDEN_RAPIDA.FECHA <= GETDATE() - 30 ORDER BY ID_ORDEN_RAPIDA"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        While Not .EOF
            sTipo = "RAPIDA"
            Set tLi = Me.ListView2.ListItems.Add(, , .Fields("ID_ORDEN_RAPIDA"))
            If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(1) = Trim(.Fields("FECHA"))
            tLi.SubItems(2) = sTipo
            If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(3) = Trim(.Fields("NOMBRE"))
            If Not IsNull(.Fields("TOTAL")) Then tLi.SubItems(4) = Trim(.Fields("TOTAL"))
            If Not IsNull(.Fields("COMENTARIO")) Then tLi.SubItems(5) = Trim(.Fields("COMENTARIO"))
            .MoveNext
        Wend
    .Close
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Pagos()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim sTipo As String
    Dim tLi As ListItem
    Dim tRs As ADODB.Recordset
    ListView3.ListItems.Clear
    sBuscar = "SELECT OC.Id_Orden_Compra, OC.NUM_ORDEN, OC.Id_Proveedor, P.Nombre, ((OC.Total - OC.Discount) + OC.TAX + OC.Freight + OC.Otros_Cargos) AS Total_Pagar, OC.COMENTARIO, OC.TIPO, OC.FECHA FROM ORDEN_COMPRA AS OC JOIN PROVEEDOR AS P ON P.Id_Proveedor = OC.Id_Proveedor WHERE OC.Confirmada IN ('X') AND OC.FECHA <= GETDATE() - 30 ORDER BY OC.ID_ORDEN_COMPRA"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        While Not .EOF
            If .Fields("TIPO") = "I" Then
                sTipo = "INTERNACIONAL"
            Else
                If .Fields("TIPO") = "N" Then
                    sTipo = "NACIONAL"
                Else
                    If .Fields("TIPO") = "X" Then
                        sTipo = "INDIRECTA"
                    Else
                        sTipo = "INDEFINIDA"
                    End If
                End If
            End If
            Set tLi = Me.ListView3.ListItems.Add(, , .Fields("NUM_ORDEN"))
            If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(1) = Trim(.Fields("FECHA"))
            tLi.SubItems(2) = sTipo
            If Not IsNull(.Fields("Nombre")) Then tLi.SubItems(3) = Trim(.Fields("Nombre"))
            If Not IsNull(.Fields("Total_Pagar")) Then tLi.SubItems(4) = Trim(.Fields("Total_Pagar"))
            If Not IsNull(.Fields("COMENTARIO")) Then tLi.SubItems(5) = Trim(.Fields("COMENTARIO"))
            .MoveNext
        Wend
    .Close
    End With
    sBuscar = "SELECT ORDEN_RAPIDA.ID_ORDEN_RAPIDA, ORDEN_RAPIDA.FECHA, PROVEEDOR_CONSUMO.NOMBRE, (SELECT SUM(TOTAL) AS TOTAL From ORDEN_RAPIDA_DETALLE WHERE (ID_ORDEN_RAPIDA = ORDEN_RAPIDA.ID_ORDEN_RAPIDA)) AS TOTAL, ORDEN_RAPIDA.COMENTARIO FROM ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR WHERE (ORDEN_RAPIDA.ESTADO = 'A') AND ORDEN_RAPIDA.FECHA <= GETDATE() - 30 ORDER BY ID_ORDEN_RAPIDA"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        While Not .EOF
            sTipo = "RAPIDA"
            Set tLi = Me.ListView3.ListItems.Add(, , .Fields("ID_ORDEN_RAPIDA"))
            If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(1) = Trim(.Fields("FECHA"))
            tLi.SubItems(2) = sTipo
            If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(3) = Trim(.Fields("NOMBRE"))
            If Not IsNull(.Fields("TOTAL")) Then tLi.SubItems(4) = Trim(.Fields("TOTAL"))
            If Not IsNull(.Fields("COMENTARIO")) Then tLi.SubItems(5) = Trim(.Fields("COMENTARIO"))
            .MoveNext
        Wend
    .Close
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub DeudaCliente()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim BusClie As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView4.ListItems.Clear
    sBuscar = "SELECT VENTAS.ID_VENTA, VENTAS.FOLIO, CLIENTE.NOMBRE, CUENTAS.FECHA, CUENTAS.FECHA_VENCE, CUENTAS.deuda , Ventas.Total FROM VENTAS INNER JOIN CUENTA_VENTA ON VENTAS.ID_VENTA = CUENTA_VENTA.ID_VENTA INNER JOIN CUENTAS ON CUENTA_VENTA.ID_CUENTA = CUENTAS.ID_CUENTA INNER JOIN CLIENTE ON VENTAS.ID_CLIENTE = CLIENTE.ID_CLIENTE WHERE (CUENTAS.PAGADA = 'N') AND (VENTAS.FACTURADO IN (0, 1)) AND (CUENTAS.FECHA_VENCE <= GETDATE() - 30) ORDER BY VENTAS.ID_VENTA"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView4.ListItems.Add(, , .Fields("ID_VENTA"))
                If Not IsNull(.Fields("FOLIO")) Then tLi.SubItems(1) = .Fields("FOLIO")
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(2) = .Fields("NOMBRE")
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(3) = .Fields("FECHA")
                If Not IsNull(.Fields("FECHA_VENCE")) Then tLi.SubItems(4) = .Fields("FECHA_VENCE")
                If Not IsNull(.Fields("deuda")) Then tLi.SubItems(5) = Format(.Fields("deuda"), "###,###,##0.00")
                If Not IsNull(.Fields("Total")) Then tLi.SubItems(6) = Format(.Fields("Total"), "###,###,##0.00")
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Clientes()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim BusClie As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView5.ListItems.Clear
    sBuscar = "SELECT ID_CLIENTE, NOMBRE, RFC From Cliente WHERE (ID_CLIENTE NOT IN (SELECT Ventas.ID_CLIENTE From Ventas WHERE (FECHA <= GETDATE() - 365))) AND VALORACION = 'A'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView5.ListItems.Add(, , .Fields("ID_CLIENTE"))
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE")
                If Not IsNull(.Fields("RFC")) Then tLi.SubItems(2) = .Fields("RFC")
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Productos()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim BusClie As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView6.ListItems.Clear
    sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION From ALMACEN3 WHERE (ID_PRODUCTO NOT IN (SELECT VENTAS_DETALLE.ID_PRODUCTO FROM VENTAS_DETALLE)) ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView6.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("DESCRIPCION")) Then tLi.SubItems(1) = .Fields("DESCRIPCION")
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Proveedor()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim BusClie As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView7.ListItems.Clear
    sBuscar = "SELECT ID_PROVEEDOR, NOMBRE, RFC From Proveedor WHERE (ID_PROVEEDOR NOT IN (SELECT ID_PROVEEDOR From ORDEN_COMPRA WHERE (FECHA <= GETDATE() - 365))) AND (ELIMINADO = 'N') ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView7.ListItems.Add(, , .Fields("ID_PROVEEDOR"))
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE")
                If Not IsNull(.Fields("RFC")) Then tLi.SubItems(2) = .Fields("RFC")
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Usuarios()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim BusClie As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView8.ListItems.Clear
    sBuscar = "SELECT ID_USUARIO, NOMBRE, APELLIDOS, PUESTO From Usuarios WHERE (ID_USUARIO NOT IN (SELECT ID_USUARIO From Ventas WHERE (FECHA <= GETDATE() - 30))) AND (ESTADO = 'A') AND (ID_USUARIO NOT IN (SELECT ID_USUARIO From ORDEN_COMPRA WHERE (FECHA <= GETDATE() - 30))) AND (ID_USUARIO NOT IN (SELECT ID_USUARIO From ENTRADAS WHERE (FECHA <= GETDATE() - 30))) AND (ID_USUARIO NOT IN (SELECT ID_USUARIO From ABONOS_CUENTA WHERE (FECHA <= GETDATE() - 30))) AND (ID_USUARIO NOT IN(SELECT ID_USUARIO From ORDEN_RAPIDA WHERE (FECHA <= GETDATE() - 30))) AND (ID_USUARIO NOT IN (SELECT USUARIO From PED_CLIEN WHERE (FECHA <= GETDATE() - 30)))"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView8.ListItems.Add(, , .Fields("ID_USUARIO"))
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & " " & .Fields("APELLIDOS")
                If Not IsNull(.Fields("PUESTO")) Then tLi.SubItems(2) = .Fields("PUESTO")
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
