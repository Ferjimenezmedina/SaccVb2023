VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOrden_Compra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ORDENES DE COMPRA"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   9840
      ScaleHeight     =   5595
      ScaleWidth      =   1395
      TabIndex        =   25
      Top             =   0
      Width           =   1455
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   29
         Top             =   4200
         Width           =   975
         Begin VB.Image Image9 
            Height          =   870
            Left            =   120
            MouseIcon       =   "frmOrden_Compra.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "frmOrden_Compra.frx":030A
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
            TabIndex        =   30
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin VB.TextBox txtId_Proveedor 
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   4695
      Begin VB.TextBox txtComentarios 
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox txtEnviara 
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label8 
         Caption         =   "COMENTARIOS"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "ENVIAR  A"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   4815
      Begin VB.CommandButton cmdImprimir 
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
         Height          =   375
         Left            =   3240
         Picture         =   "frmOrden_Compra.frx":23EC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton opnIndirecta 
         Caption         =   "Indirecta"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton opnInternacional 
         Caption         =   "Internacional"
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton opnNacional 
         Caption         =   "Nacional"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox txtDescuento 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Text            =   "0"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtFlete 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Text            =   "0"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtCargos 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Text            =   "0"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtImpuesto 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Text            =   "0"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtSubtotal 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTotal 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "0"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblFolio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "FOLIO"
         Height          =   255
         Left            =   3120
         TabIndex        =   27
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "DESCUENTO"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "OTROS CARGOS"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "FLETE"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "IMPUESTO"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SUBTOTAL"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lvwProveedores 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2355
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwCotizaciones 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   9615
      _ExtentX        =   16960
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
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblProveedor 
      Alignment       =   2  'Center
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
      Left            =   5040
      TabIndex        =   26
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmOrden_Compra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As Recordset
Dim cont As Integer
Dim NoRe As Integer
Private Sub cmdImprimir_Click()
    On Error GoTo ManejaError
    If Puede_Guardar Then
    Dim ID_PROVEEDOR As String
    Dim Total As String
    Dim ENVIARA As String
    Dim DISCOUNT As String
    Dim FREIGHT As String
    Dim TAX As String
    Dim OTROS_CARGOS As String
    Dim COMENTARIO As String
    Dim Tipo As String
    Dim NUM_ORDEN As Integer
    Dim ID_ORDEN_COMPRA As Integer
    
    ID_PROVEEDOR = Me.txtId_Proveedor.Text
    Total = Me.txtSubtotal.Text
    ENVIARA = Me.txtEnviara.Text
    DISCOUNT = Me.txtDescuento.Text
    FREIGHT = Me.txtFlete.Text
    TAX = Me.txtImpuesto.Text
    OTROS_CARGOS = Me.txtCargos.Text
    COMENTARIO = Me.txtComentarios.Text
    If Me.opnInternacional.Value = True Then
        Tipo = "I"
    ElseIf Me.opnNacional.Value = True Then
        Tipo = "N"
    Else
        Tipo = "X"
    End If
        
        'INICIO TRAER ULTIMA ORDEN DE COMPRA SEGUN TIPO
        NUM_ORDEN = 0
        sqlQuery = "SELECT TOP 1 NUM_ORDEN FROM ORDEN_COMPRA WHERE TIPO = '" & Tipo & "' ORDER BY NUM_ORDEN DESC"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            If Not (.BOF And .EOF) Then
                    If Not IsNull(.Fields("NUM_ORDEN")) Then NUM_ORDEN = .Fields("NUM_ORDEN")
                .Close
            End If
            NUM_ORDEN = NUM_ORDEN + 1
        End With
        'FIN TRAER ULTIMA ORDEN DE COMPRA
            
        'INICIO TRER ULTIMO ID_ORDEN_COMPRA
        ID_ORDEN_COMPRA = 0
        sqlQuery = "SELECT TOP 1 ID_ORDEN_COMPRA FROM ORDEN_COMPRA ORDER BY ID_ORDEN_COMPRA DESC"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            If Not (.BOF And .EOF) Then
                    If Not IsNull(.Fields("ID_ORDEN_COMPRA")) Then ID_ORDEN_COMPRA = .Fields("ID_ORDEN_COMPRA")
                .Close
            End If
            ID_ORDEN_COMPRA = ID_ORDEN_COMPRA + 1
        End With
        'FIN
        Total = Replace(Total, ",", ".")
        DISCOUNT = Replace(DISCOUNT, ",", ".")
        FREIGHT = Replace(FREIGHT, ",", ".")
        TAX = Replace(TAX, ",", ".")
        OTROS_CARGOS = Replace(OTROS_CARGOS, ",", ".")
        sqlQuery = "INSERT INTO ORDEN_COMPRA (ID_PROVEEDOR, FECHA, TOTAL, ENVIARA, DISCOUNT, FREIGHT, TAX, OTROS_CARGOS, COMENTARIO, CONFIRMADA, TIPO, NUM_ORDEN) VALUES (" & ID_PROVEEDOR & ", '" & Format(Date, "dd/mm/yyyy") & "', " & Total & ", '" & ENVIARA & "', " & DISCOUNT & ", " & FREIGHT & ", " & TAX & ", " & OTROS_CARGOS & ", '" & COMENTARIO & "', 'S', '" & Tipo & "', " & NUM_ORDEN & ")"
        Set tRs = cnn.Execute(sqlQuery)
        
        Dim ID_PRODUCTO As String
        Dim DESCRIPCION As String
        Dim CANTIDAD As Double
        Dim Precio As Double
        Dim DIAS_ENTREGA As Integer
        
            NoRe = Me.lvwCotizaciones.ListItems.Count
            For cont = 1 To NoRe
                ID_PRODUCTO = Me.lvwCotizaciones.ListItems.Item(cont).SubItems(3)
                DESCRIPCION = Me.lvwCotizaciones.ListItems.Item(cont).SubItems(4)
                CANTIDAD = Me.lvwCotizaciones.ListItems.Item(cont).SubItems(5)
                DIAS_ENTREGA = Me.lvwCotizaciones.ListItems.Item(cont).SubItems(6)
                Precio = Me.lvwCotizaciones.ListItems.Item(cont).SubItems(7)
                
                sqlQuery = "INSERT INTO ORDEN_COMPRA_DETALLE (ID_ORDEN_COMPRA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO, DIAS_ENTREGA) VALUES (" & ID_ORDEN_COMPRA & ", '" & ID_PRODUCTO & "', '" & DESCRIPCION & "', " & CANTIDAD & ", " & Precio & ", " & DIAS_ENTREGA & ")"
                Set tRs = cnn.Execute(sqlQuery)
                
            Next cont
        
        'CAMNIAR ESTADO DE COTIZACION
        Dim ID_COTIZACION As Integer
        NoRe = Me.lvwCotizaciones.ListItems.Count
        For cont = 1 To NoRe
            ID_COTIZACION = Me.lvwCotizaciones.ListItems.Item(cont)
            sqlQuery = "UPDATE COTIZA_REQUI SET ESTADO_ACTUAL = 'C' WHERE ID_COTIZACION = " & ID_COTIZACION
            Set tRs = cnn.Execute(sqlQuery)
        Next cont
        
        Me.lvwProveedores.ListItems.Clear
        Me.lvwCotizaciones.ListItems.Clear
        Llenar_Lista_Proveedores
    End If
    
    Exit Sub
    'EL SIGUIENTE CODIGO ES DE IMPRECION DE INFORMES. NO DEBE EJECUTARCE
        If Me.opnInternacional.Value = True Then
                        deAPTONER.TRAER_ULTIMA_ORDEN_COMPRA "I"
                        With deAPTONER.rsTRAER_ULTIMA_ORDEN_COMPRA
                            If Not IsNull(!NUM_ORDEN) Then NOC = !NUM_ORDEN + 1
                            .Close
                        End With
                        deAPTONER.ACTUALIZA_NUM_ORDEN_COMPRA nOrdenCompra, NOC
                        Set crReport = crApplication.OpenReport("C:\APTONER\ORDEN_COMPRA.rpt")
                        crReport.ParameterFields.Item(1).ClearCurrentValueAndRange
                        crReport.ParameterFields.Item(2).ClearCurrentValueAndRange
                        crReport.ParameterFields.Item(1).AddCurrentValue nOrdenCompra
                        crReport.ParameterFields.Item(2).AddCurrentValue VarMen.Text1(0).Text
                        crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        crReport.PrintOut
        ElseIf Me.opnNacional.Value = True Then
                        deAPTONER.TRAER_ULTIMA_ORDEN_COMPRA "N"
                        With deAPTONER.rsTRAER_ULTIMA_ORDEN_COMPRA
                            If Not IsNull(!NUM_ORDEN) Then NOC = !NUM_ORDEN + 1
                            .Close
                        End With
                        deAPTONER.ACTUALIZA_NUM_ORDEN_COMPRA nOrdenCompra, NOC
                        Set crReport = crApplication.OpenReport("C:\APTONER\ORDEN_COMPRA_NACIONAL.rpt")
                        crReport.ParameterFields.Item(1).ClearCurrentValueAndRange
                        crReport.ParameterFields.Item(2).ClearCurrentValueAndRange
                        crReport.ParameterFields.Item(1).AddCurrentValue nOrdenCompra
                        crReport.ParameterFields.Item(2).AddCurrentValue VarMen.Text1(0).Text
                        crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        crReport.PrintOut
        Else
                        deAPTONER.TRAER_ULTIMA_ORDEN_COMPRA "X"
                        With deAPTONER.rsTRAER_ULTIMA_ORDEN_COMPRA
                            If Not IsNull(!NUM_ORDEN) Then NOC = !NUM_ORDEN + 1
                            .Close
                        End With
                        deAPTONER.ACTUALIZA_NUM_ORDEN_COMPRA nOrdenCompra, NOC
                        Set crReport = crApplication.OpenReport("C:\APTONER\ORDEN_COMPRA_NACIONAL.rpt")
                        crReport.ParameterFields.Item(1).ClearCurrentValueAndRange
                        crReport.ParameterFields.Item(2).ClearCurrentValueAndRange
                        crReport.ParameterFields.Item(1).AddCurrentValue nOrdenCompra
                        crReport.ParameterFields.Item(2).AddCurrentValue VarMen.Text1(0).Text
                        crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        crReport.PrintOut
        End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
    Err.Clear
End Sub
Private Sub Form_Activate()
On Error GoTo ManejaError
    If Hay_Proveedores Then
        Llenar_Lista_Proveedores
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
    Err.Clear
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
        With Me.lvwCotizaciones
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID COTIZACION", 0
        .ColumnHeaders.Add , , "ID REQUISICION", 0
        .ColumnHeaders.Add , , "ID PROVEEDOR", 0
        .ColumnHeaders.Add , , "CLAVE DEL PRODUCTO", 2200
        .ColumnHeaders.Add , , "DESCRIPCION", 3500
        .ColumnHeaders.Add , , "CANTIDAD", 1000, 2
        .ColumnHeaders.Add , , "DIAS ENTREGA", 1440, 2
        .ColumnHeaders.Add , , "PRECIO", 1440, 2
        .ColumnHeaders.Add , , "FECHA", 0, 2
    End With
    With Me.lvwProveedores
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID PROVEEDOR", 0
        .ColumnHeaders.Add , , "PROVEEDOR", 4500, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
    Err.Clear
End Sub
Sub Llenar_Lista_Proveedores()
On Error GoTo ManejaError
    Dim nId_Proveedor As Integer
    sqlQuery = "SELECT P.ID_PROVEEDOR, P.NOMBRE, P.DIRECCION, P.COLONIA, P.CIUDAD, P.CP, P.RFC, P.TELEFONO1, P.TELEFONO2, P.TELEFONO3, P.NOTAS, P.ESTADO, P.PAIS FROM PROVEEDOR AS P JOIN COTIZA_REQUI AS C ON C.ID_PROVEEDOR = P.ID_PROVEEDOR WHERE P.ELIMINADO = 'N' AND C.ESTADO_ACTUAL = 'X' ORDER BY P.NOMBRE"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.BOF And .EOF) Then
            Me.lvwProveedores.ListItems.Clear
            Do While Not .EOF
                If nId_Proveedor <> .Fields("ID_PROVEEDOR") Then
                    Set tLi = lvwProveedores.ListItems.Add(, , .Fields("ID_PROVEEDOR"))
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = Trim(.Fields("NOMBRE"))
                    If Not IsNull(.Fields("DIRECCION")) Then tLi.SubItems(2) = Trim(.Fields("DIRECCION"))
                    If Not IsNull(.Fields("COLONIA")) Then tLi.SubItems(3) = Trim(.Fields("COLONIA"))
                    If Not IsNull(.Fields("CP")) Then tLi.SubItems(4) = Trim(.Fields("CP"))
                    If Not IsNull(.Fields("RFC")) Then tLi.SubItems(5) = Trim(.Fields("RFC"))
                    If Not IsNull(.Fields("TELEFONO1")) Then tLi.SubItems(6) = Trim(.Fields("TELEFONO1"))
                    If Not IsNull(.Fields("TELEFONO2")) Then tLi.SubItems(7) = Trim(.Fields("TELEFONO2"))
                    If Not IsNull(.Fields("TELEFONO3")) Then tLi.SubItems(8) = Trim(.Fields("TELEFONO3"))
                    If Not IsNull(.Fields("NOTAS")) Then tLi.SubItems(9) = Trim(.Fields("NOTAS"))
                    If Not IsNull(.Fields("ESTADO")) Then tLi.SubItems(10) = Trim(.Fields("ESTADO"))
                    If Not IsNull(.Fields("PAIS")) Then tLi.SubItems(11) = Trim(.Fields("PAIS"))
                    'INICIO PARA NO REPETIR PROVEEDORES
                    nId_Proveedor = .Fields("ID_PROVEEDOR")
                    ' FIN
                End If
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
    Err.Clear
End Sub
Sub Llenar_Lista_Cotizaciones(nId_Proveedor As Integer)
On Error GoTo ManejaError
    sqlQuery = "SELECT ID_COTIZACION, ID_REQUISICION, ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, CANTIDAD, DIAS_ENTREGA, PRECIO, FECHA FROM COTIZA_REQUI WHERE ESTADO_ACTUAL = 'X' AND ID_PROVEEDOR = " & nId_Proveedor
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        Me.lvwCotizaciones.ListItems.Clear
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvwCotizaciones.ListItems.Add(, , .Fields("ID_COTIZACION"))
                If Not IsNull(.Fields("ID_REQUISICION")) Then tLi.SubItems(1) = Trim(.Fields("ID_REQUISICION"))
                If Not IsNull(.Fields("ID_PROVEEDOR")) Then tLi.SubItems(2) = Trim(.Fields("ID_PROVEEDOR"))
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(3) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("DESCRIPCION")) Then tLi.SubItems(4) = Trim(.Fields("DESCRIPCION"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(5) = Trim(.Fields("CANTIDAD"))
                If Not IsNull(.Fields("DIAS_ENTREGA")) Then tLi.SubItems(6) = Trim(.Fields("DIAS_ENTREGA"))
                If Not IsNull(.Fields("PRECIO")) Then tLi.SubItems(7) = Trim(.Fields("PRECIO"))
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(8) = Trim(.Fields("FECHA"))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
    Err.Clear
End Sub

Function Hay_Cotizaciones() As Boolean
On Error GoTo ManejaError
    sqlQuery = "SELECT COUNT(ID_COTIZACION)ID_COTIZACION FROM COTIZA_REQUI WHERE ESTADO_ACTUAL = 'X'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_COTIZACION") <> 0 Then
            Hay_Cotizaciones = True
        Else
            Hay_Cotizaciones = False
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
    Err.Clear
End Function
Function Hay_Proveedores() As Boolean
On Error GoTo ManejaError
    sqlQuery = "SELECT COUNT(ID_PROVEEDOR)ID_PROVEEDOR FROM PROVEEDOR WHERE ELIMINADO = 'N'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_PROVEEDOR") <> 0 Then
            Hay_Proveedores = True
        Else
            Hay_Proveedores = False
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
    Err.Clear
End Function
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwProveedores_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    If Hay_Cotizaciones Then
        Llenar_Lista_Cotizaciones (Item)
        Me.txtId_Proveedor.Text = Item
        lblFolio.Caption = Item '?
        Me.lblProveedor.Caption = Item.SubItems(1)
        Sumar_Importe
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Sumar_Importe()
On Error GoTo ManejaError
    Me.txtSubtotal.Text = ""
    NoRe = Me.lvwCotizaciones.ListItems.Count
    For cont = 1 To NoRe
        Me.txtSubtotal.Text = Val(Me.txtSubtotal.Text) + (Me.lvwCotizaciones.ListItems.Item(cont).SubItems(5) * Me.lvwCotizaciones.ListItems.Item(cont).SubItems(7))
    Next cont
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub

Private Sub opnIndirecta_Click()
    txtImpuesto.Text = Val(txtSubtotal.Text) * 0.15
    Me.txtTotal.Text = CDbl(Val(Me.txtSubtotal.Text)) - CDbl(Val(Me.txtDescuento.Text)) + CDbl(Val(Me.txtImpuesto.Text)) + CDbl(Val(Me.txtFlete.Text)) + CDbl(Val(Me.txtCargos.Text))
End Sub

Private Sub opnInternacional_Click()
    txtImpuesto.Text = "0.00"
    Me.txtTotal.Text = CDbl(Val(Me.txtSubtotal.Text)) - CDbl(Val(Me.txtDescuento.Text)) + CDbl(Val(Me.txtImpuesto.Text)) + CDbl(Val(Me.txtFlete.Text)) + CDbl(Val(Me.txtCargos.Text))
End Sub

Private Sub opnNacional_Click()
    txtImpuesto.Text = Val(txtSubtotal.Text) * 0.15
    Me.txtTotal.Text = CDbl(Val(Me.txtSubtotal.Text)) - CDbl(Val(Me.txtDescuento.Text)) + CDbl(Val(Me.txtImpuesto.Text)) + CDbl(Val(Me.txtFlete.Text)) + CDbl(Val(Me.txtCargos.Text))
End Sub

Private Sub txtCargos_Change()
On Error GoTo ManejaError
    'Me.txtImpuesto.Text = (CDbl(Val(Me.txtSubtotal.Text)) - CDbl(Val(Me.txtDescuento.Text))) * 0.15
    Me.txtTotal.Text = CDbl(Val(Me.txtSubtotal.Text)) - CDbl(Val(Me.txtDescuento.Text)) + CDbl(Val(Me.txtImpuesto.Text)) + CDbl(Val(Me.txtFlete.Text)) + CDbl(Val(Me.txtCargos.Text))
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub txtCargos_GotFocus()
    txtCargos.BackColor = &HFFE1E1
End Sub
Private Sub txtCargos_LostFocus()
      txtCargos.BackColor = &H80000005
End Sub
Private Sub txtComentarios_GotFocus()
    txtComentarios.BackColor = &HFFE1E1
End Sub
Private Sub txtComentarios_LostFocus()
      txtComentarios.BackColor = &H80000005
End Sub
Private Sub txtDescuento_Change()
On Error GoTo ManejaError
    If Not opnInternacional.Value Then
        Me.txtImpuesto.Text = (CDbl(Val(Me.txtSubtotal.Text)) - CDbl(Val(Me.txtDescuento.Text))) * 0.15
    Else
        Me.txtImpuesto.Text = "0.00"
    End If
    Me.txtTotal.Text = CDbl(Val(Me.txtSubtotal.Text)) - CDbl(Val(Me.txtDescuento.Text)) + CDbl(Val(Me.txtImpuesto.Text)) + CDbl(Val(Me.txtFlete.Text)) + CDbl(Val(Me.txtCargos.Text))
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub txtDescuento_GotFocus()
    txtDescuento.BackColor = &HFFE1E1
End Sub
Private Sub txtDescuento_LostFocus()
      txtDescuento.BackColor = &H80000005
End Sub
Private Sub txtEnviara_GotFocus()
    txtEnviara.BackColor = &HFFE1E1
End Sub
Private Sub txtEnviara_LostFocus()
      txtEnviara.BackColor = &H80000005
End Sub
Private Sub txtFlete_Change()
On Error GoTo ManejaError
    'Me.txtImpuesto.Text = (CDbl(Val(Me.txtSubtotal.Text)) - CDbl(Val(Me.txtDescuento.Text))) * 0.15
    Me.txtTotal.Text = CDbl(Val(Me.txtSubtotal.Text)) - CDbl(Val(Me.txtDescuento.Text)) + CDbl(Val(Me.txtImpuesto.Text)) + CDbl(Val(Me.txtFlete.Text)) + CDbl(Val(Me.txtCargos.Text))
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub txtFlete_GotFocus()
    txtFlete.BackColor = &HFFE1E1
End Sub
Private Sub txtFlete_LostFocus()
      txtFlete.BackColor = &H80000005
End Sub
Private Sub txtImpuesto_Change()
On Error GoTo ManejaError
    Me.txtImpuesto.Text = (CDbl(Val(Me.txtSubtotal.Text)) - CDbl(Val(Me.txtDescuento.Text))) * 0.15
    Me.txtTotal.Text = CDbl(Val(Me.txtSubtotal.Text)) - CDbl(Val(Me.txtDescuento.Text)) + CDbl(Val(Me.txtImpuesto.Text)) + CDbl(Val(Me.txtFlete.Text)) + CDbl(Val(Me.txtCargos.Text))
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub txtImpuesto_GotFocus()
    txtImpuesto.BackColor = &HFFE1E1
End Sub
Private Sub txtImpuesto_LostFocus()
      txtImpuesto.BackColor = &H80000005
End Sub
Private Sub txtSubtotal_Change()
On Error GoTo ManejaError
    Me.txtImpuesto.Text = (CDbl(Val(Me.txtSubtotal.Text)) - CDbl(Val(Me.txtDescuento.Text))) * 0.15
    Me.txtTotal.Text = CDbl(Val(Me.txtSubtotal.Text)) - CDbl(Val(Me.txtDescuento.Text)) + CDbl(Val(Me.txtImpuesto.Text)) + CDbl(Val(Me.txtFlete.Text)) + CDbl(Val(Me.txtCargos.Text))
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub txtSubtotal_GotFocus()
    txtSubtotal.BackColor = &HFFE1E1
End Sub
Private Sub txtSubtotal_LostFocus()
      txtSubtotal.BackColor = &H80000005
End Sub
Function Puede_Guardar() As Boolean
On Error GoTo ManejaError
    If Me.txtId_Proveedor.Text = "" Then
        MsgBox "SELECCIONE EL PROVEEDOR", vbInformation, "SACC"
        Puede_Guardar = False
        Exit Function
    End If
    Puede_Guardar = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
    Err.Clear
End Function
Private Sub txtTotal_GotFocus()
    txtTotal.BackColor = &HFFE1E1
End Sub
Private Sub txtTotal_LostFocus()
      txtTotal.BackColor = &H80000005
End Sub
