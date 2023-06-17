VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMuestraProgramadas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ventas Programadas en espera de ser cerradas"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView Lvw1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
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
   Begin MSComctlLib.ListView Lvw2 
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   8415
      _ExtentX        =   14843
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
End
Attribute VB_Name = "FrmMuestraProgramadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Form_Load()
   Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
        "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With

    With Lvw1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Pedido", 1000
        .ColumnHeaders.Add , , "Id Capturo", 0
        .ColumnHeaders.Add , , "Cliente", 6500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "No de Orden", 1500
        .ColumnHeaders.Add , , "Capturo", 2000
        .ColumnHeaders.Add , , "Id Cliente", 0
    End With
    With Lvw2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 4500
        .ColumnHeaders.Add , , "Cantidad Pedida", 2000
        .ColumnHeaders.Add , , "Cantidad en Existencia", 0
        .ColumnHeaders.Add , , "Cantidad Pendiente", 0
        .ColumnHeaders.Add , , "Precio Unitario", 2000
    End With
    Actualizar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Actualizar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim BusClie As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT NO_PEDIDO, USUARIO, NOMBRE, FECHA, NO_ORDEN, P.ID_CLIENTE FROM PED_CLIEN AS P JOIN CLIENTE AS C ON C.ID_CLIENTE = P.ID_CLIENTE WHERE P.ESTADO = 'C' ORDER BY NO_PEDIDO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            Lvw1.ListItems.Clear
            Do While Not .EOF
                Set tLi = Lvw1.ListItems.Add(, , .Fields("NO_PEDIDO"))
                If Not IsNull(.Fields("USUARIO")) Then tLi.SubItems(1) = .Fields("USUARIO")
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(2) = .Fields("NOMBRE")
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(3) = .Fields("FECHA")
                If Not IsNull(.Fields("NO_ORDEN")) Then tLi.SubItems(4) = .Fields("NO_ORDEN")
                If Not IsNull(.Fields("USUARIO")) Then
                    sBuscar = "SELECT NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & Val(.Fields("USUARIO"))
                    If Val(.Fields("USUARIO")) >= 0 Then
                        Set tRs2 = cnn.Execute(sBuscar)
                        If Not (tRs2.EOF And tRs2.BOF) Then
                            tLi.SubItems(5) = tRs2.Fields("NOMBRE") & " " & tRs2.Fields("APELLIDOS")
                        Else
                            tLi.SubItems(5) = .Fields("USUARIO")
                        End If
                    Else
                        tLi.SubItems(5) = .Fields("USUARIO")
                    End If
                End If
                If Not IsNull(.Fields("ID_CLIENTE")) Then tLi.SubItems(6) = .Fields("ID_CLIENTE")
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Lvw1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tLi As ListItem
    Dim DesClente As String
    Dim IdDescuento As String
    Dim IdDescuento2 As String
    sBuscar = "SELECT DESCUENTO, ID_DESCUENTO FROM CLIENTE WHERE ID_CLIENTE = " & Item.SubItems(6)
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        If Not IsNull(tRs.Fields("DESCUENTO")) Then
            DesClente = tRs.Fields("DESCUENTO")
        Else
            DesClente = 0
        End If
        If Not IsNull(tRs.Fields("ID_DESCUENTO")) Then
            IdDescuento = tRs.Fields("ID_DESCUENTO")
        Else
            IdDescuento = 0
        End If
    End If
    sBuscar = "SELECT * FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO = " & CDbl(Item)
    Set tRs = cnn.Execute(sBuscar)
    Lvw2.ListItems.Clear
    If (tRs.BOF And tRs.EOF) Then
        Lvw2.ListItems.Clear
        MsgBox "PEDIDO VACIO!", vbInformation, "SACC"
    Else
        tRs.MoveFirst
        Do While Not tRs.EOF
            Set tLi = Lvw2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO") & "")
            tLi.SubItems(1) = tRs.Fields("CANTIDAD_PEDIDA") & ""
            tLi.SubItems(2) = tRs.Fields("CANTIDAD_EXISTENCIA") & ""
            tLi.SubItems(3) = tRs.Fields("CANTIDAD_PENDIENTE") & ""
            '*********************************** AGREGAR PRECIO DE VENTA *************************************
            ' POR :   H VALDEZ
            ' FECHA:  20 DE MAYO DE 2011
            '*************************************************************************************************
            sBuscar = "SELECT PRECIO_VENTA FROM LICITACIONES WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND ID_CLIENTE = " & Item.SubItems(6)
            Set tRs2 = cnn.Execute(sBuscar)
            If Not (tRs2.BOF And tRs2.EOF) Then
                tLi.SubItems(4) = tRs2.Fields("PRECIO_VENTA")
            Else
                sBuscar = "SELECT PRECIO_COSTO, GANANCIA, CLASIFICACION FROM ALMACEN3 WHERE ID_PRODUCTO = '" & Trim(tRs.Fields("ID_PRODUCTO")) & "'"
                Set tRs1 = cnn.Execute(sBuscar)
                If Not (tRs1.EOF And tRs1.BOF) Then
                    tLi.SubItems(4) = Format((CDbl(tRs1.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs1.Fields("GANANCIA")))), "###,###,##0.00")
                Else
                    tLi.SubItems(4) = "0.00"
                End If
                If (DesClente <> "" And CDbl(DesClente) > 0) Then
                    tLi.SubItems(4) = CDbl(Replace(tLi.SubItems(3), ",", "")) * (CDbl(Replace(DesClente, ",", "")) / 100)
                    If (IdDescuento <> "") Then
                        sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & IdDescuento & "' AND CLASIFICACION = '" & tRs1.Fields("CLASIFICACION") & "'"
                        Set tRs2 = cnn.Execute(sBuscar)
                        IdDescuento2 = tRs2.Fields("PORCENTAJE")
                        tLi.SubItems(4) = (CDbl(tRs1.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs1.Fields("GANANCIA")))) - (CDbl(tRs1.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs1.Fields("GANANCIA")))) * CDbl(IdDescuento2) / 100
                    End If
                End If
            End If
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
