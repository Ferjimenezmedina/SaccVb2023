Attribute VB_Name = "mdlUtilerias"
Public BanLista As Byte
Public bInvCre As Boolean
Public INV As Integer
Public PedDir As Integer
Public PedInd As Integer
Public Sub Limpiar_Campos(FORMA As Form)
    Dim LimpCamp As Integer
    For LimpCamp = 0 To FORMA.Controls.Count - 1
        If TypeOf FORMA.Controls(LimpCamp) Is TextBox Then
            FORMA.Controls(LimpCamp).Text = ""
        End If
    Next
End Sub
Public Sub Abrir_Campos(FORMA As Form)
    Dim AbriCamp As Integer
    For AbriCamp = 0 To FORMA.Controls.Count - 1
        If TypeOf FORMA.Controls(AbriCamp) Is TextBox Then
            FORMA.Controls(AbriCamp).Locked = True
        End If
    Next
End Sub
Public Function Mayusculas(C As Integer) As Integer
    Mayusculas = Asc(UCase(Chr(C)))
End Function
Sub Main()
    frmProd.Show
End Sub
Public Function MESES(MES As String)
    Select Case MES
        Case 1: MESES = "ENERO"
        Case 2: MESES = "FEBRERO"
        Case 3: MESES = "MARZO"
        Case 4: MESES = "ABRIL"
        Case 5: MESES = "MAYO"
        Case 6: MESES = "JUNIO"
        Case 7: MESES = "JULIO"
        Case 8: MESES = "AGOSTO"
        Case 9: MESES = "SEPTIEMBRE"
        Case 10: MESES = "OCTUBRE"
        Case 11: MESES = "NOVIEMBRE"
        Case 12: MESES = "DICIEMBRE"
    End Select
End Function
Public Function Hay_Pedidos_Directos() As Boolean
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    sBuscar = "Select Count(*) AS CONTA From Pedido AS P Join Detalle_Pedido AS DP on DP.Id_Pedido=P.Id_Pedido Where P.Tipo='D' AND DP.Entregado='0'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        Hay_Pedidos_Directos = True
        PedDir = tRs.Fields("CONTA")
    Else
        Hay_Pedidos_Directos = False
    End If
End Function
Public Function Hay_Pedidos_Indirectos() As Boolean
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    sBuscar = "Select Count(*) AS CONTA From Pedido AS P Join Detalle_Pedido AS DP ON DP.Id_Pedido=P.Id_Pedido Where P.Tipo='I' AND DP.Entregado='0'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        Hay_Pedidos_Indirectos = True
        PedInd = tRs.Fields("CONTA")
    Else
        Hay_Pedidos_Indirectos = False
    End If
End Function
Public Function Hay_Usuarios() As Boolean
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    sBuscar = "SELECT ID_USUARIO FROM USUARIOS"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        Hay_Usuarios = True
    Else
        Hay_Usuarios = False
    End If
End Function
Public Sub KillProcess(ByVal processName As String)
On Error GoTo ErrHandler
    Dim oWMI
    Dim ret
    Dim sService
    Dim oWMIServices
    Dim oWMIService
    Dim oServices
    Dim oService
    Dim servicename
    Set oWMI = GetObject("winmgmts:")
    Set oServices = oWMI.InstancesOf("win32_process")
    For Each oService In oServices
        servicename = LCase(Trim(CStr(oService.Name) & ""))
        If InStr(1, servicename, LCase(processName), vbTextCompare) > 0 Then
            ret = oService.Terminate
        End If
    Next
    Set oServices = Nothing
    Set oWMI = Nothing
ErrHandler:
Err.Clear
End Sub

