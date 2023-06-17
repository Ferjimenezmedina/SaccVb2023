VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReviComa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEDIR MAXIMOS Y MINIMOS"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "P.Separados"
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
      Left            =   6840
      Picture         =   "frmReviComa.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdPedir 
      Caption         =   "P. Juntos"
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
      Left            =   8400
      Picture         =   "frmReviComa.frx":29D2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9720
      TabIndex        =   2
      Top             =   5640
      Width           =   975
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Image cmdCancelar 
         Height          =   705
         Left            =   120
         MouseIcon       =   "frmReviComa.frx":53A4
         MousePointer    =   99  'Custom
         Picture         =   "frmReviComa.frx":56AE
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   11055
      Begin VB.CommandButton Command1 
         Caption         =   "Almacen 3"
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
         Left            =   2160
         Picture         =   "frmReviComa.frx":7160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdTraer 
         Caption         =   "Almacen 2"
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
         Left            =   600
         Picture         =   "frmReviComa.frx":9B32
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1200
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lvwFaltas 
      Height          =   4575
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmReviComa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim sqlQuery As String
Dim tLi As ListItem
Dim lvSI As ListSubItem
Dim tRs As Recordset
Dim tRs2 As Recordset
Dim intIndex As Integer
Dim bBandExis As Boolean
Private Sub cmdPedir_Click()
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim NoRe As Integer
    Dim nPedido As Integer
    Dim cProducto As String
    Dim Uno As String
    Dim cMin As Double
    Dim cMax As Double
    Dim CantPed As Double
    Dim CantPedF As Double
    Dim exist As Integer
    Dim Almacen As String
    Dim sqlQuery As String
    cMin = 0
    cMax = 0
    CantPed = 0
        If MsgBox("¿DESEA HACER UN PEDIDO POR LOS ARTICULOS SELECCIONADOS?", vbQuestion + vbYesNo + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then
            Uno = "S"
            NoRe = Me.lvwFaltas.ListItems.Count
            For Cont = 1 To NoRe
                If Me.lvwFaltas.ListItems.Item(Cont).Checked = True Then
                    CantPedF = CDbl(lvwFaltas.ListItems.Item(Cont).SubItems(4))
                        If (CantPedF > 0) Then
                            If Me.lvwFaltas.ListItems.Item(Cont).SubItems(6) = "M" Then
                                If Uno = "S" Then
                                    sqlQuery = "INSERT INTO PEDIDO (SUCURSAL, PIDIO, FECHA, TIPO, COMENTARIO) VALUES ('BODEGA', 'SISTEMA', '" & Format(Date, "dd/mm/yyyy") & "', 'I', 'MAXIMOS Y MINIMOS')"
                                    cnn.Execute (sqlQuery)
                                    sqlQuery = "SELECT TOP 1 ID_PEDIDO FROM PEDIDO ORDER BY ID_PEDIDO DESC"
                                    Set tRs = cnn.Execute(sqlQuery)
                                    nPedido = tRs.Fields("ID_PEDIDO")
                                    Uno = "N"
                                End If
                                sqlQuery = "INSERT INTO DETALLE_PEDIDO (ID_PRODUCTO, CANTIDAD, ID_PEDIDO, DESCRIPCION, ALMACEN) VALUES ('" & Trim(lvwFaltas.ListItems.Item(Cont)) & "', " & CantPedF & ", " & nPedido & ", '" & lvwFaltas.ListItems.Item(Cont).SubItems(1) & "', '" & lvwFaltas.ListItems.Item(Cont).SubItems(5) & "')"
                                cnn.Execute (sqlQuery)
                            End If
                        End If
                End If
            Next Cont
            Uno = "S"
            NoRe = Me.lvwFaltas.ListItems.Count
            For Cont = 1 To NoRe
                If Me.lvwFaltas.ListItems.Item(Cont).Checked = True Then
                    CantPedF = CDbl(lvwFaltas.ListItems.Item(Cont).SubItems(4))
                        If (CantPedF > 0) Then
                            If Me.lvwFaltas.ListItems.Item(Cont).SubItems(6) = "L" Then
                                If Uno = "S" Then
                                    sqlQuery = "INSERT INTO PEDIDO (SUCURSAL, PIDIO, FECHA, TIPO, COMENTARIO) VALUES ('BODEGA', 'SISTEMA', '" & Format(Date, "dd/mm/yyyy") & "', 'I', 'LICITACION')"
                                    cnn.Execute (sqlQuery)
                                    sqlQuery = "SELECT TOP 1 ID_PEDIDO FROM PEDIDO ORDER BY ID_PEDIDO DESC"
                                    Set tRs = cnn.Execute(sqlQuery)
                                    nPedido = tRs.Fields("ID_PEDIDO")
                                    Uno = "N"
                                End If
                                sqlQuery = "INSERT INTO DETALLE_PEDIDO (ID_PRODUCTO, CANTIDAD, ID_PEDIDO, DESCRIPCION, ALMACEN) VALUES ('" & Trim(lvwFaltas.ListItems.Item(Cont).SubItems(1)) & "', " & CantPedF & ", " & nPedido & ", '" & lvwFaltas.ListItems.Item(Cont).SubItems(1) & "', '" & lvwFaltas.ListItems.Item(Cont).SubItems(5) & "')"
                                cnn.Execute (sqlQuery)
                            End If
                        End If
                End If
            Next Cont
            MsgBox "PEDIDO TERMINADO", vbInformation, "MENSAJE DEL SISTEMA"
            lvwFaltas.ListItems.Clear
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

Private Sub cmdCancelar_Click()

On Error GoTo ManejaError

    Unload Me
    
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

Private Sub cmdTraer_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tRs2 As Recordset
    Dim Ultimo As Integer
    Dim Cont As Integer
    Dim Fin As Integer
    Dim Resu As Long
    Dim CantPed As Long
    Dim CantMin As Long
    
    sBuscar = "SELECT * FROM VSMAXMINA2"
    Set tRs = cnn.Execute(sBuscar)
    
    lvwFaltas.ListItems.Clear
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                sqlQuery = "SELECT ISNULL(SUM(D.CANTIDAD), 0) AS CANTIDAD FROM DETALLE_PEDIDO AS D JOIN PEDIDO AS P ON D.ID_PEDIDO = P.ID_PEDIDO WHERE D.ID_PRODUCTO = '" & .Fields("ID_PRODUCTO") & "' AND (D.ENTREGADO = '0' OR D.ENTREGADO = 'R') AND P.COMENTARIO = 'MAXIMOS Y MINIMOS' AND P.PIDIO = 'SISTEMA'"
                Set tRs2 = cnn.Execute(sqlQuery)
                If Not (tRs2.EOF And tRs2.BOF) Then
                    CantPed = tRs2.Fields("CANTIDAD")
                End If
                If CDbl(.Fields("FALTANTE")) - CantPed > 0 Then
                    CantMin = CDbl(.Fields("FALTANTE")) - CantPed
                    CantMin = CantMin - CDbl(.Fields("FALTANTEMIN"))
                    
                    Set tLi = lvwFaltas.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                        tLi.SubItems(1) = .Fields("DESCRIPCION")
                        If CantMin > 0 Then
                            tLi.SubItems(2) = CantMin
                            tLi.SubItems(3) = CDbl(.Fields("FALTANTEMIN"))
                        Else
                            tLi.SubItems(2) = 0
                            tLi.SubItems(3) = CDbl(.Fields("FALTANTE")) - CantPed
                        End If
                        tLi.SubItems(4) = CDbl(.Fields("FALTANTE")) - CantPed
                        tLi.SubItems(5) = "A2"
                        tLi.SubItems(6) = "M"
                    Ultimo = lvwFaltas.ListItems.Count
                    lvwFaltas.ListItems.Item(Ultimo).Checked = True
                End If
                .MoveNext
            Loop
        End If
    End With
    Fin = lvwFaltas.ListItems.Count
    sBuscar = "SELECT V.NOMBRE, V.ID_PRODUCTO, V.DESCRIPCION, V.CANTIDAD, V.CANT_MIN, NO_CONTRATO, ISNULL(E.CANTIDAD, 0) AS EXISTENCIA FROM VSTOTALVENTALICITA AS V JOIN ALMACEN2 AS A ON V.ID_PRODUCTO = A.ID_PRODUCTO LEFT JOIN EXISTENCIAS AS E ON V.ID_PRODUCTO = E.ID_PRODUCTO WHERE E.SUCURSAL = 'BODEGA'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Cont = 1
                Do While Cont <= Fin
                    Cont = Cont + 1
                Loop
                If Cont > Fin Then
                    If CDbl(.Fields("CANT_MIN")) - CDbl(.Fields("CANTIDAD")) > 0 Then
                        Resu = ((CDbl(.Fields("CANT_MIN")) - CDbl(.Fields("CANTIDAD"))) * 0.2)
                        If Resu < 1 Then
                            Resu = 1
                        End If
                        sqlQuery = "SELECT ISNULL(SUM(D.CANTIDAD), 0) AS CANTIDAD FROM DETALLE_PEDIDO AS D JOIN PEDIDO AS P ON D.ID_PEDIDO = P.ID_PEDIDO WHERE D.ID_PRODUCTO = '" & .Fields("ID_PRODUCTO") & "' AND (D.ENTREGADO = '0' OR D.ENTREGADO = 'R') AND P.COMENTARIO = 'LICITACION' AND P.PIDIO = 'SISTEMA'"
                        Set tRs2 = cnn.Execute(sqlQuery)
                        If Not (tRs2.EOF And tRs2.BOF) Then
                            CantPed = tRs2.Fields("CANTIDAD")
                        End If
                        tRs2.Close
                        CantPed = CantPed + .Fields("EXISTENCIA")
                        If Resu - CantPed > 0 Then
                            Set tLi = lvwFaltas.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                                tLi.SubItems(1) = .Fields("ID_PRODUCTO")
                                tLi.SubItems(2) = Resu - CantPed
                                tLi.SubItems(3) = 0
                                tLi.SubItems(4) = Resu - CantPed
                                tLi.SubItems(5) = "A2"
                                tLi.SubItems(6) = "L"
                        End If
                    End If
                End If
                .MoveNext
                Ultimo = lvwFaltas.ListItems.Count
                lvwFaltas.ListItems.Item(Ultimo).Checked = True
                lvwFaltas.ListItems.Item(Ultimo).Bold = True
                lvwFaltas.ListItems.Item(Ultimo).ForeColor = vbRed
            Loop
        End If
        .Close
    End With
    
    If lvwFaltas.ListItems.Count = 0 Then
        MsgBox "NO HAY PRODUCTOS ABAJO DEL MINIMO", vbInformation, "MENSAJE DEL SISTEMA"
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
            "Provider=SQLOLEDB.1;Password=SaccPass*20;Persist Security Info=True;User ID=AdmSACC2210;Initial Catalog=APTONER;" & _
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

Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tRs2 As Recordset
    Dim Ultimo As Integer
    Dim Cont As Integer
    Dim Fin As Integer
    Dim Resu As Long
    Dim CantPed As Long
    Dim CantMin As Long
    
    
    sBuscar = "SELECT * FROM VSMAXMINA3"
    Set tRs = cnn.Execute(sBuscar)
    lvwFaltas.ListItems.Clear
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                CantPed = 0
                sqlQuery = "SELECT ISNULL(SUM(D.CANTIDAD), 0) AS CANTIDAD FROM DETALLE_PEDIDO AS D JOIN PEDIDO AS P ON D.ID_PEDIDO = P.ID_PEDIDO WHERE D.ID_PRODUCTO = '" & .Fields("ID_PRODUCTO") & "' AND (D.ENTREGADO = '0' OR D.ENTREGADO = 'R') AND P.COMENTARIO = 'MAXIMOS Y MINIMOS' AND P.PIDIO = 'SISTEMA'"
                Set tRs2 = cnn.Execute(sqlQuery)
                If Not (tRs2.EOF And tRs2.BOF) Then
                    CantPed = tRs2.Fields("CANTIDAD")
                End If
                If CDbl(.Fields("FALTANTE")) - CantPed > 0 Then
                    CantMin = CDbl(.Fields("FALTANTE")) - CantPed
                    CantMin = CantMin - CDbl(.Fields("FALTANTEMIN"))
                    
                    Set tLi = lvwFaltas.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                        tLi.SubItems(1) = .Fields("DESCRIPCION")
                        If CantMin > 0 Then
                            tLi.SubItems(2) = CantMin
                            tLi.SubItems(3) = CDbl(.Fields("FALTANTEMIN"))
                        Else
                            tLi.SubItems(2) = 0
                            tLi.SubItems(3) = CDbl(.Fields("FALTANTE")) - CantPed
                        End If
                        tLi.SubItems(4) = CDbl(.Fields("FALTANTE")) - CantPed
                        tLi.SubItems(5) = "A3"
                        tLi.SubItems(6) = "M"
                End If
                .MoveNext
                Ultimo = lvwFaltas.ListItems.Count
                If lvwFaltas.ListItems.Count > 0 Then lvwFaltas.ListItems.Item(Ultimo).Checked = True
            Loop
        End If
        .Close
    End With
    Fin = lvwFaltas.ListItems.Count
    sBuscar = "SELECT V.NOMBRE, V.ID_PRODUCTO, V.DESCRIPCION, V.CANTIDAD, V.CANT_MIN, NO_CONTRATO, ISNULL(E.CANTIDAD, 0) AS EXISTENCIA FROM VSTOTALVENTALICITA AS V JOIN ALMACEN3 AS A ON V.ID_PRODUCTO = A.ID_PRODUCTO LEFT JOIN EXISTENCIAS AS E ON V.ID_PRODUCTO = E.ID_PRODUCTO WHERE E.SUCURSAL = 'BODEGA'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Cont = 1
                Do While Cont <= Fin
                    Cont = Cont + 1
                Loop
                If Cont > Fin Then
                    If CDbl(.Fields("CANT_MIN")) - CDbl(.Fields("CANTIDAD")) > 0 Then
                        Resu = ((CDbl(.Fields("CANT_MIN")) - CDbl(.Fields("CANTIDAD"))) * 0.2)
                        If Resu = 0 Then
                            Resu = 1
                        End If
                        sqlQuery = "SELECT ISNULL(SUM(D.CANTIDAD), 0) AS CANTIDAD FROM DETALLE_PEDIDO AS D JOIN PEDIDO AS P ON D.ID_PEDIDO = P.ID_PEDIDO WHERE D.ID_PRODUCTO = '" & .Fields("ID_PRODUCTO") & "' AND (D.ENTREGADO = '0' OR D.ENTREGADO = 'R') AND P.COMENTARIO = 'LICITACION' AND P.PIDIO = 'SISTEMA'"
                        Set tRs2 = cnn.Execute(sqlQuery)
                        If Not (tRs2.EOF And tRs2.BOF) Then
                            CantPed = tRs2.Fields("CANTIDAD")
                        End If
                        tRs2.Close
                        CantPed = CantPed + .Fields("EXISTENCIA")
                        If Resu - CantPed > 0 Then
                            Set tLi = lvwFaltas.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                                tLi.SubItems(1) = .Fields("ID_PRODUCTO")
                                tLi.SubItems(2) = Resu - CantPed
                                tLi.SubItems(3) = 0
                                tLi.SubItems(4) = Resu - CantPed
                                tLi.SubItems(5) = "A3"
                                tLi.SubItems(6) = "L"
                        End If
                    End If
                End If
                .MoveNext
                Ultimo = lvwFaltas.ListItems.Count
                lvwFaltas.ListItems.Item(Ultimo).Checked = True
                'lvwFaltas.ListItems.Item(Ultimo).Bold = True
                'lvwFaltas.ListItems.Item(Ultimo).ForeColor = vbRed
            Loop
        End If
        .Close
    End With
    
    If lvwFaltas.ListItems.Count = 0 Then
        MsgBox "NO HAY PRODUCTOS ABAJO DEL MINIMO", vbInformation, "MENSAJE DEL SISTEMA"
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


Private Sub Command2_Click()
On Error GoTo ManejaError

    Dim Cont As Integer
    Dim NoRe As Integer
    Dim nPedido As Integer
    Dim cProducto As String
    Dim Uno As String
    Dim cMin As Double
    Dim cMax As Double
    Dim CantPed As Double
    Dim CantPedF As Double
    Dim exist As Integer
    Dim Almacen As String
    Dim sqlQuery As String
    
    cMin = 0
    cMax = 0
    CantPed = 0
        If MsgBox("¿DESEA HACER UN PEDIDO POR LOS ARTICULOS SELECCIONADOS?", vbQuestion + vbYesNo + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then
            Uno = "S"
            NoRe = Me.lvwFaltas.ListItems.Count
            For Cont = 1 To NoRe
                If Me.lvwFaltas.ListItems.Item(Cont).Checked = True Then
                    CantPedF = CDbl(lvwFaltas.ListItems.Item(Cont).SubItems(4))
                        If (CantPedF > 0) Then
                            If Me.lvwFaltas.ListItems.Item(Cont).SubItems(6) = "M" Then
                                If Uno = "S" Then
                                    sqlQuery = "INSERT INTO PEDIDO (SUCURSAL, PIDIO, FECHA, TIPO, COMENTARIO) VALUES ('BODEGA', 'SISTEMA', '" & Format(Date, "dd/mm/yyyy") & "', 'I', 'MAXIMOS Y MINIMOS')"
                                    cnn.Execute (sqlQuery)
                                    sqlQuery = "SELECT TOP 1 ID_PEDIDO FROM PEDIDO ORDER BY ID_PEDIDO DESC"
                                    Set tRs = cnn.Execute(sqlQuery)
                                    nPedido = tRs.Fields("ID_PEDIDO")
                                    Uno = "N"
                                End If
                                If CDbl(lvwFaltas.ListItems.Item(Cont).SubItems(4)) > 0 Then
                                    sqlQuery = "INSERT INTO DETALLE_PEDIDO (ID_PRODUCTO, CANTIDAD, ID_PEDIDO, DESCRIPCION, ALMACEN) VALUES ('" & Trim(lvwFaltas.ListItems.Item(Cont).SubItems(1)) & "', " & Replace(lvwFaltas.ListItems.Item(Cont).SubItems(4), ",", ".") & ", " & nPedido & ", '" & lvwFaltas.ListItems.Item(Cont).SubItems(1) & "', '" & lvwFaltas.ListItems.Item(Cont).SubItems(5) & "')"
                                    cnn.Execute (sqlQuery)
                                End If
                                If CDbl(lvwFaltas.ListItems.Item(Cont).SubItems(3)) > 0 Then
                                    sqlQuery = "INSERT INTO DETALLE_PEDIDO (ID_PRODUCTO, CANTIDAD, ID_PEDIDO, DESCRIPCION, ALMACEN, ENTREGADO) VALUES ('" & Trim(lvwFaltas.ListItems.Item(Cont).SubItems(1)) & "', " & Replace(lvwFaltas.ListItems.Item(Cont).SubItems(3), ",", ".") & ", " & nPedido & ", '" & lvwFaltas.ListItems.Item(Cont).SubItems(1) & "', '" & lvwFaltas.ListItems.Item(Cont).SubItems(5) & "', 'R')"
                                    cnn.Execute (sqlQuery)
                                    sBuscar = "INSERT INTO REQUISICION (FECHA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, ACTIVO, CONTADOR, COTIZADA, ALMACEN, URGENTE) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', '" & Trim(lvwFaltas.ListItems.Item(Cont).SubItems(1)) & "' , '" & lvwFaltas.ListItems.Item(Cont).SubItems(1) & "'," & Replace(lvwFaltas.ListItems.Item(Cont).SubItems(1), ",", ".") & ", 0, 0, 0, '" & lvwFaltas.ListItems.Item(Cont).SubItems(5) & "', 'S')"
                                    cnn.Execute (sqlQuery)
                                End If
                            End If
                        End If
                End If
            Next Cont
            Uno = "S"
            NoRe = Me.lvwFaltas.ListItems.Count
            For Cont = 1 To NoRe
                If Me.lvwFaltas.ListItems.Item(Cont).Checked = True Then
                    CantPedF = CDbl(lvwFaltas.ListItems.Item(Cont).SubItems(4))
                        If (CantPedF > 0) Then
                            If Me.lvwFaltas.ListItems.Item(Cont).SubItems(6) = "L" Then
                                If Uno = "S" Then
                                    sqlQuery = "INSERT INTO PEDIDO (SUCURSAL, PIDIO, FECHA, TIPO, COMENTARIO) VALUES ('BODEGA', 'SISTEMA', '" & Format(Date, "dd/mm/yyyy") & "', 'I', 'LICITACION')"
                                    cnn.Execute (sqlQuery)
                                    sqlQuery = "SELECT TOP 1 ID_PEDIDO FROM PEDIDO ORDER BY ID_PEDIDO DESC"
                                    Set tRs = cnn.Execute(sqlQuery)
                                    nPedido = tRs.Fields("ID_PEDIDO")
                                    Uno = "N"
                                End If
                                sqlQuery = "INSERT INTO DETALLE_PEDIDO (ID_PRODUCTO, CANTIDAD, ID_PEDIDO, DESCRIPCION, ALMACEN) VALUES ('" & Trim(lvwFaltas.ListItems.Item(Cont).SubItems(1)) & "', " & CantPedF & ", " & nPedido & ", '" & lvwFaltas.ListItems.Item(Cont).SubItems(1) & "', '" & lvwFaltas.ListItems.Item(Cont).SubItems(5) & "')"
                                cnn.Execute (sqlQuery)
                            End If
                        End If
                End If
            Next Cont
            MsgBox "PEDIDO TERMINADO", vbInformation, "MENSAJE DEL SISTEMA"
            lvwFaltas.ListItems.Clear
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

Private Sub Form_Load()

On Error GoTo ManejaError

    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Dim sPathBase As String
    sPathBase = "LINUX"
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SaccPass*20;Persist Security Info=True;User ID=AdmSACC2210;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    With lvwFaltas
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "PRODUCTO", 2000
        .ColumnHeaders.Add , , "DESCRIPCION", 2000
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "C. URGENTE", 1000
        .ColumnHeaders.Add , , "C. TOTAL", 1000
        .ColumnHeaders.Add , , "ALMACEN", 1440
        .ColumnHeaders.Add , , "TIPO", 500
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
'Private Sub Imprimit()
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.CurrentX = (Printer.Width - Printer.TextWidth(Menu.Text5(0).Text)) / 2
'    Printer.Print Menu.Text5(0).Text
'    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & Menu.Text5(8).Text)) / 2
'    Printer.Print "R.F.C. " & Menu.Text5(8).Text
'    Printer.CurrentX = (Printer.Width - Printer.TextWidth(Menu.Text5(1).Text & " COL. " & Menu.Text5(4).Text)) / 2
'    Printer.Print Menu.Text5(1).Text & " COL. " & Menu.Text5(4).Text
'    Printer.CurrentX = (Printer.Width - Printer.TextWidth(Menu.Text5(5).Text & ", " & Menu.Text5(6).Text & " C.P. " & Menu.Text5(9).Text)) / 2
'    Printer.Print Menu.Text5(5).Text & ", " & Menu.Text5(6).Text & " C.P. " & Menu.Text5(9).Text
'    Printer.Print "         Fecha : " & Date
'End Sub


