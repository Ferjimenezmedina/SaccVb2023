VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmPedClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos de Clientes"
   ClientHeight    =   6090
   ClientLeft      =   825
   ClientTop       =   1995
   ClientWidth     =   10365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   9120
      ScaleHeight     =   6795
      ScaleWidth      =   1995
      TabIndex        =   16
      Top             =   0
      Width           =   2055
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   17
         Top             =   4680
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
            TabIndex        =   18
            Top             =   960
            Width           =   975
         End
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmPedClientes.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "FrmPedClientes.frx":030A
            Top             =   240
            Width           =   705
         End
      End
   End
   Begin VB.CommandButton CmdGuardar 
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
      Left            =   7800
      Picture         =   "FrmPedClientes.frx":1DBC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton CmdQuitar 
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
      Left            =   6480
      Picture         =   "FrmPedClientes.frx":478E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3201
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
   Begin VB.CommandButton CmdAceptar 
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
      Left            =   6240
      Picture         =   "FrmPedClientes.frx":7160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox TxtCantidad 
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox TxtCanExis 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox TxtClvProd 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton CmdBucar 
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
      Left            =   7800
      Picture         =   "FrmPedClientes.frx":9B32
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por Descripcion"
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Clave"
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      Top             =   120
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox TxtBusProd 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   4335
   End
   Begin MSComctlLib.ListView LvwProd 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3201
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
   Begin VB.Label Label4 
      Caption         =   "Cantidad del Pedido"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad en Existencia"
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Clave de Producto :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar Producto :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "FrmPedClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim ind As Integer
Dim NOORD2 As String
Private Sub AgreLis()
On Error GoTo ManejaError
    Me.cmdGuardar.Enabled = True
    Dim LI As ListItem
    Dim NumeroRegistros As Integer
    Dim Conta As Integer
    Dim Exp As Integer
    Exp = 0
    NumeroRegistros = ListView2.ListItems.Count
    For Conta = 1 To NumeroRegistros
        If ListView2.ListItems.Item(Conta) = TxtClvProd.Text Then
            ListView2.ListItems.Item(Conta).SubItems(1) = Format(CDbl(ListView2.ListItems.Item(Conta).SubItems(1)) + CDbl(txtCantidad.Text), "0.00")
            If CDbl(ListView2.ListItems.Item(Conta).SubItems(1)) <= CDbl(ListView2.ListItems.Item(Conta).SubItems(2)) Then
                ListView2.ListItems.Item(Conta).SubItems(3) = "0.00"
            Else
                ListView2.ListItems.Item(Conta).SubItems(3) = CDbl(ListView2.ListItems.Item(Conta).SubItems(1)) - CDbl(ListView2.ListItems.Item(Conta).SubItems(2))
            End If
            Exp = 1
        End If
    Next Conta
    If Exp = 0 Then
        Set LI = ListView2.ListItems.Add(, , TxtClvProd.Text & "")
            LI.SubItems(1) = txtCantidad.Text & ""
            LI.SubItems(2) = TxtCanExis.Text & ""
        If CDbl(txtCantidad.Text) - CDbl(TxtCanExis.Text) <= 0 Then
            LI.SubItems(3) = "0.00"
        Else
            LI.SubItems(3) = CDbl(txtCantidad.Text) - CDbl(TxtCanExis.Text)
        End If
    End If
    TxtClvProd.Text = ""
    TxtCanExis.Text = ""
    txtCantidad.Text = ""
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub CmdBucar_Click()
On Error GoTo ManejaError
    If TxtBusProd.Text <> "" Then
        Buscar
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub CmdAceptar_Click()
On Error GoTo ManejaError
    AgreLis
    LvwProd.SetFocus
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub CmdQuitar_Click()
On Error GoTo ManejaError
    If ind <> 0 Then
        ListView2.ListItems.Remove (ind)
        ind = 0
        Me.cmdQuitar.Enabled = False
        ListView2.SetFocus
    End If
    If ListView2.ListItems.Count = 0 Then
        Me.cmdGuardar.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub CmdGuardar_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim NumeroRegistros As Integer
    NumeroRegistros = ListView2.ListItems.Count
    Dim Conta As Integer
    For Conta = 1 To NumeroRegistros
        sBuscar = "INSERT INTO PED_CLIEN_DETALLE (ID_PRODUCTO, NO_PEDIDO, CANTIDAD_PEDIDA, CANTIDAD_EXISTENCIA, CANTIDAD_PENDIENTE) VALUES ('" & ListView2.ListItems(Conta) & "', " & NOORD2 & ", " & CDbl(ListView2.ListItems(Conta).SubItems(1)) & ", " & CDbl(ListView2.ListItems(Conta).SubItems(2)) & ", " & CDbl(ListView2.ListItems(Conta).SubItems(3)) & ");"
        cnn.Execute (sBuscar)
        If CDbl(ListView2.ListItems(Conta).SubItems(3)) <> 0 Then
            sBuscar = "INSERT INTO REQUISICION (FECHA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, ACTIVO, CONTADOR, COTIZADA) VALUES ('" & Now & "', '" & ListView2.ListItems(Conta) & "' , 'DESCRIPCION'," & CDbl(ListView2.ListItems(Conta).SubItems(3)) & ", 0, 0, 0)"
            cnn.Execute (sBuscar)
        End If
        If CDbl(ListView2.ListItems(Conta).SubItems(1)) - CDbl(ListView2.ListItems(Conta).SubItems(3)) <> 0 Then
            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD - " & CDbl(ListView2.ListItems(Conta).SubItems(1)) - CDbl(ListView2.ListItems(Conta).SubItems(3)) & " WHERE ID_PRODUCTO = '" & ListView2.ListItems(Conta) & "' AND SUCURSAL = 'BODEGA'"
            cnn.Execute (sBuscar)
        End If
    Next Conta
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
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    If Me.Option1.Value = True Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & TxtBusProd.Text & "%' ORDER BY ID_PRODUCTO"
    Else
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION FROM ALMACEN3 WHERE DESCRIPCION LIKE '%" & TxtBusProd.Text & "%' ORDER BY ID_PRODUCTO"
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            MsgBox "No se ha encontrado el producto"
        Else
            LvwProd.ListItems.Clear
            '.MoveFirst
                Do While Not .EOF
                    Set tLi = LvwProd.ListItems.Add(, , Trim(.Fields("ID_PRODUCTO")) & "")
                    tLi.SubItems(1) = .Fields("DESCRIPCION") & ""
                    .MoveNext
                Loop
        End If
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

Private Sub Form_Load()

End Sub

Private Sub LvwProd_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    TxtClvProd.Text = Item
    Dim sBuscar As String
    Dim tRs2 As Recordset
    sBuscar = "SELECT CANTIDAD, SUCURSAL FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & TxtClvProd.Text & "' AND SUCURSAL = 'BODEGA'"
    Set tRs2 = cnn.Execute(sBuscar)
    With tRs2
        If (.BOF And .EOF) Then
            TxtCanExis.Text = "0.00"
        Else
            .MoveFirst
            If Not IsNull(.Fields("CANTIDAD")) Then
                TxtCanExis.Text = .Fields("CANTIDAD")
            Else
                TxtCanExis.Text = "0.00"
            End If
            .MoveNext
        End If
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
Private Sub LvwProd_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        txtCantidad.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Me.cmdQuitar.Enabled = True
    ind = Item.Index
    Me.cmdQuitar.SetFocus
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub ListView2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.cmdQuitar.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub TxtBusProd_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And TxtBusProd.Text <> "" Then
        Buscar
        LvwProd.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub TxtBusProd_GotFocus()
    TxtBusProd.BackColor = &HFFE1E1
End Sub
Private Sub TxtBusProd_LostFocus()
      TxtBusProd.BackColor = &H80000005
End Sub
Private Sub TxtClvProd_Change()
On Error GoTo ManejaError
    If txtCantidad.Text = "" Or TxtClvProd.Text = "" Then
        Me.CmdAceptar.Enabled = False
    Else
        Me.CmdAceptar.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub TxtCantidad_Change()
On Error GoTo ManejaError
    If txtCantidad.Text = "" Or TxtClvProd.Text = "" Then
        Me.CmdAceptar.Enabled = False
    Else
        Me.CmdAceptar.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtCantidad_GotFocus()
On Error GoTo ManejaError
    txtCantidad.BackColor = &HFFE1E1
    txtCantidad.SelStart = 0
    txtCantidad.SelLength = Len(txtCantidad.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub TxtCantidad_LostFocus()
      txtCantidad.BackColor = &H80000005
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        AgreLis
        TxtBusProd.SetFocus
    End If
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub

Sub Reconexion()
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Const sPathBase As String = "LINUX"
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
        MsgBox "LA CONEXION SE RESABLECIO CON EXITO. PUEDE CONTINUAR CON SU TRABAJO.", vbInformation, "MENSAJE DEL SISTEMA"
    End With
Exit Sub
ManejaError:
    MsgBox Err.Number & Err.Description
    Err.Clear
    If MsgBox("NO PUDIMOS RESTABLECER LA CONEXIÓN, ¿DESEA REINTENTARLO?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
End Sub


