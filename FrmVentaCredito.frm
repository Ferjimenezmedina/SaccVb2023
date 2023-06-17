VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmVentaCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Venta a Credito"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   8640
      ScaleHeight     =   5115
      ScaleWidth      =   1875
      TabIndex        =   15
      Top             =   0
      Width           =   1935
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   975
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmVentaCredito.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "FrmVentaCredito.frx":030A
            Top             =   240
            Width           =   705
         End
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
            TabIndex        =   17
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   5640
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5760
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command3 
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
      Left            =   7320
      Picture         =   "FrmVentaCredito.frx":1DBC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   2520
      Width           =   4815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmVentaCredito.frx":478E
      Left            =   4320
      List            =   "FrmVentaCredito.frx":47A1
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2990
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
   Begin VB.CommandButton Command1 
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
      Left            =   7320
      Picture         =   "FrmVentaCredito.frx":47B9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label5 
      Caption         =   "Credito Disponible :"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Limite de Credito :"
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Dias de Credito :"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "FrmVentaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim ClvCliente As Integer
Dim ClvUsuario As String
Dim VarDescuento As String
Dim DesClente As Double
Private Sub Combo1_Click()
    If Text2.Text = "" Or Text3.Text = "" Or Combo1.Text = "" Or Text4.Text = "" Then
        Me.Command3.Enabled = False
    Else
        Me.Command3.Enabled = True
    End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text2.Text <> "" Then
        Me.Command3.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    Dim FechaVence As String
    Text3.Text = Replace(Text3.Text, ",", ".")
    Text4.Text = Replace(Text4.Text, ",", ".")
    FechaVence = Date + CDbl(Combo1.Text)
    Dim sBuscar As String
    Dim tRs As Recordset
    DesClente = Replace(DesClente, ",", ".")
    If VarDescuento = "" Then
        VarDescuento = "0.00"
    End If
    sBuscar = "INSERT INTO CUENTAS (PAGADA, ID_CLIENTE, ID_USUARIO, FECHA, DIAS_CREDITO, FECHA_VENCE, DESCUENTO, SUCURSAL) VALUES ( 'N', " & ClvCliente & ", '" & ClvUsuario & "', '" & Date & "', " & CDbl(Combo1.Text) & ", '" & FechaVence & "', " & CDbl(VarDescuento) & ", '" & Menu.Text4(0).Text & "');"
    cnn.Execute (sBuscar)
    sBuscar = "SELECT ID_CUENTA FROM CUENTAS ORDER BY ID_CUENTA DESC"
    Set tRs = cnn.Execute(sBuscar)
    Text5.Text = tRs.Fields("ID_CUENTA")
    sBuscar = "INSERT INTO VENTAS (ID_CLIENTE, NOMBRE, SUCURSAL, ID_USUARIO, FECHA, DIAS_CREDITO, FECHA_VENCE, DESCUENTO) VALUES (" & ClvCliente & ", '" & Text2.Text & "', '" & Menu.Text4(0).Text & "', '" & ClvUsuario & "', '" & Date & "', " & CDbl(Combo1.Text) & ", '" & FechaVence & "', " & CDbl(VarDescuento) & ");"
    cnn.Execute (sBuscar)
    sBuscar = "SELECT ID_VENTA FROM VENTAS WHERE SUCURSAL = '" & Menu.Text4(0).Text & "' ORDER BY ID_VENTA DESC"
    Set tRs = cnn.Execute(sBuscar)
    Text6.Text = tRs.Fields("ID_VENTA")
    sBuscar = "INSERT INTO CUENTA_VENTA (ID_VENTA, ID_CUENTA) VALUES (" & Text6.Text & ", " & Text5.Text & ");"
    cnn.Execute (sBuscar)
    FrmRegVentCred.Show vbModal
    'Unload Me
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
    Me.Command3.Enabled = False
    ClvUsuario = Menu.Text1(0).Text
    Me.Command1.Enabled = False
    Const sPathBase As String = "LINUX"
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Clave del Cliente", 1800
        .ColumnHeaders.Add , , "Nombre", 7450
        .ColumnHeaders.Add , , "RFC", 2450
        .ColumnHeaders.Add , , "Limite de credito", 2450
        .ColumnHeaders.Add , , "Descuento", 1550
        .ColumnHeaders.Add , , "Dias de Credito", 1550
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
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_CLIENTE, NOMBRE, RFC, DIAS_CREDITO, LIMITE_CREDITO, DESCUENTO FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND LIMITE_CREDITO <> 0"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            Text1.Text = ""
            MsgBox "No se encontro cliente con credito registrado a ese nombre"
        Else
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & ""
                If Not IsNull(.Fields("RFC")) Then tLi.SubItems(2) = .Fields("RFC") & ""
                If Not IsNull(.Fields("LIMITE_CREDITO")) Then tLi.SubItems(3) = .Fields("LIMITE_CREDITO") & ""
                If .Fields("DESCUENTO") = "" Then
                    tLi.SubItems(4) = "0.00"
                Else
                    tLi.SubItems(4) = .Fields("DESCUENTO") & ""
                End If
                If Not IsNull(.Fields("DIAS_CREDITO")) Then tLi.SubItems(5) = .Fields("DIAS_CREDITO") & ""
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
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Text2.Text = Item.SubItems(1)
    Text3.Text = Item.SubItems(3)
    Combo1.Text = Item.SubItems(5)
    If Combo1.Text <> "" Then
        Combo1.Enabled = False
    Else
        Combo1.Enabled = True
    End If
    Dim sBuscar As String
    Dim ACUM As Double
    Dim tRs As Recordset
    sBuscar = "SELECT TOTAL_COMPRA FROM CUENTAS WHERE ID_CLIENTE = " & Item
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                If .Fields("TOTAL_COMPRA") <> Null Then
                    ACUM = ACUM + CDbl(.Fields("TOTAL_COMPRA"))
                End If
                .MoveNext
            Loop
        End If
    End With
    sBuscar = "SELECT CANT_ABONO FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & Item
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                If .Fields("TOTAL_COMPRA") <> Null Then
                    ACUM = ACUM - CDbl(.Fields("CANT_ABONO"))
                End If
                .MoveNext
            Loop
        End If
    End With
    Text4.Text = CDbl(Text3.Text) - ACUM
    ClvCliente = Item
    VarDescuento = Item.SubItems(4)
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
Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo1.SetFocus
    End If
End Sub
Private Sub Text1_Change()
    If Text1.Text = "" Then
        Me.Command1.Enabled = False
    Else
        Me.Command1.Enabled = True
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Text1.Text <> "" Then
        If KeyAscii = 13 Then
            Buscar
            ListView1.SetFocus
        End If
    End If
End Sub
Private Sub Command1_Click()
    Buscar
End Sub
Private Sub Text2_Change()
    If Text2.Text = "" Or Text3.Text = "" Or Combo1.Text = "" Or Text4.Text = "" Then
        Me.Command3.Enabled = False
    Else
        Me.Command3.Enabled = True
    End If
End Sub
Private Sub Text3_Change()
    If Text2.Text = "" Or Text3.Text = "" Or Combo1.Text = "" Or Text4.Text = "" Then
        Me.Command3.Enabled = False
    Else
        Me.Command3.Enabled = True
    End If
End Sub
Private Sub Text4_Change()
    If Text2.Text = "" Or Text3.Text = "" Or Combo1.Text = "" Or Text4.Text = "" Then
        Me.Command3.Enabled = False
    Else
        Me.Command3.Enabled = True
    End If
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

