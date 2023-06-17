VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmVerCotizaClien 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Cotizaciones a Clientes"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   10560
      ScaleHeight     =   6315
      ScaleWidth      =   1155
      TabIndex        =   18
      Top             =   0
      Width           =   1215
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   19
         Top             =   4920
         Width           =   975
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmVerCotizaClien.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "FrmVerCotizaClien.frx":030A
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
            TabIndex        =   20
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5400
      Width           =   8175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4920
      Width           =   6255
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
      Left            =   9120
      Picture         =   "FrmVerCotizaClien.frx":1DBC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por Descripcion"
      Height          =   195
      Left            =   7080
      TabIndex        =   8
      Top             =   4200
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Clave"
      Height          =   195
      Left            =   5640
      TabIndex        =   7
      Top             =   4200
      Value           =   -1  'True
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   3255
      Left            =   5640
      TabIndex        =   6
      Top             =   720
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5741
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5640
      TabIndex        =   5
      Top             =   360
      Width           =   4695
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1815
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   5175
      _ExtentX        =   9128
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   5175
      _ExtentX        =   9128
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
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Precio Cotizó"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Tel."
      Height          =   255
      Left            =   8400
      TabIndex        =   14
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Direccion"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre del Proveedor"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   240
      X2              =   10320
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Buscar articulo"
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Productos"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Cotizaciones"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FrmVerCotizaClien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Private Sub Form_Load()
On Error GoTo ManejaError
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
        .ColumnHeaders.Add , , "NO. COTIZACION", 0
        .ColumnHeaders.Add , , "NOMBRE", 3700
        .ColumnHeaders.Add , , "DIRECCION", 1500
        .ColumnHeaders.Add , , "COLONIA", 1500
        .ColumnHeaders.Add , , "CIUDAD", 1500
        .ColumnHeaders.Add , , "TELEFONO", 1000
        .ColumnHeaders.Add , , "COMENTARIOS", 1000
    End With
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "NO. COTIZACION", 0
        .ColumnHeaders.Add , , "CLAVE", 1700
        .ColumnHeaders.Add , , "DESCRIPCION", 3500
        .ColumnHeaders.Add , , "CANTIDAD", 1500
    End With
    With ListView3
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "NO. COTIZACION", 0
        .ColumnHeaders.Add , , "CLAVE", 1700
        .ColumnHeaders.Add , , "DESCRIPCION", 3500
        .ColumnHeaders.Add , , "CANTIDAD", 1500
    End With
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_COTIZACION, NOMBRE, DIRECCION, COLONIA, CIUDAD, TELEFONO, COMENTARIOS FROM COTIZACION WHERE PENDIENTE = 'S'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        ListView1.ListItems.Clear
        tRs.MoveFirst
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_COTIZACION") & "")
            tLi.SubItems(1) = tRs.Fields("NOMBRE") & ""
            tLi.SubItems(2) = tRs.Fields("DIRECCION") & ""
            tLi.SubItems(3) = tRs.Fields("COLONIA") & ""
            tLi.SubItems(4) = tRs.Fields("CIUDAD") & ""
            tLi.SubItems(5) = tRs.Fields("TELEFONO") & ""
            tLi.SubItems(6) = tRs.Fields("COMENTARIOS") & ""
            tRs.MoveNext
        Loop
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
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_COTIZACION, ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM COTIZACION_DETALLE WHERE ID_COTIZACION = " & Item
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        ListView2.ListItems.Clear
        tRs.MoveFirst
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_COTIZACION") & "")
            tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO") & ""
            tLi.SubItems(2) = tRs.Fields("DESCRIPCION") & ""
            tLi.SubItems(3) = tRs.Fields("CANTIDAD") & ""
            tRs.MoveNext
        Loop
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
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    If Option1.Value Then
        sBuscar = "SELECT ID_COTIZACION, ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM COTIZACION_DETALLE WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%'"
    Else
        sBuscar = "SELECT ID_COTIZACION, ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM COTIZACION_DETALLE WHERE DESCRIPCION LIKE '%" & Text1.Text & "%'"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        ListView3.ListItems.Clear
        tRs.MoveFirst
        Do While Not tRs.EOF
            Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_COTIZACION") & "")
            tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO") & ""
            tLi.SubItems(2) = tRs.Fields("DESCRIPCION") & ""
            tLi.SubItems(3) = tRs.Fields("CANTIDAD") & ""
            tRs.MoveNext
        Loop
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
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tRs1 As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_PROVEEDOR, PRECIO FROM COTIZACION_PROV WHERE ID_COTIZACION = " & Item & " AND ID_PRODUCTO = '" & Item.SubItems(1) & "' AND DESCRIPCION = '" & Item.SubItems(2) & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        tRs.MoveFirst
        sBuscar = "SELECT NOMBRE, TELEFONO1, DIRECCION FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs.Fields("ID_PROVEEDOR")
        Set tRs1 = cnn.Execute(sBuscar)
        Text5.Text = tRs.Fields("PRECIO")
        Text2.Text = tRs1.Fields("NOMBRE")
        Text4.Text = tRs1.Fields("TELEFONO1")
        Text3.Text = tRs1.Fields("DIRECCION")
    Else
        Text5.Text = ""
        Text2.Text = ""
        Text4.Text = ""
        Text3.Text = ""
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
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tRs1 As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_PROVEEDOR, PRECIO FROM COTIZACION_PROV WHERE ID_COTIZACION = " & Item & " AND ID_PRODUCTO = '" & Item.SubItems(1) & "' AND DESCRIPCION = '" & Item.SubItems(2) & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        tRs.MoveFirst
        sBuscar = "SELECT NOMBRE, TELEFONO1, DIRECCION FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs.Fields("ID_PROVEEDOR")
        Set tRs1 = cnn.Execute(sBuscar)
        Text5.Text = tRs.Fields("PRECIO")
        Text2.Text = tRs1.Fields("NOMBRE")
        Text4.Text = tRs1.Fields("TELEFONO1")
        Text3.Text = tRs1.Fields("DIRECCION")
    Else
        Text5.Text = ""
        Text2.Text = ""
        Text4.Text = ""
        Text3.Text = ""
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
Private Sub Text1_Change()
    If Text1.Text <> "" Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
    End If
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text1.Text <> "" Then
        Buscar
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
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub Text3_GotFocus()
    Text3.BackColor = &HFFE1E1
End Sub
Private Sub Text3_LostFocus()
    Text3.BackColor = &H80000005
End Sub
Private Sub Text4_GotFocus()
    Text4.BackColor = &HFFE1E1
End Sub
Private Sub Text4_LostFocus()
    Text4.BackColor = &H80000005
End Sub
Private Sub Text5_GotFocus()
    Text5.BackColor = &HFFE1E1
End Sub
Private Sub Text5_LostFocus()
    Text5.BackColor = &H80000005
End Sub
