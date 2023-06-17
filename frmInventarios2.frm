VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmInventarios2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventarios"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11040
      Picture         =   "frmInventarios2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7560
      Visible         =   0   'False
      Width           =   90
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
      Height          =   195
      Left            =   10320
      Picture         =   "frmInventarios2.frx":29D2
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7560
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   10800
      TabIndex        =   31
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   10560
      TabIndex        =   29
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   11160
      MultiLine       =   -1  'True
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9240
      TabIndex        =   26
      Top             =   1560
      Width           =   975
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmInventarios2.frx":53A4
         MousePointer    =   99  'Custom
         Picture         =   "frmInventarios2.frx":56AE
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EXCEL"
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
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10320
      TabIndex        =   24
      Top             =   1560
      Width           =   975
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
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmInventarios2.frx":71F0
         MousePointer    =   99  'Custom
         Picture         =   "frmInventarios2.frx":74FA
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdBuscar2 
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
      Left            =   4680
      Picture         =   "frmInventarios2.frx":95DC
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtIdProducto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   3855
   End
   Begin VB.CommandButton cmdActualizar 
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
      Left            =   6360
      Picture         =   "frmInventarios2.frx":BFAE
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Index           =   0
      Left            =   4320
      TabIndex        =   10
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdTraer 
         Caption         =   "Traer"
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
         Left            =   360
         Picture         =   "frmInventarios2.frx":E980
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtTraer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No. Inventario"
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
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   11175
      Begin MSComctlLib.ListView lvwInventario 
         Height          =   3615
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4095
      Begin VB.CheckBox chk3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Almacen 3"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chk2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Almacen 2"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chk1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Almacen 1"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cboSucursales 
         Height          =   315
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   3375
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Nuevo"
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
         Left            =   1440
         Picture         =   "frmInventarios2.frx":11352
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SUCURSALES"
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
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   9360
      TabIndex        =   30
      Top             =   7560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar Producto"
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
      TabIndex        =   22
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label lblSuc 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7680
      TabIndex        =   20
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lblFec 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7680
      TabIndex        =   19
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblInv 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7680
      TabIndex        =   18
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sucursal:"
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
      Left            =   6600
      TabIndex        =   17
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha:"
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
      Left            =   6600
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Inventario:"
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
      Left            =   6600
      TabIndex        =   15
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblEstado 
      BackColor       =   &H00FFFFFF&
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
      Left            =   120
      TabIndex        =   5
      Top             =   6960
      Width           =   11175
   End
End
Attribute VB_Name = "frmInventarios2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Dim StrRep As String
Dim StrRep2 As String
Dim StrRep3 As String
Dim bBanderaLvw As Byte
Dim Cont As Integer
Dim NoRe As Integer
Dim ID As Integer
Private Sub cboSucursales_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.cmdBuscar.Value = True
    End If
End Sub
Private Sub Chk1_Click()
    chk2.Value = 0
    chk3.Value = 0
    txtIdProducto.SetFocus
End Sub
Private Sub Chk2_Click()
    chk1.Value = 0
    chk3.Value = 0
    txtIdProducto.SetFocus
End Sub
Private Sub chk3_Click()
    chk2.Value = 0
    chk1.Value = 0
    txtIdProducto.SetFocus
End Sub
Private Sub cmdActualizar_Click()
    If Puede_Guardar Then
        If Me.lblInv.Caption = "" Then
            sqlQuery = "INSERT INTO INVENTARIOS (ID_SUCURSAL, FECHA) VALUES('" & Me.cboSucursales.Text & "', '" & Format(Date, "dd/mm/yyyy") & "')"
            cnn.Execute (sqlQuery)
            sqlQuery = "SELECT TOP 1 ID_INVENTARIO FROM INVENTARIOS ORDER BY ID_INVENTARIO DESC"
            'InputBox "", "", sqlQuery
            Set tRs = cnn.Execute(sqlQuery)
            ID = tRs.Fields("ID_INVENTARIO")
            NoRe = Me.lvwInventario.ListItems.Count
            For Cont = 1 To NoRe
                If Val(Me.lvwInventario.ListItems.Item(Cont)) <> 0 Then
                    sqlQuery = "INSERT INTO INVENTARIO_DETALLE (ID_INVENTARIO, ID_PRODUCTO, CANTIDAD) VALUES('" & ID & "', '" & Me.lvwInventario.ListItems.Item(Cont).SubItems(2) & "', " & Me.lvwInventario.ListItems.Item(Cont) & ")"
                    'InputBox "", "", sqlQuery
                    cnn.Execute (sqlQuery)
                End If
            Next Cont
            Traer_Ultimo_Id
        Else
            ID = Val(Me.lblInv.Caption)
            NoRe = Me.lvwInventario.ListItems.Count
            For Cont = 1 To NoRe
                If Me.lvwInventario.ListItems(Cont).SubItems(6) = 0 Then
                    If Val(Me.lvwInventario.ListItems.Item(Cont)) <> 0 Then
                        sqlQuery = "INSERT INTO INVENTARIO_DETALLE (ID_INVENTARIO, ID_PRODUCTO, CANTIDAD) VALUES('" & ID & "', '" & Me.lvwInventario.ListItems.Item(Cont).SubItems(2) & "', " & Me.lvwInventario.ListItems.Item(Cont) & ")"
                        'InputBox "", "", sqlQuery
                        cnn.Execute (sqlQuery)
                    End If
                ElseIf Me.lvwInventario.ListItems(Cont).SubItems(6) = 2 Then
                    sqlQuery = "UPDATE INVENTARIO_DETALLE SET CANTIDAD = " & Val(Me.lvwInventario.ListItems(Cont)) & " WHERE ID_INVENTARIO = " & Me.lblInv.Caption & " AND ID_PRODUCTO = '" & Me.lvwInventario.ListItems(Cont).SubItems(2) & "'"
                    'InputBox "", "", sqlQuery
                    cnn.Execute (sqlQuery)
                End If
            Next Cont
        End If
        Me.lblEstado.Caption = "Inventario guardado a las " & Time
    End If
End Sub
Private Sub cmdBuscar_Click()
    If Puede_Buscar Then
        Me.lblEstado.Caption = "Buscando productos... por favor espere..."
        Me.lblEstado.ForeColor = vbBlack
        DoEvents
        Llenar_Lista_Existencias Me.cboSucursales.Text
        Me.lblEstado.Caption = "Listo, " & Me.lvwInventario.ListItems.Count & " registros encontrados"
        Me.lblEstado.ForeColor = vbBlack
        DoEvents
        Me.cmdBuscar2.Enabled = True
        Me.lvwInventario.SetFocus
    End If
End Sub
Private Sub Command3_Click()
    Dim sqlQuery As String
    Command3.Visible = False
End Sub
Private Sub Command4_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim cant As Integer
    If chk1 = 1 Then
        sBuscar = "SELECT ID_PRODUCTO,CANTIDAD,SUCURSAL from vsinvalm1 WHERE SUCURSAL = '" & cboSucursales.Text & "' AND ID_PRODUCTO='" & Text1.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = '" & Text3.Text & "' WHERE ID_PRODUCTO = '" & Text1.Text & "'  AND  SUCURSAL = '" & cboSucursales.Text & "'"
            Set tRs = cnn.Execute(sBuscar)
        End If
    End If
    If chk2 = 1 Then
        sBuscar = "SELECT ID_PRODUCTO,CANTIDAD,SUCURSAL from vsinvalm2 WHERE SUCURSAL = '" & cboSucursales.Text & "' AND ID_PRODUCTO='" & Text1.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
               sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = '" & Text3.Text & "' WHERE ID_PRODUCTO = '" & Text1.Text & "' AND  SUCURSAL = '" & cboSucursales.Text & "'"
               Set tRs = cnn.Execute(sBuscar)
        End If
    End If
    If chk3 = 1 Then
       sBuscar = "SELECT ID_PRODUCTO,CANTIDAD,SUCURSAL from vsinvalm3 WHERE SUCURSAL = '" & cboSucursales.Text & "' AND ID_PRODUCTO='" & Text1.Text & "'"
       Set tRs = cnn.Execute(sBuscar)
       If Not (tRs.EOF And tRs.BOF) Then
           sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = '" & Text3.Text & "' WHERE ID_PRODUCTO = '" & Text1.Text & "' AND  SUCURSAL = '" & cboSucursales.Text & "'"
           Set tRs = cnn.Execute(sBuscar)
       End If
    End If
    Text1.Text = ""
    Text3.Text = ""
    Command4.Visible = False
End Sub
Private Sub cmdBuscar2_Click()
On Error GoTo ManejaError
    txtIdProducto.SetFocus
    Me.lblEstado.Caption = "Buscando " & Me.txtIdProducto.Text & "... por favor espere..."
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
        If Me.lblSuc.Caption = "" Then
            Llenar_Lista_Existencias_Producto Me.cboSucursales.Text
        Else
            Llenar_Lista_Existencias_Producto Me.lblSuc.Caption
        End If
        If Me.lblInv.Caption <> "" Then
            Me.lblEstado.Caption = "Cargando inventarios..."
            Me.lblEstado.ForeColor = vbBlack
            DoEvents
            Llenar_Lista_Inventario Me.lblInv.Caption
        End If
    Me.lblEstado.Caption = "Listo, " & Me.lvwInventario.ListItems.Count & " registros encontrados"
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    Me.lvwInventario.SetFocus
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdTraer_Click()
    If Puede_Traer Then
        Me.lblEstado.Caption = "Buscando productos... por favor espere..."
        DoEvents
        If Traer_Datos = True Then
            DoEvents
            Llenar_Lista_Existencias Me.lblSuc.Caption
            Me.lblEstado.Caption = "Cargando inventarios..."
            DoEvents
            Llenar_Lista_Inventario Me.txtTraer.Text
            DoEvents
        End If
        Me.lblEstado.Caption = "Listo, " & Me.lvwInventario.ListItems.Count & " registros encontrados"
        DoEvents
        Me.cmdBuscar2.Enabled = True
        Me.lvwInventario.SetFocus
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With lvwInventario
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CANT INV", 2200
        .ColumnHeaders.Add , , "ID_EXISTENCIA", 0
        .ColumnHeaders.Add , , "PRODUCTO", 2000
        .ColumnHeaders.Add , , "Descripcion", 5000
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "MAXIMO", 1000
        .ColumnHeaders.Add , , "MINIMO", 1000
        .ColumnHeaders.Add , , "MARCA", 1500
        .ColumnHeaders.Add , , "ESTADO", 0
        .ColumnHeaders.Add , , "LOCALIZACION", 2200
    End With
    sqlQuery = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sqlQuery)
        With tRs
            Me.cboSucursales.Clear
            Do While Not .EOF
                Me.cboSucursales.AddItem .Fields("NOMBRE")
                .MoveNext
            Loop
        End With
        Me.cboSucursales.ListIndex = 0
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Puede_Buscar() As Boolean
On Error GoTo ManejaError
    If Me.cboSucursales.Text = "" Then
        MsgBox "INTRODUSCA LA SUCURSAL", vbInformation, "SACC"
        Me.cboSucursales.SetFocus
        Puede_Buscar = False
        Exit Function
    End If
    If Me.chk1.Value = 0 And Me.chk2.Value = 0 And Me.chk3.Value = 0 Then
        MsgBox "DEBE SELECCIONAR AL MENOS UN ALMACEN", vbInformation, "SACC"
        Me.cboSucursales.SetFocus
        Puede_Buscar = False
        Exit Function
    End If
    Puede_Buscar = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Lista_Existencias(Id_Sucursal As String)
On Error GoTo ManejaError
    sqlQuery = ""
    If Me.chk3.Value = 1 Then
        sqlQuery = "select ID_EXISTENCIA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, MARCA, C_MINIMA, C_MAXIMA,LOCALIZACION from vsinvalm3 WHERE SUCURSAL = '" & Id_Sucursal & "'"
    End If
    If Me.chk1.Value = 1 Then
        If sqlQuery <> "" Then
            sqlQuery = sqlQuery & " Union select ID_EXISTENCIA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, MARCA, C_MINIMA, C_MAXIMA,LOCALIZACION from vsinvalm1 WHERE SUCURSAL = '" & Id_Sucursal & "'"
        Else
            sqlQuery = "select ID_EXISTENCIA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, MARCA, C_MINIMA, C_MAXIMA,LOCALIZACION from vsinvalm1 WHERE SUCURSAL = '" & Id_Sucursal & "'"
        End If
    End If
    If Me.chk2.Value = 1 Then
        If sqlQuery <> "" Then
            sqlQuery = sqlQuery & " Union select ID_EXISTENCIA, ID_PRODUCTO, DESCRIPCION, CANTIDAD,C_MINIMA, C_MAXIMA, MARCA,LOCALIZACION from vsinvalm2 WHERE SUCURSAL = '" & Id_Sucursal & "'"
             Else
            sqlQuery = "select ID_EXISTENCIA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, MARCA, C_MINIMA, C_MAXIMA,LOCALIZACION from vsinvalm2 WHERE SUCURSAL = '" & Id_Sucursal & "'"
           
        End If
        With tRs
            Me.lvwInventario.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwInventario.ListItems.Add(, , "")
                If Not IsNull(.Fields("ID_EXISTENCIA")) Then tLi.SubItems(1) = .Fields("ID_EXISTENCIA")
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = .Fields("ID_PRODUCTO")
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(3) = .Fields("Descripcion")
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(4) = .Fields("CANTIDAD")
                If Not IsNull(.Fields("C_MAXIMA")) Then tLi.SubItems(5) = .Fields("C_MAXIMA")
                If Not IsNull(.Fields("C_MINIMA")) Then tLi.SubItems(6) = .Fields("C_MINIMA")
                If Not IsNull(.Fields("MARCA")) Then tLi.SubItems(7) = .Fields("MARCA")
                tLi.SubItems(8) = "0"
                .MoveNext
            Loop
        End With
    End If
    sqlQuery = sqlQuery & " ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        Me.lvwInventario.ListItems.Clear
        Do While Not .EOF
            Set tLi = Me.lvwInventario.ListItems.Add(, , "")
            If Not IsNull(.Fields("ID_EXISTENCIA")) Then tLi.SubItems(1) = .Fields("ID_EXISTENCIA")
            If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = .Fields("ID_PRODUCTO")
            If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(3) = .Fields("Descripcion")
            If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(4) = .Fields("CANTIDAD")
            If Not IsNull(.Fields("C_MINIMA")) Then tLi.SubItems(5) = .Fields("C_MINIMA")
            If Not IsNull(.Fields("C_MAXIMA")) Then tLi.SubItems(6) = .Fields("C_MAXIMA")
            If Not IsNull(.Fields("MARCA")) Then tLi.SubItems(7) = .Fields("MARCA")
            tLi.SubItems(8) = "0"
            If Not IsNull(.Fields("LOCALIZACION")) Then tLi.SubItems(9) = .Fields("LOCALIZACION")
            .MoveNext
        Loop
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image10_Click()
    If lvwInventario.ListItems.Count > 0 Then
        Dim StrCopi As String
        Dim Con As Integer
        Dim Con2 As Integer
        Dim NumColum As Integer
        Dim Ruta As String
        Me.CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "Guardar como"
        CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
        Me.CommonDialog1.ShowSave
        Ruta = Me.CommonDialog1.FileName
        StrCopi = "Cant_inv" & Chr(9) & "Producto" & Chr(9) & "Descripcion" & Chr(9) & "Cantidad" & Chr(9) & "Maximo" & Chr(9) & "Minimo" & Chr(9) & "Marca" & Chr(13)
        If Ruta <> "" Then
            NumColum = lvwInventario.ColumnHeaders.Count
            For Con = 1 To lvwInventario.ListItems.Count
                StrCopi = StrCopi & lvwInventario.ListItems.Item(Con) & Chr(13)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & lvwInventario.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(9)
            Next
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwInventario_Click()
On Error GoTo ManejaError
    Me.lvwInventario.StartLabelEdit
    If Me.lvwInventario.SelectedItem.SubItems(6) = 1 Then
        Me.lvwInventario.SelectedItem.SubItems(6) = 2
    ElseIf Me.lvwInventario.SelectedItem.SubItems(6) = 3 Then
        MsgBox "EL INVENTARIO YA FUE DESCONTADO", vbInformation, "SACC"
    End If
ManejaError:
    Err.Clear
End Sub
Private Sub lvwInventario_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwInventario.SortKey = ColumnHeader.Index - 1
    lvwInventario.Sorted = True
    lvwInventario.SortOrder = 1 Xor lvwInventario.SortOrder
End Sub
Private Sub lvwInventario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Me.lvwInventario.SelectedItem.SubItems(9) = 3 Then
            MsgBox "EL INVENTARIO YA FUE DESCONTADO", vbInformation, "SACC"
        Else
            Me.lvwInventario.StartLabelEdit
        End If
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
Function Puede_Traer() As Boolean
On Error GoTo ManejaError
    If Me.txtTraer.Text = "" Then
        MsgBox "INTRODUSCA EL NUMERO DE INVENTARIO", vbInformation, "SACC"
        Me.txtTraer.SetFocus
        Puede_Traer = False
        Exit Function
    End If
    Puede_Traer = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Traer_Datos() As Boolean
On Error GoTo ManejaError
    Traer_Datos = False
    sqlQuery = "SELECT * FROM INVENTARIOS WHERE ID_INVENTARIO = " & Me.txtTraer.Text & " AND ESTADO_ACTUAL = 'A'"
    'InputBox "", "", sqlQuery
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.EOF And .BOF) Then
            If Not IsNull(.Fields("ID_INVENTARIO")) Then Me.lblInv.Caption = .Fields("ID_INVENTARIO")
            If Not IsNull(.Fields("ID_SUCURSAL")) Then Me.lblSuc.Caption = .Fields("ID_SUCURSAL")
            If Not IsNull(.Fields("FECHA")) Then Me.lblFec.Caption = .Fields("FECHA")
            Traer_Datos = True
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Lista_Inventario(nInventario As Integer)
On Error GoTo ManejaError
    NoRe = Me.lvwInventario.ListItems.Count
    sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD, ESTADO_ACTUAL,C_MAXIMA,C_MINIMA FROM INVENTARIO_DETALLE WHERE ID_INVENTARIO = " & nInventario
    Set tRs = cnn.Execute(sqlQuery)
        With tRs
            Do While Not .EOF
                For Cont = 1 To NoRe
                    If Me.lvwInventario.ListItems(Cont).SubItems(2) = .Fields("ID_PRODUCTO") Then
                        If Not IsNull(.Fields("CANTIDAD")) Then Me.lvwInventario.ListItems.Item(Cont) = .Fields("CANTIDAD")
                        If .Fields("ESTADO_ACTUAL") = "I" Then
                            Me.lvwInventario.ListItems(Cont).SubItems(6) = "3"
                        Else
                            Me.lvwInventario.ListItems(Cont).SubItems(6) = "1"
                        End If
                    End If
                Next Cont
                .MoveNext
            Loop
        End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtIdProducto_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.cmdBuscar2.Value = True
    End If
 txtIdProducto.SetFocus
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtTraer_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Me.cmdTraer.Value = True
    Else
        Dim Valido As String
        Valido = "1234567890"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Traer_Ultimo_Id()
On Error GoTo ManejaError
    sqlQuery = "SELECT TOP 1 * FROM INVENTARIOS ORDER BY ID_INVENTARIO DESC"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.EOF And .BOF) Then
            If Not IsNull(.Fields("ID_INVENTARIO")) Then Me.lblInv.Caption = .Fields("ID_INVENTARIO")
            If Not IsNull(.Fields("ID_SUCURSAL")) Then Me.lblSuc.Caption = .Fields("ID_SUCURSAL")
            If Not IsNull(.Fields("FECHA")) Then Me.lblFec.Caption = .Fields("FECHA")
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Puede_Guardar() As Boolean
On Error GoTo ManejaError
    If Me.lvwInventario.ListItems.Count = 0 Then
        MsgBox "SELECCIONE LA SUCURSAL O INTRODUSCA EL NUMERO DE INVENTARIO", vbInformation, "SACC"
        Me.cboSucursales.SetFocus
        Puede_Guardar = False
        Exit Function
    End If
    Puede_Guardar = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Lista_Existencias_Producto(Id_Sucursal As String)
On Error GoTo ManejaError
    sqlQuery = ""
    If Me.chk3.Value = 1 Then
        sqlQuery = "SELECT EXISTENCIAS.ID_EXISTENCIA, EXISTENCIAS.ID_PRODUCTO, ALMACEN3.Descripcion, EXISTENCIAS.CANTIDAD, ALMACEN3.MARCA, ALMACEN3.C_MAXIMA, ALMACEN3.C_MINIMA FROM ALMACEN3, EXISTENCIAS WHERE EXISTENCIAS.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO AND EXISTENCIAS.SUCURSAL = '" & Id_Sucursal & "' AND EXISTENCIAS.ID_PRODUCTO LIKE '%" & Me.txtIdProducto.Text & "%'"
    End If
    If Me.chk1.Value = 1 Then
        If sqlQuery <> "" Then
            sqlQuery = sqlQuery & " UNION SELECT EXISTENCIAS.ID_EXISTENCIA, EXISTENCIAS.ID_PRODUCTO, ALMACEN1.Descripcion, EXISTENCIAS.CANTIDAD, ALMACEN1.MARCA, ALMACEN1.C_MAXIMA, ALMACEN1.C_MINIMA FROM ALMACEN1, EXISTENCIAS WHERE EXISTENCIAS.ID_PRODUCTO = ALMACEN1.ID_PRODUCTO AND EXISTENCIAS.SUCURSAL = '" & Id_Sucursal & "' AND EXISTENCIAS.ID_PRODUCTO LIKE '%" & Me.txtIdProducto.Text & "%'"
        Else
            sqlQuery = "SELECT EXISTENCIAS.ID_EXISTENCIA, EXISTENCIAS.ID_PRODUCTO, ALMACEN1.Descripcion, EXISTENCIAS.CANTIDAD, ALMACEN1.MARCA, ALMACEN1.C_MAXIMA, ALMACEN1.C_MINIMA FROM ALMACEN1, EXISTENCIAS WHERE EXISTENCIAS.ID_PRODUCTO = ALMACEN1.ID_PRODUCTO AND EXISTENCIAS.SUCURSAL = '" & Id_Sucursal & "' AND EXISTENCIAS.ID_PRODUCTO LIKE '%" & Me.txtIdProducto.Text & "%'"
        End If
    End If
    If Me.chk2.Value = 1 Then
        If sqlQuery <> "" Then
            sqlQuery = sqlQuery & " UNION SELECT EXISTENCIAS.ID_EXISTENCIA, EXISTENCIAS.ID_PRODUCTO, ALMACEN2.Descripcion, EXISTENCIAS.CANTIDAD, ALMACEN2.MARCA, ALMACEN2.C_MAXIMA, ALMACEN2.C_MINIMA FROM ALMACEN2, EXISTENCIAS WHERE EXISTENCIAS.ID_PRODUCTO = ALMACEN2.ID_PRODUCTO AND EXISTENCIAS.SUCURSAL = '" & Id_Sucursal & "' AND EXISTENCIAS.ID_PRODUCTO LIKE '%" & Me.txtIdProducto.Text & "%'"
        Else
            sqlQuery = "SELECT EXISTENCIAS.ID_EXISTENCIA, EXISTENCIAS.ID_PRODUCTO, ALMACEN2.Descripcion, EXISTENCIAS.CANTIDAD, ALMACEN2.MARCA, ALMACEN2.C_MAXIMA, ALMACEN2.C_MINIMA FROM ALMACEN2, EXISTENCIAS WHERE EXISTENCIAS.ID_PRODUCTO = ALMACEN2.ID_PRODUCTO AND EXISTENCIAS.SUCURSAL = '" & Id_Sucursal & "' AND EXISTENCIAS.ID_PRODUCTO LIKE '%" & Me.txtIdProducto.Text & "%'"
        End If
    End If
    sqlQuery = sqlQuery & " ORDER BY EXISTENCIAS.ID_PRODUCTO"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        Me.lvwInventario.ListItems.Clear
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Set tLi = Me.lvwInventario.ListItems.Add(, , "")
                If Not IsNull(.Fields("ID_EXISTENCIA")) Then tLi.SubItems(1) = .Fields("ID_EXISTENCIA")
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = .Fields("ID_PRODUCTO")
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(3) = .Fields("Descripcion")
                 If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(4) = .Fields("CANTIDAD")
                If Not IsNull(.Fields("C_MAXIMA")) Then tLi.SubItems(5) = .Fields("C_MAXIMA")
                If Not IsNull(.Fields("C_MINIMA")) Then tLi.SubItems(6) = .Fields("C_MINIMA")
                If Not IsNull(.Fields("MARCA")) Then tLi.SubItems(7) = .Fields("MARCA")
                tLi.SubItems(8) = "0"
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
