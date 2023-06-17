VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPerdidas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERDIDAS"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10320
      TabIndex        =   16
      Top             =   120
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmPerdidas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmPerdidas.frx":030A
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
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdTerminar 
      Caption         =   "Terminar"
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
      Left            =   6000
      Picture         =   "frmPerdidas.frx":23EC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   360
      Width           =   1095
   End
   Begin VB.CheckBox chkSelec 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccionar todo"
      Height          =   195
      Left            =   8640
      TabIndex        =   13
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8415
      Begin VB.CommandButton cmdDescontar 
         Caption         =   "Descontar"
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
         Left            =   7080
         Picture         =   "frmPerdidas.frx":4DBE
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtTraer 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
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
         Left            =   4680
         Picture         =   "frmPerdidas.frx":7790
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1095
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
         Left            =   1920
         TabIndex        =   12
         Top             =   240
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
         Left            =   1920
         TabIndex        =   11
         Top             =   720
         Width           =   975
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
         Left            =   1920
         TabIndex        =   10
         Top             =   480
         Width           =   975
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
         Left            =   3000
         TabIndex        =   9
         Top             =   240
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
         Left            =   3000
         TabIndex        =   8
         Top             =   720
         Width           =   2415
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
         Left            =   3000
         TabIndex        =   7
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   11175
      Begin MSComctlLib.ListView lvwInventario 
         Height          =   5175
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   9128
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
      TabIndex        =   14
      Top             =   6960
      Width           =   11175
   End
End
Attribute VB_Name = "frmPerdidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Dim Cont As Integer
Dim NoRe As Integer
Private Sub chkSelec_Click()
    If Me.lvwInventario.ListItems.Count <> 0 Then
        NoRe = Me.lvwInventario.ListItems.Count
        For Cont = 1 To NoRe
            If Me.lvwInventario.ListItems.Item(Cont).Checked = False Then
                Me.lvwInventario.ListItems.Item(Cont).Checked = True
            Else
                Me.lvwInventario.ListItems.Item(Cont).Checked = False
            End If
        Next Cont
    End If
End Sub
Private Sub cmdDescontar_Click()
    If Puede_Descontar Then
        Me.lblEstado.Caption = "Descontando inventario... por favor espere"
        Me.lblEstado.ForeColor = vbBlack
        NoRe = Me.lvwInventario.ListItems.Count
        For Cont = 1 To NoRe
            If Me.lvwInventario.ListItems.Item(Cont).Checked = True Then
                sqlQuery = "INSERT INTO PERDIDAS (ID_PRODUCTO, ID_INVENTARIO, FECHA, CANTIDAD, PRECIO_UNITARIO) VALUES ('" & Me.lvwInventario.ListItems.Item(Cont) & "', " & Me.txtTraer.Text & ", '" & Format(Date, "dd/mm/yyyy") & "', " & Me.lvwInventario.ListItems.Item(Cont).SubItems(6) & ", " & Replace(FormatNumber(Me.lvwInventario.ListItems.Item(Cont).SubItems(4), 2, vbUseDefault, vbUseDefault, vbFalse), ",", "") & ")"
                cnn.Execute (sqlQuery)
                sqlQuery = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD - " & Me.lvwInventario.ListItems.Item(Cont).SubItems(6) & " WHERE ID_PRODUCTO = '" & Me.lvwInventario.ListItems.Item(Cont) & "' AND SUCURSAL = '" & Me.lblSuc.Caption & "'"
                cnn.Execute (sqlQuery)
                sqlQuery = "UPDATE INVENTARIO_DETALLE SET ESTADO_ACTUAL = 'I' WHERE ID_INVENTARIO = " & Me.lblInv.Caption & " AND ID_PRODUCTO = '" & Me.lvwInventario.ListItems.Item(Cont) & "'"
                cnn.Execute (sqlQuery)
            End If
        Next Cont
        Llenar_Lista_Perdida
        Me.lblEstado.Caption = "Inventario descontado"
        Me.lblEstado.ForeColor = vbBlue
    Else
        MsgBox "SELECCIONE LOS PRODUCTOS A DESCONTAR", vbInformation, "SACC"
        If Me.txtTraer.Text = "" Then
            Me.txtTraer.SetFocus
        Else
            Me.lvwInventario.SetFocus
        End If
    End If
End Sub
Private Sub cmdTraer_Click()
    If Puede_Traer Then
        Traer_Datos
        Llenar_Lista_Perdida
    End If
End Sub
Sub Llenar_Lista_Perdida()
On Error GoTo ManejaError
    Me.lblEstado.Caption = "Cargando datos del inventario"
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    sqlQuery = "SELECT D.ID_PRODUCTO, D.CANTIDAD AS INVENTARIO, V3.Descripcion, V3.ID_EXISTENCIA, V3.CANTIDAD AS EXISTENCIA, V3.PRECIO_COSTO " & _
    "FROM INVENTARIO_DETALLE AS D RIGHT JOIN vsinvalm3 AS V3 ON D.ID_PRODUCTO = V3.ID_PRODUCTO " & _
    "WHERE D.ID_INVENTARIO = " & Me.txtTraer.Text & " AND V3.SUCURSAL = '" & Me.lblSuc.Caption & "' AND D.CANTIDAD - V3.CANTIDAD <> 0 AND D.ESTADO_ACTUAL = 'A'" & _
    "UNION SELECT D.ID_PRODUCTO, D.CANTIDAD AS INVENTARIO, V3.Descripcion, V3.ID_EXISTENCIA, V3.CANTIDAD AS EXISTENCIA, V3.PRECIO_COSTO " & _
    "FROM INVENTARIO_DETALLE AS D RIGHT JOIN vsinvalm2 AS V3 ON D.ID_PRODUCTO = V3.ID_PRODUCTO " & _
    "WHERE D.ID_INVENTARIO = " & Me.txtTraer.Text & " AND V3.SUCURSAL = '" & Me.lblSuc.Caption & "' AND D.CANTIDAD - V3.CANTIDAD <> 0 AND D.ESTADO_ACTUAL = 'A'" & _
    "UNION SELECT D.ID_PRODUCTO, D.CANTIDAD AS INVENTARIO, V3.Descripcion, V3.ID_EXISTENCIA, V3.CANTIDAD AS EXISTENCIA, V3.PRECIO_COSTO " & _
    "FROM INVENTARIO_DETALLE AS D RIGHT JOIN vsinvalm1 AS V3 ON D.ID_PRODUCTO = V3.ID_PRODUCTO " & _
    "WHERE D.ID_INVENTARIO = " & Me.txtTraer.Text & " AND V3.SUCURSAL = '" & Me.lblSuc.Caption & "' AND D.CANTIDAD - V3.CANTIDAD <> 0 AND D.ESTADO_ACTUAL = 'A'" & _
    "ORDER BY D.ID_PRODUCTO"
    Set tRs = cnn.Execute(sqlQuery)
    Me.lvwInventario.ListItems.Clear
    With tRs
        Do While Not .EOF
            Set tLi = Me.lvwInventario.ListItems.Add(, , .Fields("ID_PRODUCTO"))
            If Not IsNull(.Fields("INVENTARIO")) Then tLi.SubItems(2) = .Fields("INVENTARIO")
            If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion")
            If Not IsNull(.Fields("EXISTENCIA")) Then tLi.SubItems(3) = .Fields("EXISTENCIA")
            If Not IsNull(.Fields("PRECIO_COSTO")) Then tLi.SubItems(4) = FormatNumber(.Fields("PRECIO_COSTO"), 2, vbUseDefault, vbUseDefault, vbTrue)
            If Not IsNull(.Fields("ID_EXISTENCIA")) Then tLi.SubItems(5) = .Fields("ID_EXISTENCIA")
            tLi.SubItems(6) = (.Fields("EXISTENCIA") - .Fields("INVENTARIO"))
            tLi.SubItems(7) = FormatNumber((.Fields("EXISTENCIA") - .Fields("INVENTARIO")) * .Fields("PRECIO_COSTO"), 2, vbFalse, vbFalse, vbTrue)
            .MoveNext
        Loop
    End With
    Me.lblEstado.Caption = "Listo"
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Puede_Traer() As Boolean
On Error GoTo ManejaError
    If Me.txtTraer.Text = "" Then
        Me.lblEstado.Caption = "Introdusca el numero de inventario"
        Me.lblEstado.ForeColor = vbRed
        Me.txtTraer.SetFocus
        Puede_Traer = False
        Exit Function
    End If
    Me.lblEstado.Caption = "Buscando inventario... por favor espere"
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    sqlQuery = "SELECT COUNT(I.ID_INVENTARIO) ID_INVENTARIO FROM INVENTARIOS AS I JOIN INVENTARIO_DETALLE AS D ON I.ID_INVENTARIO = D.ID_INVENTARIO WHERE D.ESTADO_ACTUAL = 'A' AND I.ESTADO_ACTUAL = 'A' AND D.ID_INVENTARIO = " & Me.txtTraer.Text
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_INVENTARIO") >= 1 Then
            Me.lblEstado.Caption = .Fields("ID_INVENTARIO") & " registros activos"
            Me.lblEstado.ForeColor = vbBlue
        Else
            Me.lblEstado.Caption = "No hay registros... introdusca otro numero de inventario"
            Me.lblEstado.ForeColor = vbRed
            Me.txtTraer.SetFocus
            Puede_Traer = False
            Exit Function
        End If
    End With
    Puede_Traer = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With Me.lvwInventario
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE PRODUCTO", 1700
        .ColumnHeaders.Add , , "Descripcion", 3700
        .ColumnHeaders.Add , , "C INV", 1000
        .ColumnHeaders.Add , , "C EXISTENCIA", 1000
        .ColumnHeaders.Add , , "PRECIO UNITARIO", 1000
        .ColumnHeaders.Add , , "ID EXISTENCIA", 0
        .ColumnHeaders.Add , , "PERDIDA", 1000
        .ColumnHeaders.Add , , "TOTAL PERDIDA", 1200
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub txtTraer_GotFocus()
    Me.txtTraer.SelStart = 0
    Me.txtTraer.SelLength = Len(Me.txtTraer.Text)
End Sub
Private Sub txtTraer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.cmdTraer.Value = True
    End If
End Sub
Sub Traer_Datos()
On Error GoTo ManejaError
    Me.lblEstado.Caption = "Buscando datos del inventario"
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    sqlQuery = "SELECT I.ID_SUCURSAL, I.FECHA FROM INVENTARIOS AS I WHERE I.ID_INVENTARIO = " & Me.txtTraer.Text & " AND ESTADO_ACTUAL = 'A'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.EOF And .BOF) Then
            If Not IsNull(.Fields("ID_SUCURSAL")) Then Me.lblSuc.Caption = .Fields("ID_SUCURSAL")
            If Not IsNull(.Fields("FECHA")) Then Me.lblFec.Caption = .Fields("FECHA")
            Me.lblInv.Caption = Me.txtTraer.Text
            Me.lvwInventario.SetFocus
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Puede_Descontar() As Boolean
On Error GoTo ManejaError
    Puede_Descontar = False
    NoRe = Me.lvwInventario.ListItems.Count
    For Cont = 1 To NoRe
        If Me.lvwInventario.ListItems.Item(Cont).Checked = True Then
            Puede_Descontar = True
        End If
    Next Cont
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub cmdTerminar_Click()
    If Puede_Terminar Then
        If MsgBox("¿DESEA TERMINAR EL INVETNARIO?", vbDefaultButton1 + vbYesNo, "SACC") = vbYes Then
            sqlQuery = "UPDATE INVENTARIOS SET ESTADO_ACTUAL = 'I' WHERE ID_INVENTARIO = " & Me.lblInv.Caption
            cnn.Execute (sqlQuery)
            Me.lblInv.Caption = ""
            Me.lblSuc.Caption = ""
            Me.lblFec.Caption = ""
            Me.lblEstado.Caption = ""
            Me.lvwInventario.ListItems.Clear
        End If
    End If
End Sub
Function Puede_Terminar() As Boolean
On Error GoTo ManejaError
    If Me.lblInv.Caption = "" Then
        MsgBox "INTRODUSCA EL NUMERO DE INVENTARIO", vbInformation, "SACC"
        Me.txtTraer.SetFocus
        Puede_Terminar = False
        Exit Function
    End If
    Puede_Terminar = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
