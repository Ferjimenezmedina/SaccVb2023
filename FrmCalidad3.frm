VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCalidad3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calidad"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10800
      TabIndex        =   0
      Top             =   3120
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCalidad3.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCalidad3.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7920
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmCalidad3.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblEstado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCantidad"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblArticulo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblComanda"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtNumArticulo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdRemanofactura"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtEdo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtNoSirve"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.TextBox txtNoSirve 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   3840
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtEdo 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   3840
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton cmdRemanofactura 
         Cancel          =   -1  'True
         Caption         =   "Remanofactura"
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
         Left            =   8760
         Picture         =   "FrmCalidad3.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   10095
         Begin MSComctlLib.ListView lvwJR 
            Height          =   2055
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3625
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2055
            Left            =   4440
            TabIndex        =   6
            Top             =   240
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   3625
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.TextBox txtNumArticulo 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   3720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Comanda"
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
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Articulo"
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
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblComanda 
         Caption         =   "Comanda"
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
         Left            =   1320
         TabIndex        =   14
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblArticulo 
         Caption         =   "Articulo"
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
         Left            =   1320
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblCantidad 
         Caption         =   "Cantidad"
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
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Cantidad"
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
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   600
         Width           =   3855
      End
   End
End
Attribute VB_Name = "FrmCalidad3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Dim pro As String
Private Sub cmdRemanofactura_Click()
    Dim Art_rema As String
    Dim cProducto As String
    Dim NoArt As Integer
    Dim CantRegresa As Integer
    Dim Tipo As String
    Art_rema = Strings.Left(Me.lblArticulo.Caption, Strings.Len(Me.lblArticulo.Caption) - 3) & "REM"
    If MsgBox("DESEA MANDAR A AUTORIZAR " & Art_rema, vbQuestion + vbYesNo + vbDefaultButton1, "SACC") = vbYes Then
        Do
            CantRegresa = Val(InputBox("INTRODUSCA LA CANTIDAD DE " & lblArticulo.Caption & " QUE SE VA MANDAR A REMA", "SACC", Val(lblCantidad.Caption)))
        Loop Until CantRegresa <= Val(Me.lblCantidad.Caption) And CantRegresa <> 0
        sqlQuery = "SELECT Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO = '" & Art_rema & "'"
        Set tRs = cnn.Execute(sqlQuery)
        If Not (tRs.EOF And tRs.BOF) Then
            cProducto = tRs.Fields("Descripcion")
            sqlQuery = "SELECT TOP 1 ARTICULO FROM COMANDAS_DETALLES_2 WHERE ID_COMANDA = " & Me.lblComanda.Caption & " ORDER BY ARTICULO DESC"
            Set tRs = cnn.Execute(sqlQuery)
            If Mid(Art_rema, 3, 1) = "T" Then
                Tipo = "T" 'Toner
            ElseIf Mid(Art_rema, 3, 1) = "I" Then
                Tipo = "I" 'Tinta
            End If
            sqlQuery = "INSERT INTO COMANDAS_DETALLES_2 (ID_COMANDA,ARTICULO,ID_PRODUCTO,CANTIDAD,ESTADO_ACTUAL,TIPO) VALUES (" & lblComanda.Caption & ", " & Val(tRs.Fields("ARTICULO")) + 1 & ", '" & Art_rema & "', " & CantRegresa & ", 'Z', '" & Tipo & "');"
            cnn.Execute (sqlQuery)
            If txtEdo = "M" Then
                sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'N', CANTIDAD_NO_SIRVIO = " & CantRegresa + CDbl(txtNoSirve.Text) & ", FECHA_FIN = '" & Format(Date, "dd/mm/yyyy") & "' WHERE ID_COMANDA = " & Me.lblComanda.Caption & " AND ARTICULO = " & Me.txtNumArticulo.Text ' PONER BIEN ESTADO
            Else
                sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'N', CANTIDAD_NO_SIRVIO = " & CantRegresa & ", FECHA_FIN = '" & Format(Date, "dd/mm/yyyy") & "' WHERE ID_COMANDA = " & Me.lblComanda.Caption & " AND ARTICULO = " & Me.txtNumArticulo.Text ' PONER BIEN ESTADO
            End If
            cnn.Execute (sqlQuery)
            sqlQuery = "SELECT COUNT(TEMPORAL) TEMPORAL FROM JR_TEMPORALES WHERE ID_COMANDA = " & Me.lblComanda.Caption & " AND ID_REPARACION = '" & Me.lblArticulo.Caption & "'"
            Set tRs = cnn.Execute(sqlQuery)
            If tRs.Fields("TEMPORAL") = 0 Then
                sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD, ISNULL(E.CANTIDAD, 0) AS EXISTENCIA FROM JUEGO_REPARACION AS J LEFT JOIN EXISTENCIAS AS E ON J.ID_PRODUCTO = E.ID_PRODUCTO WHERE ID_REPARACION = '" & cProducto & "'"
                Set tRs = cnn.Execute(sqlQuery)
            Else
                sqlQuery = "SELECT J.ID_PRODUCTO, J.CANTIDAD, ISNULL(E.CANTIDAD, 0) AS EXISTENCIA FROM JR_TEMPORALES AS J LEFT JOIN EXISTENCIAS AS E ON J.ID_PRODUCTO = E.ID_PRODUCTO WHERE ID_REPARACION = '" & cProducto & "' AND ID_COMANDA = " & Me.txtNumArticulo.Text
                Set tRs = cnn.Execute(sqlQuery)
            End If
            'With tRs
            '    Do While Not .EOF
            '        If .Fields("EXISTENCIA") = "0" Then
            '            sqlQuery = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, SUCURSAL, CANTIDAD) VALUES(" & (CantRegresa * .Fields("CANTIDAD")) & ", '" & .Fields("Id_Producto") & "', 'BODEGA');"
            '        Else
            '            sqlQuery = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD + " & (CantRegresa * .Fields("CANTIDAD")) & " WHERE ID_PRODUCTO = '" & .Fields("Id_Producto") & "' AND SUCURSAL = 'BODEGA'"
            '        End If
            '        cnn.Execute (sqlQuery)
            '        .MoveNext
            '    Loop
            'End With
            Unload Me
        Else
            MsgBox "ESTE ARTICULO NO TIENE JUEGO DE REPARACION PARA REMANOFACTURA", vbInformation, "SACC"
        End If
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
    With lvwJR
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Pieza", 2500
        .ColumnHeaders.Add , , "Descripcion", 2900
        .ColumnHeaders.Add , , "Cantidad", 1200
    End With
    If txtNoSirve.Text = "" Then txtNoSirve.Text = "0"
    If Tiene_JR_Temporal Then
        Llenar_Liata_JR_Temporal
        Me.lblEstado.Caption = "MODIFICADO"
        Me.lblEstado.ForeColor = vbRed
    Else
        Llenar_Liata_JR
        Me.lblEstado.Caption = "NORMAL"
        Me.lblEstado.ForeColor = vbBlue
    End If
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "I_S", 800
        .ColumnHeaders.Add , , "ID_COMANDA", 1500
        .ColumnHeaders.Add , , "ID_PRODUCTO", 1500
        .ColumnHeaders.Add , , "CANTIDAD", 800
        .ColumnHeaders.Add , , "CANTIDAD_NO_SIRVIO", 800
        .ColumnHeaders.Add , , "ESTADO_ACTUAL", 0
    End With
    Llenar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Tiene_JR_Temporal() As Boolean
On Error GoTo ManejaError
    sqlQuery = "SELECT COUNT(TEMPORAL) TEMPORAL FROM JR_TEMPORALES WHERE ID_COMANDA = " & Me.lblComanda.Caption & " AND ID_REPARACION = '" & Me.lblArticulo.Caption & "'"
    Set tRs = cnn.Execute(sqlQuery)
    If tRs.Fields("TEMPORAL") = 0 Then
        Tiene_JR_Temporal = False
    Else
        Tiene_JR_Temporal = True
    End If
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Liata_JR_Temporal()
On Error GoTo ManejaError
    sqlQuery = "SELECT * FROM JR_TEMPORALES WHERE ID_COMANDA = " & Me.lblComanda.Caption & " AND ID_REPARACION = '" & Me.lblArticulo.Caption & "'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        While Not .EOF
            Set tLi = Me.lvwJR.ListItems.Add(, , Trim(.Fields("ID_PRODUCTO")))
            If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = .Fields("CANTIDAD")
            .MoveNext
        Wend
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar()
On Error GoTo ManejaError
    sqlQuery = "SELECT * FROM COMANDAS_DETALLES_2 WHERE ID_COMANDA = " & Me.lblComanda.Caption & ""
    Set tRs = cnn.Execute(sqlQuery)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("I_S"))
            tLi.SubItems(1) = tRs.Fields("ID_COMANDA")
            tLi.SubItems(2) = tRs.Fields("ID_PRODUCTO")
            tLi.SubItems(3) = tRs.Fields("CANTIDAD")
            tLi.SubItems(4) = tRs.Fields("CANTIDAD_NO_SIRVIO")
            tLi.SubItems(5) = tRs.Fields("ESTADO_ACTUAL")
            tRs.MoveNext
        Loop
     End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Liata_JR()
On Error GoTo ManejaError
    sqlQuery = "SELECT * FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & Me.lblArticulo.Caption & "'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        While Not .EOF
            Set tLi = Me.lvwJR.ListItems.Add(, , Trim(.Fields("ID_PRODUCTO")))
            If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = .Fields("CANTIDAD")
            .MoveNext
        Wend
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Puede_Subir_Inventario() As Boolean
On Error GoTo ManejaError
    Puede_Subir_Inventario = False
    sqlQuery = "SELECT TIPO FROM COMANDAS_2 WHERE ID_COMANDA = " & Me.lblComanda.Caption
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not IsNull(.Fields("TIPO")) Then
            If .Fields("TIPO") = "P" Then
                Puede_Subir_Inventario = True
            End If
        End If
    End With
    Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Subir_Inventarios()
On Error GoTo ManejaError
    Dim nExi As Double
    sqlQuery = "SELECT ID_EXISTENCIA, CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Me.lblArticulo.Caption & "' AND SUCURSAL = 'BODEGA'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.EOF And .BOF) Then
            nExi = .Fields("ID_EXISTENCIA")
            sqlQuery = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs.Fields("CANTIDAD")) + CDbl(lblCantidad.Caption) - CDbl(txtNoSirve.Text) & "WHERE ID_EXISTENCIA = " & nExi
            cnn.Execute (sqlQuery)
        Else
            sqlQuery = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & Me.lblArticulo.Caption & "', " & CDbl(lblCantidad.Caption) - CDbl(txtNoSirve.Text) & ", 'BODEGA');"
            cnn.Execute (sqlQuery)
        End If
    End With
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView1_Click()
    pro = Item
End Sub
Private Sub ListView1_DblClick()
    pro = Item
End Sub
