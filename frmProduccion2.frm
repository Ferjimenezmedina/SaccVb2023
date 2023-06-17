VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmProduccion2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Producción"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7800
      TabIndex        =   13
      Top             =   3240
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmProduccion2.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmProduccion2.frx":030A
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
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmProduccion2.frx":23EC
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
      Tab(0).Control(8)=   "cmdRemanofactura"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdTerminar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.Frame Frame1 
         Height          =   2415
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   7095
         Begin MSComctlLib.ListView lvwJR 
            Height          =   2055
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   3625
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            NumItems        =   0
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
         Left            =   6120
         Picture         =   "frmProduccion2.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemanofactura 
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
         Left            =   4560
         Picture         =   "frmProduccion2.frx":4DDA
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtNumArticulo 
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   3840
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
         Left            =   240
         TabIndex        =   12
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
         Left            =   240
         TabIndex        =   11
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
         Left            =   1440
         TabIndex        =   10
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
         Left            =   1440
         TabIndex        =   9
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
         Left            =   1440
         TabIndex        =   8
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
         Left            =   240
         TabIndex        =   7
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
         Left            =   3480
         TabIndex        =   6
         Top             =   600
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmProduccion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Private Sub cmdRemanofactura_Click()
On Error GoTo ManejaError
    Dim Art_rema As String
    Dim cProducto As String
    Dim NoArt As Integer
    Dim CantRegresa As Integer
    Dim Tipo As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim iAfectados As Long
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
            If CantRegresa = Val(Me.lblCantidad.Caption) Then
                sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'N', CANTIDAD_NO_SIRVIO = " & CantRegresa & ", FECHA_FIN = '" & Date & "' WHERE ID_COMANDA = " & Me.lblComanda.Caption & " AND ARTICULO = " & Me.txtNumArticulo.Text ' PONER BIEN ESTADO
            Else
                sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'M', CANTIDAD_NO_SIRVIO = " & CantRegresa & " WHERE ID_COMANDA = " & Me.lblComanda.Caption & " AND ARTICULO = " & Me.txtNumArticulo.Text ' PONER BIEN ESTADO
            End If
            cnn.Execute (sqlQuery)
            sqlQuery = "SELECT COUNT(TEMPORAL) TEMPORAL FROM JR_TEMPORALES WHERE ID_COMANDA = " & Me.lblComanda.Caption & " AND ID_REPARACION = '" & Me.lblArticulo.Caption & "'"
            Set tRs = cnn.Execute(sqlQuery)
            If tRs.Fields("TEMPORAL") = 0 Then
                sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & cProducto & "'"
                Set tRs = cnn.Execute(sqlQuery)
            Else
                sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JR_TEMPORALES WHERE ID_REPARACION = '" & cProducto & "' AND ID_COMANDA = " & Me.txtNumArticulo.Text
                Set tRs = cnn.Execute(sqlQuery)
            End If
            ' No se regresaran las existencias al enviar a Rema... los productos de la recarga seran retirados de la rema
            ' Modificado 06/Sep/2012 Armando H Valdez
            'sqlQuery = "SELECT CANTIDAD FROM EXISTENCIAS WHERE WHERE ID_PRODUCTO = '" & tRs.Fields("Id_Producto") & "' AND SUCURSAL = 'BODEGA'"
            'Set tRs2 = cnn.Execute(sqlQuery)
            'With tRs2
            '    Do While Not .EOF
            '        If .EOF And .BOF Then
            '            sqlQuery = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, SUCURSAL, CANTIDAD) VALUES(" & (CantRegresa * .Fields("CANTIDAD")) & ", '" & .Fields("Id_Producto") & "', 'BODEGA');"
            '            Set tRs = cnn.Execute(sqlQuery, iAfectados, adCmdText)
            '            If iAfectados < 1 Then
            '                sqlQuery = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, SUCURSAL, CANTIDAD) VALUES(" & (CantRegresa * .Fields("CANTIDAD")) & ", '" & .Fields("Id_Producto") & "', 'BODEGA');"
            '                Set tRs = cnn.Execute(sqlQuery, iAfectados, adCmdText)
            '                If iafectado < 1 Then
            '                    MsgBox "El producto " & .Fields("Id_Producto") & " no se regreso al inventario correctamente", vbInformation, "SACC"
            '                End If
            '            End If
            '        Else
            '            sqlQuery = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(.Fields("CANTIDAD")) + (CDbl(CantRegresa) * CDbl(.Fields("CANTIDAD"))) & " WHERE ID_PRODUCTO = '" & .Fields("Id_Producto") & "' AND SUCURSAL = 'BODEGA'"
            '            Set tRs = cnn.Execute(sqlQuery, iAfectados, adCmdText)
            '            If iAfectados < 1 Then
            '                sqlQuery = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(.Fields("CANTIDAD")) + (CDbl(CantRegresa) * CDbl(.Fields("CANTIDAD"))) & " WHERE ID_PRODUCTO = '" & .Fields("Id_Producto") & "' AND SUCURSAL = 'BODEGA'"
            '                Set tRs = cnn.Execute(sqlQuery, iAfectados, adCmdText)
            '                If iafectado < 1 Then
            '                    MsgBox "El producto " & .Fields("Id_Producto") & " no se regreso al inventario correctamente", vbInformation, "SACC"
            '                End If
            '            End If
            '        End If
            '        .MoveNext
            '     Loop
            'End With
            Unload Me
        Else
            MsgBox "ESTE ARTICULO NO TIENE JUEGO DE REPARACION PARA REMANOFACTURA", vbInformation, "SACC"
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdTerminar_Click()
On Error GoTo ManejaError
    sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'P' WHERE ID_COMANDA = " & Me.lblComanda.Caption & " AND ARTICULO = " & Me.txtNumArticulo.Text
    cnn.Execute (sqlQuery)
    frmProduccion.Llenar_Lista_Tinta
    Unload Me
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
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
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "PIEZA", 2440
        .ColumnHeaders.Add , , "Descripcion", 2880
        .ColumnHeaders.Add , , "CANTIDAD", 500
    End With
    Me.Caption = "JUEGO DE REPARACIÓN DE " & frmProduccion.txtArticulo.Text
    Me.lblArticulo.Caption = frmProduccion.txtArticulo.Text
    Me.lblComanda.Caption = frmProduccion.txtComanda.Text
    Me.lblCantidad.Caption = frmProduccion.txtCantidad.Text
    Me.txtNumArticulo.Text = frmProduccion.txtNumArticulo.Text
    If Tiene_JR_Temporal Then
        Llenar_Liata_JR_Temporal
        Me.lblEstado.Caption = "MODIFICADO"
        Me.lblEstado.ForeColor = vbRed
    Else
        Llenar_Liata_JR
        Me.lblEstado.Caption = "NORMAL"
        Me.lblEstado.ForeColor = vbBlue
    End If
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
Private Sub Image9_Click()
    Unload Me
End Sub
