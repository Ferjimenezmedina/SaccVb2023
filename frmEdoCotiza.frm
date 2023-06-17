VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEdoCotiza 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado de Cotizaciones"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmEdoCotiza.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lvAprovadas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lvCotizadas"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lvCanceladas"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin MSComctlLib.ListView lvCanceladas 
         Height          =   1575
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2778
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvCotizadas 
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   2760
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2778
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvAprovadas 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   4800
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2778
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cotizaciones Rechazadas"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cotizaciones Pendientes de Aprobación"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cotizaciones Aprobadas Pendientes de Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   4440
         Width           =   3255
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5760
      TabIndex        =   3
      Top             =   5400
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmEdoCotiza.frx":001C
         MousePointer    =   99  'Custom
         Picture         =   "frmEdoCotiza.frx":0326
         Top             =   120
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmEdoCotiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With lvCotizadas
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID REQUISICION", 0
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripcion", 3500
        .ColumnHeaders.Add , , "CANTIDAD", 1000, 2
        .ColumnHeaders.Add , , "FECHA", 0, 2
        .ColumnHeaders.Add , , "CONTADOR", 0, 2
    End With
    With lvAprovadas
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID COTIZACION", 0
        .ColumnHeaders.Add , , "ID REQUISICION", 0
        .ColumnHeaders.Add , , "ID PROVEEDOR", 0
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripcion", 3500
        .ColumnHeaders.Add , , "CANTIDAD", 1000, 2
        .ColumnHeaders.Add , , "DIAS ENTREGA", 1440, 2
        .ColumnHeaders.Add , , "PRECIO", 1440, 2
        .ColumnHeaders.Add , , "FECHA", 0, 2
        .ColumnHeaders.Add , , "IDS", 0
        .ColumnHeaders.Add , , "NUMOC", 100
        .ColumnHeaders.Add , , "MONEDA", 100
    End With
    With lvCanceladas
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID COTIZACION", 0
        .ColumnHeaders.Add , , "ID REQUISICION", 0
        .ColumnHeaders.Add , , "ID PROVEEDOR", 0
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripcion", 3500
        .ColumnHeaders.Add , , "CANTIDAD", 1000, 2
        .ColumnHeaders.Add , , "DIAS ENTREGA", 1440, 2
        .ColumnHeaders.Add , , "PRECIO", 1440, 2
        .ColumnHeaders.Add , , "FECHA", 0, 2
        .ColumnHeaders.Add , , "IDS", 0
        .ColumnHeaders.Add , , "NUMOC", 100
        .ColumnHeaders.Add , , "MONEDA", 100
    End With
    Llenar_Lista_Requisiciones
    Llenar_Lista_Canceladas
    Llenar_Lista_Cotizaciones
End Sub
Sub Llenar_Lista_Requisiciones()
On Error GoTo ManejaError
    Dim sqlQuery As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sqlQuery = "SELECT ID_REQUISICION, ID_PRODUCTO, DESCRIPCION, CANTIDAD, FECHA, CONTADOR FROM REQUISICION WHERE ACTIVO = 0 AND COTIZADA = 1"
    Set tRs = cnn.Execute(sqlQuery)
    lvCotizadas.ListItems.Clear
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvCotizadas.ListItems.Add(, , .Fields("ID_REQUISICION"))
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(2) = Trim(.Fields("Descripcion"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(3) = Trim(.Fields("CANTIDAD"))
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(4) = Trim(.Fields("FECHA"))
                If Not IsNull(.Fields("CONTADOR")) Then
                    tLi.SubItems(5) = Trim(.Fields("CONTADOR"))
                Else
                    tLi.SubItems(5) = "0"
                End If
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Cotizaciones()
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim CONT2 As Integer
    Dim tLi As ListItem
    sqlQuery = "SELECT NUMOC, ID_COTIZACION, ID_REQUISICION, ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, CANTIDAD, DIAS_ENTREGA, PRECIO, FECHA, ISNULL(MONEDA, '') AS MONEDA FROM COTIZA_REQUI WHERE ESTADO_ACTUAL = 'X'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        lvAprovadas.ListItems.Clear
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvAprovadas.ListItems.Add(, , .Fields("ID_COTIZACION"))
                If Not IsNull(.Fields("ID_REQUISICION")) Then tLi.SubItems(1) = Trim(.Fields("ID_REQUISICION"))
                If Not IsNull(.Fields("ID_PROVEEDOR")) Then tLi.SubItems(2) = Trim(.Fields("ID_PROVEEDOR"))
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(3) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(4) = Trim(.Fields("Descripcion"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(5) = Trim(.Fields("CANTIDAD"))
                If Not IsNull(.Fields("DIAS_ENTREGA")) Then tLi.SubItems(6) = Trim(.Fields("DIAS_ENTREGA"))
                If Not IsNull(.Fields("PRECIO")) Then tLi.SubItems(7) = Trim(.Fields("PRECIO"))
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(8) = Trim(.Fields("FECHA"))
                If Not IsNull(.Fields("NUMOC")) Then
                    If .Fields("NUMOC") <> "0" Then
                        tLi.SubItems(10) = Trim(.Fields("NUMOC"))
                    End If
                End If
                If Not IsNull(.Fields("MONEDA")) Then tLi.SubItems(11) = Trim(.Fields("MONEDA"))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Canceladas()
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim CONT2 As Integer
    Dim tLi As ListItem
    sqlQuery = "SELECT NUMOC, ID_COTIZACION, ID_REQUISICION, ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, CANTIDAD, DIAS_ENTREGA, PRECIO, FECHA, ISNULL(MONEDA, '') AS MONEDA FROM COTIZA_REQUI WHERE ESTADO_ACTUAL = 'Z'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        lvCanceladas.ListItems.Clear
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvCanceladas.ListItems.Add(, , .Fields("ID_COTIZACION"))
                If Not IsNull(.Fields("ID_REQUISICION")) Then tLi.SubItems(1) = Trim(.Fields("ID_REQUISICION"))
                If Not IsNull(.Fields("ID_PROVEEDOR")) Then tLi.SubItems(2) = Trim(.Fields("ID_PROVEEDOR"))
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(3) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(4) = Trim(.Fields("Descripcion"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(5) = Trim(.Fields("CANTIDAD"))
                If Not IsNull(.Fields("DIAS_ENTREGA")) Then tLi.SubItems(6) = Trim(.Fields("DIAS_ENTREGA"))
                If Not IsNull(.Fields("PRECIO")) Then tLi.SubItems(7) = Trim(.Fields("PRECIO"))
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(8) = Trim(.Fields("FECHA"))
                If Not IsNull(.Fields("NUMOC")) Then
                    If .Fields("NUMOC") <> "0" Then
                        tLi.SubItems(10) = Trim(.Fields("NUMOC"))
                    End If
                End If
                If Not IsNull(.Fields("MONEDA")) Then tLi.SubItems(11) = Trim(.Fields("MONEDA"))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
