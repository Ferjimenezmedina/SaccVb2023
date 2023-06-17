VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRepGastos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reporte de Gastos"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7575
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   9960
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9960
      TabIndex        =   23
      Top             =   5040
      Width           =   975
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
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRepGastos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepGastos.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9960
      TabIndex        =   1
      Top             =   6240
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepGastos.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepGastos.frx":2156
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label26 
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   12938
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Ordenes Rapidas"
      TabPicture(0)   =   "FrmRepGastos.frx":4238
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CommonDialog1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Salidas de Almacen"
      TabPicture(1)   =   "FrmRepGastos.frx":4254
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ListView2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Productos Comprados"
      TabPicture(2)   =   "FrmRepGastos.frx":4270
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "ListView3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.CommandButton Command3 
         Caption         =   "Busca"
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
         Left            =   -66720
         Picture         =   "FrmRepGastos.frx":428C
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Caption         =   "Rango de fechas"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   3615
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   375
            Left            =   240
            TabIndex        =   29
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50397185
            CurrentDate     =   40208
         End
         Begin MSComCtl2.DTPicker DTPicker6 
            Height          =   375
            Left            =   1920
            TabIndex        =   30
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50397185
            CurrentDate     =   40208
         End
         Begin VB.Label Label6 
            Caption         =   "Al"
            Height          =   255
            Left            =   1920
            TabIndex        =   32
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Del"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   1335
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8280
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         Caption         =   "Estado"
         Height          =   1335
         Left            =   3840
         TabIndex        =   19
         Top             =   480
         Width           =   1695
         Begin VB.OptionButton Option3 
            Caption         =   "Todas"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   840
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Pendientes"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Pagadas"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Rango de fechas"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   3615
         Begin VB.CheckBox Check2 
            Caption         =   "Buscar por rango de fechas"
            Height          =   195
            Left            =   600
            TabIndex        =   14
            Top             =   1080
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   1920
            TabIndex        =   15
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   50397185
            CurrentDate     =   40208
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   50397185
            CurrentDate     =   40208
         End
         Begin VB.Label Label4 
            Caption         =   "Del"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Al"
            Height          =   255
            Left            =   1920
            TabIndex        =   17
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Busca"
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
         Left            =   -66720
         Picture         =   "FrmRepGastos.frx":6C5E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Busca"
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
         Left            =   8280
         Picture         =   "FrmRepGastos.frx":9630
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rango de fechas"
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3615
         Begin VB.OptionButton Option5 
            Caption         =   "Fecha De Creación"
            Height          =   255
            Left            =   1680
            TabIndex        =   27
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Fecha De Pago"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Buscar por rango de fechas"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   2160
            TabIndex        =   7
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   50397185
            CurrentDate     =   40208
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   480
            TabIndex        =   6
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   50397185
            CurrentDate     =   40208
         End
         Begin VB.Label Label2 
            Caption         =   "Al"
            Height          =   255
            Left            =   1920
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Del"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   4
         Top             =   1920
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   9340
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5295
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   9340
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   34
         Top             =   1920
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   9340
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "FrmRepGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    Else
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    End If
End Sub
Private Sub Check2_Click()
    If Check2.Value = 1 Then
        DTPicker3.Enabled = True
        DTPicker4.Enabled = True
    Else
        DTPicker3.Enabled = False
        DTPicker4.Enabled = False
    End If
End Sub
Private Sub Command1_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If Option4.Value Then
        If Check1.Value = 1 Then
            If Option1.Value Then
                sBuscar = "SELECT * FROM VSPAGOS_OR WHERE FECHA_PAGO BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND ESTADO = 'F'"
            Else
                If Option2.Value Then
                    sBuscar = "SELECT * FROM VSPAGOS_OR WHERE FECHA_PAGO BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND ESTADO = 'A'"
                Else
                    sBuscar = "SELECT * FROM VSPAGOS_OR WHERE FECHA_PAGO BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
                End If
            End If
        Else
            If Option1.Value Then
                sBuscar = "SELECT * FROM VsRepGastos WHERE ESTADO = 'F'"
            Else
                If Option2.Value Then
                    sBuscar = "SELECT * FROM VsRepGastos WHERE ESTADO = 'A'"
                Else
                    sBuscar = "SELECT * FROM VsRepGastos"
                End If
            End If
        End If
    Else
        If Check1.Value = 1 Then
            If Option1.Value Then
                sBuscar = "SELECT * FROM VsRepGastos WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND ESTADO = 'F'"
            Else
                If Option2.Value Then
                    sBuscar = "SELECT * FROM VsRepGastos WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' AND ESTADO = 'A'"
                Else
                    sBuscar = "SELECT * FROM VsRepGastos WHERE FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
                End If
            End If
        Else
            If Option1.Value Then
                sBuscar = "SELECT * FROM VsRepGastos WHERE ESTADO = 'F'"
            Else
                If Option2.Value Then
                    sBuscar = "SELECT * FROM VsRepGastos WHERE ESTADO = 'A'"
                Else
                    sBuscar = "SELECT * FROM VsRepGastos"
                End If
            End If
        End If
    End If
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_ORDEN_RAPIDA"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(2) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("COMENTARIO")) Then tLi.SubItems(3) = tRs.Fields("COMENTARIO")
            If Not IsNull(tRs.Fields("MONEDA")) Then tLi.SubItems(4) = tRs.Fields("MONEDA")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(5) = Format(tRs.Fields("TOTAL"), "###,###,##0.00")
            If Not IsNull(tRs.Fields("ESTADO")) Then tLi.SubItems(6) = tRs.Fields("ESTADO")
            If Not IsNull(tRs.Fields("FECHA_PAGO")) Then tLi.SubItems(7) = tRs.Fields("FECHA_PAGO")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command2_Click()
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim tLi As ListItem
    If Check2.Value = 1 Then
        sBuscar = "SELECT * FROM VsRepSalidas WHERE FECHA BETWEEN '" & DTPicker4.Value & "' AND '" & DTPicker3.Value & "'"
    Else
        sBuscar = "SELECT * FROM VsRepSalidas"
    End If
    Set tRs = cnn.Execute(sBuscar)
    ListView2.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_SALIDA"))
            If Not IsNull(tRs.Fields("JUSTIFICACION")) Then tLi.SubItems(1) = tRs.Fields("JUSTIFICACION")
            If Not IsNull(tRs.Fields("PRECIO")) Then tLi.SubItems(2) = Format(tRs.Fields("PRECIO"), "###,###,##0.00")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(3) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(4) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(5) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(6) = tRs.Fields("SUCURSAL")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command3_Click()
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim tLi As ListItem
    sBuscar = "SELECT ORDEN_RAPIDA_DETALLE.ID_PRODUCTO, PRODUCTOS_CONSUMIBLES.DESCRIPCION, SUM(ORDEN_RAPIDA_DETALLE.CANTIDAD) AS CANTIDAD, SUM(ORDEN_RAPIDA_DETALLE.SUBTOTAL) AS SUBTOTAL, SUM(ORDEN_RAPIDA_DETALLE.IVA) AS IVA, SUM(ORDEN_RAPIDA_DETALLE.IVARETENIDO) AS IVA_RETENIDO, SUM(ORDEN_RAPIDA_DETALLE.ISR) AS ISR, SUM(ORDEN_RAPIDA_DETALLE.IVADIEZ) AS IVA_10, SUM(ORDEN_RAPIDA_DETALLE.ISR2) AS ISR_2, SUM(ORDEN_RAPIDA_DETALLE.DESCUENTO) AS DESCUENTO, SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOTAL FROM ORDEN_RAPIDA_DETALLE INNER JOIN PRODUCTOS_CONSUMIBLES ON ORDEN_RAPIDA_DETALLE.ID_PRODUCTO = PRODUCTOS_CONSUMIBLES.ID_PRODUCTO INNER JOIN ORDEN_RAPIDA ON ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA = ORDEN_RAPIDA.ID_ORDEN_RAPIDA WHERE (ORDEN_RAPIDA.FECHA BETWEEN '" & DTPicker5.Value & "' AND '" & DTPicker6.Value & "') AND (ORDEN_RAPIDA.ESTADO = 'F') GROUP BY ORDEN_RAPIDA_DETALLE.ID_PRODUCTO, PRODUCTOS_CONSUMIBLES.DESCRIPCION"
    Set tRs = cnn.Execute(sBuscar)
    ListView3.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("DESCRIPCION")) Then tLi.SubItems(1) = tRs.Fields("DESCRIPCION")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = Format(tRs.Fields("CANTIDAD"), "###,###,##0.00")
            If Not IsNull(tRs.Fields("SUBTOTAL")) Then tLi.SubItems(3) = Format(tRs.Fields("SUBTOTAL"), "###,###,##0.00")
            If Not IsNull(tRs.Fields("IVA")) Then tLi.SubItems(4) = Format(tRs.Fields("IVA"), "###,###,##0.00")
            If Not IsNull(tRs.Fields("IVA_RETENIDO")) Then tLi.SubItems(5) = Format(tRs.Fields("IVA_RETENIDO"), "###,###,##0.00")
            If Not IsNull(tRs.Fields("ISR")) Then tLi.SubItems(6) = Format(tRs.Fields("ISR"), "###,###,##0.00")
            If Not IsNull(tRs.Fields("IVA_10")) Then tLi.SubItems(6) = Format(tRs.Fields("IVA_10"), "###,###,##0.00")
            If Not IsNull(tRs.Fields("ISR_2")) Then tLi.SubItems(6) = Format(tRs.Fields("ISR_2"), "###,###,##0.00")
            If Not IsNull(tRs.Fields("DESCUENTO")) Then tLi.SubItems(6) = Format(tRs.Fields("DESCUENTO"), "###,###,##0.00")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(6) = Format(tRs.Fields("TOTAL"), "###,###,##0.00")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Load()
On Error GoTo ManejaError
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Value = Date - 30
    DTPicker2.Value = Date
    DTPicker3.Value = Date
    DTPicker4.Value = Date - 30
    DTPicker5.Value = Date
    DTPicker6.Value = Date - 30
    Set cnn = New ADODB.Connection
    Dim sBuscar As String
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Num. Orden", 1500
        .ColumnHeaders.Add , , "Proveedor", 5500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Comentarios", 1400
        .ColumnHeaders.Add , , "Moneda", 1000
        .ColumnHeaders.Add , , "Total", 1200, hRight
        .ColumnHeaders.Add , , "Estado", 1000
        .ColumnHeaders.Add , , "Fecha de Pago", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Salida", 1500
        .ColumnHeaders.Add , , "Motivo", 5500
        .ColumnHeaders.Add , , "Precio", 1500, hRight
        .ColumnHeaders.Add , , "Cantidad", 1400
        .ColumnHeaders.Add , , "Usuario", 1000
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Sucursal", 1000
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID PRODUCTO", 1500
        .ColumnHeaders.Add , , "DESCRPCION", 5500
        .ColumnHeaders.Add , , "CANTIDAD", 1500, hRight
        .ColumnHeaders.Add , , "SUBTOTAL", 1400
        .ColumnHeaders.Add , , "IVA", 1000
        .ColumnHeaders.Add , , "IVA RETENIDO", 1200
        .ColumnHeaders.Add , , "ISR", 1000
        .ColumnHeaders.Add , , "IVA 10", 1000
        .ColumnHeaders.Add , , "ISR 2", 1000
        .ColumnHeaders.Add , , "DESCUENTO", 1000
        .ColumnHeaders.Add , , "TOTAL", 1000
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image10_Click()
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Dim foo As Integer
    Me.CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
    Me.CommonDialog1.ShowSave
    Ruta = Me.CommonDialog1.FileName
    If SSTab1.Tab = 0 And ListView1.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.Count
            For Con = 1 To ListView1.ColumnHeaders.Count
                StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView1.ListItems.Count
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            'archivo TXT
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
    If SSTab1.Tab = 1 And ListView2.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = ListView2.ColumnHeaders.Count
            For Con = 1 To ListView2.ColumnHeaders.Count
                StrCopi = StrCopi & ListView2.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView2.ListItems.Count
                StrCopi = StrCopi & ListView2.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView2.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            'archivo TXT
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
    If SSTab1.Tab = 2 And ListView3.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = ListView3.ColumnHeaders.Count
            For Con = 1 To ListView3.ColumnHeaders.Count
                StrCopi = StrCopi & ListView3.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView3.ListItems.Count
                StrCopi = StrCopi & ListView3.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView3.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            'archivo TXT
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub

