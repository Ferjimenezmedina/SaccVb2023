VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepPrestamoActivo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte prestamos uso activo fijo"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRepPrestamoActivo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command12"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   4575
         Begin VB.OptionButton Option2 
            Caption         =   "Vence"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3600
            TabIndex        =   7
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Inicio"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3600
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   480
            TabIndex        =   4
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   50724865
            CurrentDate     =   41187
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Por fecha"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   2160
            TabIndex        =   5
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   50724865
            CurrentDate     =   41187
         End
         Begin VB.Label Label3 
            Caption         =   "Al"
            Height          =   255
            Left            =   1920
            TabIndex        =   16
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Del"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1335
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2355
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
      Begin VB.CommandButton Command12 
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
         Left            =   3720
         Picture         =   "FrmRepPrestamoActivo.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Eliminara todos los articulos marcados con el recuadro"
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5040
      TabIndex        =   10
      Top             =   840
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmRepPrestamoActivo.frx":29EE
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepPrestamoActivo.frx":2CF8
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reporte"
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5040
      TabIndex        =   8
      Top             =   2040
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
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepPrestamoActivo.frx":3287
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepPrestamoActivo.frx":3591
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   5040
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FrmRepPrestamoActivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim IdCliente As String
Private Sub Check1_Click()
    If Check1.Value = 0 Then
        Option1.Enabled = False
        Option2.Enabled = False
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    Else
        Option1.Enabled = True
        Option2.Enabled = True
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    End If
End Sub
Private Sub Command12_Click()
    BuscaCliente
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Value = Date - 30
    DTPicker2.Value = Date
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    Set cnn = New ADODB.Connection
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
        .ColumnHeaders.Add , , "Id Cliente", 0
        .ColumnHeaders.Add , , "Nombre", 4500
    End With
End Sub
Private Sub BuscaCliente()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE ID_CLIENTE IN (SELECT ID_CLIENTE FROM VSPRESTAMO WHERE ESTADO = 'P') AND NOMBRE LIKE '%" & Text1.Text & "%'"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image26_Click()
    'IdCliente
    'SELECT ID_PRESTAMO, NOMBRE, FECHA_PRESTAMO, FECHA_ENTREGA, DEPOSITO, NOTAS FROM VSPRESTAMO WHERE ESTADO = 'P' AND NOTAS LIKE '%" & Text11.Text & "%' ORDER BY ID_PRESTAMO
    'SELECT CANTIDAD, ID_PRODUCTO, PRECIO_VENTA, DESCRIPCION, ID FROM VSPRESTAMO_DETALLE WHERE ID_PRESTAMO = " & NoFolioElim
    Dim oDoc  As cPDF
    Dim Posi As Integer
    Dim sBuscar As String
    Dim tRs8 As ADODB.Recordset
    Dim PosVer As Integer
    Dim Posabo As Integer
    Dim tRs  As ADODB.Recordset
    Posi = 120
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate(App.Path & "\PrestamosActivoFijo.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
' Encabezado del reporte
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    oDoc.LoadImage Image1, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 70, 40, 43, 161, "Logo"
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
    oDoc.WTextBox 90, 200, 20, 250, "Prestamos de Activo Fijo", "F2", 10, hCenter
    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
    oDoc.MoveTo 10, 110
    oDoc.WLineTo 580, 110
    oDoc.LineStroke
' Cuerpo del reporte
    If IdCliente <> "" Then
        If Check1.Value = 0 Then
            sBuscar = "SELECT ID_PRESTAMO, NOMBRE, FECHA_PRESTAMO, FECHA_ENTREGA, DEPOSITO, NOTAS FROM VSPRESTAMO WHERE ESTADO = 'P' AND ID_CLIENTE = " & IdCliente & " ORDER BY ID_PRESTAMO"
        Else
            If Option1.Value Then
                sBuscar = "SELECT ID_PRESTAMO, NOMBRE, FECHA_PRESTAMO, FECHA_ENTREGA, DEPOSITO, NOTAS FROM VSPRESTAMO WHERE ESTADO = 'P' AND ID_CLIENTE = " & IdCliente & " AND FECHA_PRESTAMO BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY ID_PRESTAMO"
            Else
                sBuscar = "SELECT ID_PRESTAMO, NOMBRE, FECHA_PRESTAMO, FECHA_ENTREGA, DEPOSITO, NOTAS FROM VSPRESTAMO WHERE ESTADO = 'P' AND ID_CLIENTE = " & IdCliente & " AND FECHA_ENTREGA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY ID_PRESTAMO"
            End If
        End If
    Else
        If Check1.Value = 0 Then
            sBuscar = "SELECT ID_PRESTAMO, NOMBRE, FECHA_PRESTAMO, FECHA_ENTREGA, DEPOSITO, NOTAS FROM VSPRESTAMO WHERE ESTADO = 'P' ORDER BY ID_PRESTAMO"
        Else
            If Option1.Value Then
                sBuscar = "SELECT ID_PRESTAMO, NOMBRE, FECHA_PRESTAMO, FECHA_ENTREGA, DEPOSITO, NOTAS FROM VSPRESTAMO WHERE ESTADO = 'P' AND FECHA_PRESTAMO BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY ID_PRESTAMO"
            Else
                sBuscar = "SELECT ID_PRESTAMO, NOMBRE, FECHA_PRESTAMO, FECHA_ENTREGA, DEPOSITO, NOTAS FROM VSPRESTAMO WHERE ESTADO = 'P' AND FECHA_ENTREGA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY ID_PRESTAMO"
            End If
        End If
    End If
    Set tRs8 = cnn.Execute(sBuscar)
    If Not (tRs8.EOF And tRs8.BOF) Then
        Do While Not tRs8.EOF
            oDoc.WTextBox Posi, 20, 30, 330, tRs8.Fields("NOMBRE"), "F2", 10, hLeft
            oDoc.WTextBox Posi, 350, 50, 100, "Del: " & tRs8.Fields("FECHA_PRESTAMO"), "F1", 10, hLeft
            oDoc.WTextBox Posi, 450, 40, 100, "Al: " & tRs8.Fields("FECHA_ENTREGA"), "F1", 10, hLeft
            Posi = Posi + 10
            oDoc.WTextBox Posi, 20, 30, 520, "Notas : " & tRs8.Fields("NOTAS"), "F1", 10, hLeft
            Posi = Posi + 10
            oDoc.WTextBox Posi, 20, 40, 50, "Cantidad", "F2", 8, hLeft
            oDoc.WTextBox Posi, 70, 40, 100, "Producto", "F2", 8, hLeft
            oDoc.WTextBox Posi, 140, 40, 350, "DESCRIPCION", "F2", 8, hLeft
            oDoc.WTextBox Posi, 490, 40, 50, "Precio", "F2", 8, hCenter
            sBuscar = "SELECT CANTIDAD, ID_PRODUCTO, PRECIO_VENTA, DESCRIPCION, ID FROM VSPRESTAMO_DETALLE WHERE ID_PRESTAMO = " & tRs8.Fields("ID_PRESTAMO")
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                 Do While Not (tRs.EOF)
                    Posi = Posi + 10
                    oDoc.WTextBox Posi, 20, 40, 50, tRs.Fields("CANTIDAD"), "F1", 9, hCenter
                    oDoc.WTextBox Posi, 70, 40, 100, tRs.Fields("ID_PRODUCTO"), "F1", 9, hLeft
                    oDoc.WTextBox Posi, 140, 40, 350, tRs.Fields("Descripcion"), "F1", 9, hLeft
                    oDoc.WTextBox Posi, 490, 40, 50, Format(tRs.Fields("PRECIO_VENTA"), "###,###,##0.00"), "F1", 9, hRight
                    tRs.MoveNext
                    If Posi >= 730 Then
                        Posi = 120
                        oDoc.NewPage A4_Vertical
                        oDoc.WImage 70, 40, 43, 161, "Logo"
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                        oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                        oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                        oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                        oDoc.WTextBox 90, 200, 20, 250, "Prestamos de Activo Fijo", "F2", 10, hCenter
                        oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
                        oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
                        oDoc.MoveTo 10, 110
                        oDoc.WLineTo 580, 110
                        oDoc.LineStroke
                    End If
                Loop
                Posi = Posi + 30
            End If
            Posi = Posi + 10
            oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, Posi
            oDoc.WLineTo 580, Posi
            oDoc.LineStroke
            tRs8.MoveNext
        Loop
    End If
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item.SubItems(1)
    IdCliente = Item
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BuscaCliente
    End If
End Sub
