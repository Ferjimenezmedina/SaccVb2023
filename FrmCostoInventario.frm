VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCostoInventario 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Costos de Inventarios"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmCostoInventario.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Combo1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton Command9 
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
         Picture         =   "FrmCostoInventario.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Almacen 3"
         Height          =   255
         Left            =   6120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Almacen 2"
         Height          =   255
         Left            =   4920
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Almacen 1"
         Height          =   255
         Left            =   3720
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4695
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8281
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Sucursal :"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   10
      Top             =   3600
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmCostoInventario.frx":29EE
         MousePointer    =   99  'Custom
         Picture         =   "FrmCostoInventario.frx":2CF8
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   8
      Top             =   4800
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCostoInventario.frx":483A
         MousePointer    =   99  'Custom
         Picture         =   "FrmCostoInventario.frx":4B44
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label9 
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
   End
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   6
      Top             =   2400
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmCostoInventario.frx":6C26
         MousePointer    =   99  'Custom
         Picture         =   "FrmCostoInventario.frx":6F30
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   9000
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "FrmCostoInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command9_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    Dim Dolar As Double
    sBuscar = "SELECT COMPRA FROM DOLAR WHERE FECHA = '" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Dolar = CDbl(tRs.Fields("COMPRA"))
    End If
    If Option1.value Then
        If Combo1.Text = "<TODAS>" Then
            sBuscar = "SELECT ALMACEN1.ID_PRODUCTO, ALMACEN1.Descripcion, ALMACEN1.PRECIO_COSTO, EXISTENCIAS.SUCURSAL, EXISTENCIAS.CANTIDAD, ALMACEN1.PRECIO_EN FROM ALMACEN1 INNER JOIN EXISTENCIAS ON ALMACEN1.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO ORDER BY ID_PRODUCTO"
        Else
            sBuscar = "SELECT ALMACEN1.ID_PRODUCTO, ALMACEN1.Descripcion, ALMACEN1.PRECIO_COSTO, EXISTENCIAS.SUCURSAL, EXISTENCIAS.CANTIDAD, ALMACEN1.PRECIO_EN FROM ALMACEN1 INNER JOIN EXISTENCIAS ON ALMACEN1.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE EXISTENCIAS.SUCURSAL = '" & Combo1.Text & "' ORDER BY EXISTENCIAS.ID_PRODUCTO"
        End If
    Else
        If Option2.value Then
            If Combo1.Text = "<TODAS>" Then
                sBuscar = "SELECT ALMACEN2.ID_PRODUCTO, ALMACEN2.Descripcion, ALMACEN2.PRECIO_COSTO, EXISTENCIAS.SUCURSAL, EXISTENCIAS.CANTIDAD, ALMACEN2.PRECIO_EN FROM ALMACEN2 INNER JOIN EXISTENCIAS ON ALMACEN2.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO ORDER BY ID_PRODUCTO"
            Else
                sBuscar = "SELECT ALMACEN2.ID_PRODUCTO, ALMACEN2.Descripcion, ALMACEN2.PRECIO_COSTO, EXISTENCIAS.SUCURSAL, EXISTENCIAS.CANTIDAD, ALMACEN2.PRECIO_EN FROM ALMACEN2 INNER JOIN EXISTENCIAS ON ALMACEN2.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE EXISTENCIAS.SUCURSAL = '" & Combo1.Text & "' ORDER BY EXISTENCIAS.ID_PRODUCTO"
            End If
        Else
            If Combo1.Text = "<TODAS>" Then
                sBuscar = "SELECT ALMACEN3.ID_PRODUCTO, ALMACEN3.Descripcion, ALMACEN3.PRECIO_COSTO, EXISTENCIAS.SUCURSAL, EXISTENCIAS.CANTIDAD, ALMACEN3.PRECIO_EN FROM ALMACEN3 INNER JOIN EXISTENCIAS ON ALMACEN3.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE ALMACEN3.CLASIFICACION NOT IN ('SERVICIOS') ORDER BY EXISTENCIAS.ID_PRODUCTO"
            Else
                sBuscar = "SELECT ALMACEN3.ID_PRODUCTO, ALMACEN3.Descripcion, ALMACEN3.PRECIO_COSTO, EXISTENCIAS.SUCURSAL, EXISTENCIAS.CANTIDAD, ALMACEN3.PRECIO_EN FROM ALMACEN3 INNER JOIN EXISTENCIAS ON ALMACEN3.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE ALMACEN3.CLASIFICACION NOT IN ('SERVICIOS') AND EXISTENCIAS.SUCURSAL = '" & Combo1.Text & "' ORDER BY EXISTENCIAS.ID_PRODUCTO"
            End If
        End If
    End If
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            sBuscar = "SELECT MAX(ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA) AS ID, ORDEN_COMPRA_DETALLE.ID_PRODUCTO, ORDEN_COMPRA_DETALLE.DESCRIPCION, MAX(ORDEN_COMPRA_DETALLE.PRECIO) AS PRECIO, ORDEN_COMPRA.MONEDA FROM ORDEN_COMPRA_DETALLE INNER JOIN ORDEN_COMPRA ON ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA = ORDEN_COMPRA.ID_ORDEN_COMPRA WHERE (ORDEN_COMPRA_DETALLE.ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "') GROUP BY ORDEN_COMPRA_DETALLE.ID_PRODUCTO, ORDEN_COMPRA_DETALLE.DESCRIPCION, ORDEN_COMPRA.MONEDA ORDER BY ORDEN_COMPRA_DETALLE.ID_PRODUCTO, ID DESC"
            Set tRs1 = cnn.Execute(sBuscar)
            If Not (tRs1.EOF And tRs1.BOF) Then
                If Not IsNull(tRs1.Fields("PRECIO")) Then tLi.SubItems(3) = tRs1.Fields("PRECIO")
                If tRs1.Fields("MONEDA") = "PESOS" Then
                    If Not IsNull(tRs1.Fields("PRECIO")) Then tLi.SubItems(4) = tRs1.Fields("PRECIO")
                Else
                    If Not IsNull(tRs1.Fields("PRECIO")) Then tLi.SubItems(4) = Format(CDbl(tRs1.Fields("PRECIO")) * Dolar, "###,###,##0.00")
                End If
                sBuscar = "SELECT MONEDA FROM ORDEN_COMPRA WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID")
                Set tRs1 = cnn.Execute(sBuscar)
                If Not IsNull(tRs1.Fields("MONEDA")) Then tLi.SubItems(5) = tRs1.Fields("MONEDA")
            Else
                If Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(3) = tRs.Fields("PRECIO_COSTO")
                If tRs.Fields("PRECIO_EN") = "PESOS" Then
                    If Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(4) = tRs.Fields("PRECIO_COSTO")
                Else
                    If Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(4) = Format(CDbl(tRs.Fields("PRECIO_COSTO")) * Dolar, "###,###,##0.00")
                End If
                If Not IsNull(tRs.Fields("PRECIO_EN")) Then tLi.SubItems(5) = tRs.Fields("PRECIO_EN")
            End If
            If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(6) = tRs.Fields("SUCURSAL")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
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
        .ColumnHeaders.Add , , "ID PRODUCTO", 2000
        .ColumnHeaders.Add , , "DESCRIPCION", 4500
        .ColumnHeaders.Add , , "CANTIDAD", 1200
        .ColumnHeaders.Add , , "COSTO", 1200
        .ColumnHeaders.Add , , "COSTO EN PESOS", 1200
        .ColumnHeaders.Add , , "MONEDA", 1000
        .ColumnHeaders.Add , , "SUCURSAL", 1000
    End With
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Combo1.AddItem "<TODAS>"
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image10_Click()
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    If ListView1.ListItems.COUNT = 0 Then
        Command9.value = True
    End If
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
    CommonDialog1.ShowSave
    Ruta = Me.CommonDialog1.FileName
    If ListView1.ListItems.COUNT > 0 Then
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.COUNT
            For Con = 1 To ListView1.ColumnHeaders.COUNT
                StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView1.ListItems.COUNT
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
Private Sub Image26_Click()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim sBuscar As String
    Dim Total As Double
    Dim Dolar As Double
    'text1.text = no_orden
    'nvomen.Text1(0).Text = id_usuario
    If ListView1.ListItems.COUNT = 0 Then
        Command9.value = True
    End If
    sBuscar = "SELECT COMPRA FROM DOLAR WHERE FECHA = '" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Dolar = CDbl(tRs.Fields("COMPRA"))
    End If
    If Not (ListView1.ListItems.COUNT = 0) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\RepCostoInventario.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        ' asi se agregan los logos... solo te falto poner un control IMAGE1 para cargar la imagen en el
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Horizontal
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs1 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 10, 100, 760, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 10, 100, 760, tRs1.Fields("DIRECCION") & " Col." & tRs1.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 70, 10, 100, 760, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 80, 10, 100, 760, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 40, 500, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox 100, 10, 100, 760, "REPORTE DE COSTOS DE INVENTARIOS", "F3", 8, hCenter
        Posi = 120
        oDoc.WTextBox Posi, 10, 20, 55, "Producto", "F2", 8, hCenter
        oDoc.WTextBox Posi, 60, 20, 450, "DESCRIPCION", "F2", 8, hCenter
        oDoc.WTextBox Posi, 510, 20, 50, "Cantidad", "F2", 8, hCenter
        oDoc.WTextBox Posi, 560, 20, 50, "Costo", "F2", 8, hCenter
        oDoc.WTextBox Posi, 610, 20, 50, "Moneda", "F2", 8, hCenter
        oDoc.WTextBox Posi, 660, 20, 50, "Sucursal", "F2", 8, hCenter
        oDoc.WTextBox Posi, 710, 20, 50, "Total", "F2", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 760, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        For Cont = 1 To ListView1.ListItems.COUNT
            oDoc.WTextBox Posi, 10, 20, 55, ListView1.ListItems(Cont), "F3", 7, hLeft
            oDoc.WTextBox Posi, 60, 20, 450, ListView1.ListItems(Cont).SubItems(1), "F3", 7, hLeft
            oDoc.WTextBox Posi, 510, 20, 50, ListView1.ListItems(Cont).SubItems(2), "F3", 7, hCenter
            oDoc.WTextBox Posi, 560, 20, 50, Format(ListView1.ListItems(Cont).SubItems(3), "#,###,##0.00"), "F3", 7, hRight
            oDoc.WTextBox Posi, 610, 20, 50, ListView1.ListItems(Cont).SubItems(5), "F3", 7, hRight
            oDoc.WTextBox Posi, 660, 20, 50, ListView1.ListItems(Cont).SubItems(6), "F3", 7, hRight
            If ListView1.ListItems(Cont).SubItems(5) = "PESOS" Then
                oDoc.WTextBox Posi, 710, 20, 50, Format(ListView1.ListItems(Cont).SubItems(2) * ListView1.ListItems(Cont).SubItems(3), "#,###,##0.00"), "F3", 7, hRight
                Total = Total + CDbl(ListView1.ListItems(Cont).SubItems(2) * ListView1.ListItems(Cont).SubItems(3))
            Else
                oDoc.WTextBox Posi, 710, 20, 50, Format(ListView1.ListItems(Cont).SubItems(2) * (ListView1.ListItems(Cont).SubItems(3) * Dolar), "#,###,##0.00"), "F3", 7, hRight
                Total = Total + CDbl(ListView1.ListItems(Cont).SubItems(2) * (ListView1.ListItems(Cont).SubItems(3) * Dolar))
            End If
            Posi = Posi + 12
            If Posi >= 520 Then
                oDoc.NewPage A4_Horizontal
                oDoc.WImage 70, 40, 43, 161, "Logo"
                sBuscar = "SELECT * FROM EMPRESA"
                Set tRs1 = cnn.Execute(sBuscar)
                oDoc.WTextBox 40, 10, 100, 760, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
                oDoc.WTextBox 60, 10, 100, 760, tRs1.Fields("DIRECCION") & " Col." & tRs1.Fields("COLONIA"), "F3", 8, hCenter
                oDoc.WTextBox 70, 10, 100, 760, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
                oDoc.WTextBox 80, 10, 100, 760, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
                oDoc.WTextBox 40, 500, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
                ' ENCABEZADO DEL DETALLE
                oDoc.WTextBox 100, 10, 100, 760, "REPORTE DE ESTADO DE ORDENES DE COMPRA", "F3", 8, hCenter
                Posi = 120
                oDoc.WTextBox Posi, 10, 20, 55, "Producto", "F2", 8, hCenter
                oDoc.WTextBox Posi, 60, 20, 450, "DESCRIPCION", "F2", 8, hCenter
                oDoc.WTextBox Posi, 510, 20, 50, "Cantidad", "F2", 8, hCenter
                oDoc.WTextBox Posi, 560, 20, 50, "Costo", "F2", 8, hCenter
                oDoc.WTextBox Posi, 610, 20, 50, "Moneda", "F2", 8, hCenter
                oDoc.WTextBox Posi, 660, 20, 50, "Sucursal", "F2", 8, hCenter
                oDoc.WTextBox Posi, 710, 20, 50, "Total", "F2", 8, hCenter
                Posi = Posi + 12
                ' Linea
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, Posi
                oDoc.WLineTo 760, Posi
                oDoc.LineStroke
                Posi = Posi + 6
            End If
        Next
        ' Linea
        Posi = Posi + 30
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 760, Posi
        oDoc.LineStroke
        Posi = Posi + 16
        oDoc.WTextBox Posi, 560, 20, 200, "TOTAL EN INVENTARIO : " & Format(Total, "###,###,##0.00"), "F3", 10, hRight
        ' TEXTO ABAJO
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
