VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCostoJR 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Costo de Juego de Reparación"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7920
      TabIndex        =   8
      Top             =   2640
      Width           =   975
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
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmCostoJR.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCostoJR.frx":030A
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7920
      TabIndex        =   6
      Top             =   5040
      Width           =   975
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCostoJR.frx":0899
         MousePointer    =   99  'Custom
         Picture         =   "FrmCostoJR.frx":0BA3
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7920
      TabIndex        =   4
      Top             =   3840
      Width           =   975
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmCostoJR.frx":2C85
         MousePointer    =   99  'Custom
         Picture         =   "FrmCostoJR.frx":2F8F
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmCostoJR.frx":4AD1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   600
         Width           =   5175
      End
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
         Left            =   6360
         Picture         =   "FrmCostoJR.frx":4AED
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4695
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   7455
         _ExtentX        =   13150
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
         Caption         =   "Producto :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   7920
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "FrmCostoJR"
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
    sBuscar = "SELECT * FROM VsCostoJR WHERE ID_REPARACION LIKE '%" & Text1.Text & "%' ORDER BY ID_REPARACION"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_REPARACION"))
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            sBuscar = "SELECT MAX(ID_ORDEN_COMPRA) AS ID, ID_PRODUCTO, DESCRIPCION, PRECIO From ORDEN_COMPRA_DETALLE WHERE (ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "') GROUP BY ID_PRODUCTO, DESCRIPCION, PRECIO ORDER BY ID_PRODUCTO, ID DESC"
            Set tRs1 = cnn.Execute(sBuscar)
            If Not (tRs1.EOF And tRs1.BOF) Then
                If Not IsNull(tRs1.Fields("PRECIO")) Then tLi.SubItems(3) = tRs1.Fields("PRECIO")
                sBuscar = "SELECT MONEDA FROM ORDEN_COMPRA WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID")
                Set tRs1 = cnn.Execute(sBuscar)
                If Not IsNull(tRs1.Fields("MONEDA")) Then tLi.SubItems(4) = tRs1.Fields("MONEDA")
            Else
                If Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(3) = tRs.Fields("PRECIO_COSTO")
                If Not IsNull(tRs.Fields("PRECIO_EN")) Then tLi.SubItems(4) = tRs.Fields("PRECIO_EN")
            End If
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
        .ColumnHeaders.Add , , "COMPONENTE", 4500
        .ColumnHeaders.Add , , "CANTIDAD", 1200
        .ColumnHeaders.Add , , "COSTO", 1200
        .ColumnHeaders.Add , , "MONEDA", 1000
    End With
End Sub
Private Sub Image10_Click()
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    If ListView1.ListItems.Count = 0 Then
        Command9.Value = True
    End If
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
    CommonDialog1.ShowSave
    Ruta = Me.CommonDialog1.FileName
    If ListView1.ListItems.Count > 0 Then
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
    Dim tRs2 As ADODB.Recordset
    Dim sBuscar As String
    Dim Total As Double
    Dim Dolar As Double
    Dim sIdRep As String
    'text1.text = no_orden
    'nvomen.Text1(0).Text = id_usuario
    If ListView1.ListItems.Count = 0 Then
        Command9.Value = True
    End If
    sBuscar = "SELECT COMPRA FROM DOLAR WHERE FECHA = '" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Dolar = CDbl(tRs.Fields("COMPRA"))
    End If
    If Not (ListView1.ListItems.Count = 0) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\RepCostoJR.pdf") Then
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
        oDoc.NewPage A4_Vertical
        oDoc.WImage 60, 20, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs1 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 10, 100, 560, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 10, 100, 560, tRs1.Fields("DIRECCION") & " Col." & tRs1.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 70, 10, 100, 560, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 80, 10, 100, 560, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 40, 400, 20, 150, "Fecha de Impresion: " & Date, "F3", 8, hCenter
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox 100, 10, 100, 560, "REPORTE DE COSTOS DE JUEGOS DE REPARACION", "F3", 8, hCenter
        Posi = 120
        oDoc.WTextBox Posi, 10, 20, 120, "Juego de Rep.", "F2", 8, hCenter
        oDoc.WTextBox Posi, 130, 20, 120, "Producto", "F2", 8, hCenter
        oDoc.WTextBox Posi, 250, 20, 60, "Cantidad", "F2", 8, hCenter
        oDoc.WTextBox Posi, 310, 20, 60, "Costo", "F2", 8, hCenter
        oDoc.WTextBox Posi, 370, 20, 60, "Moneda", "F2", 8, hCenter
        oDoc.WTextBox Posi, 430, 20, 60, "Total", "F2", 8, hCenter
        oDoc.WTextBox Posi, 490, 20, 50, "P. Venta", "F3", 8, hCenter
        oDoc.WTextBox Posi, 540, 20, 60, "Ganancia", "F3", 8, hCenter
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 760, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        sIdRep = ListView1.ListItems(1)
        For Cont = 1 To ListView1.ListItems.Count
            If sIdRep <> ListView1.ListItems(Cont) Then
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 430, Posi
                oDoc.WLineTo 490, Posi
                oDoc.LineStroke
                oDoc.WTextBox Posi, 430, 20, 60, Format(Total, "$#,###,##0.00"), "F3", 7, hRight
                sBuscar = "SELECT PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView1.ListItems(Cont) & "'"
                Set tRs2 = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    oDoc.WTextBox Posi, 490, 20, 50, Format(tRs2.Fields("PRECIO_COSTO") * ((tRs2.Fields("GANANCIA") / 100) + 1), "$#,###,##0.00"), "F3", 7, hRight
                    oDoc.WTextBox Posi, 540, 20, 50, Format(tRs2.Fields("PRECIO_COSTO") * ((tRs2.Fields("GANANCIA") / 100) + 1) - Total, "$#,###,##0.00"), "F3", 7, hRight
                End If
                Total = 0
                sIdRep = ListView1.ListItems(Cont)
                Posi = Posi + 24
            End If
            oDoc.WTextBox Posi, 10, 20, 120, ListView1.ListItems(Cont), "F3", 7, hLeft
            oDoc.WTextBox Posi, 130, 20, 120, ListView1.ListItems(Cont).SubItems(1), "F3", 7, hLeft
            oDoc.WTextBox Posi, 250, 20, 60, ListView1.ListItems(Cont).SubItems(2), "F3", 7, hCenter
            oDoc.WTextBox Posi, 310, 20, 60, Format(ListView1.ListItems(Cont).SubItems(3), "#,###,##0.00"), "F3", 7, hRight
            oDoc.WTextBox Posi, 370, 20, 60, ListView1.ListItems(Cont).SubItems(4), "F3", 7, hRight
            If ListView1.ListItems(Cont).SubItems(4) = "PESOS" Then
                oDoc.WTextBox Posi, 430, 20, 60, Format(ListView1.ListItems(Cont).SubItems(2) * ListView1.ListItems(Cont).SubItems(3), "$#,###,##0.00"), "F3", 7, hRight
                Total = Total + CDbl(ListView1.ListItems(Cont).SubItems(2) * ListView1.ListItems(Cont).SubItems(3))
            Else
                oDoc.WTextBox Posi, 430, 20, 60, Format(ListView1.ListItems(Cont).SubItems(2) * (ListView1.ListItems(Cont).SubItems(3) * Dolar), "$#,###,##0.00"), "F3", 7, hRight
                Total = Total + CDbl(ListView1.ListItems(Cont).SubItems(2) * (ListView1.ListItems(Cont).SubItems(3) * Dolar))
            End If
            Posi = Posi + 12
            If Posi >= 700 Then
                oDoc.NewPage A4_Vertical
                oDoc.WImage 60, 20, 43, 161, "Logo"
                sBuscar = "SELECT * FROM EMPRESA"
                Set tRs1 = cnn.Execute(sBuscar)
                oDoc.WTextBox 40, 10, 100, 560, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
                oDoc.WTextBox 60, 10, 100, 560, tRs1.Fields("DIRECCION") & " Col." & tRs1.Fields("COLONIA"), "F3", 8, hCenter
                oDoc.WTextBox 70, 10, 100, 560, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
                oDoc.WTextBox 80, 10, 100, 560, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
                oDoc.WTextBox 40, 400, 20, 150, "Fecha de Impresion: " & Date, "F3", 8, hCenter
                ' ENCABEZADO DEL DETALLE
                oDoc.WTextBox 100, 10, 100, 560, "REPORTE DE COSTOS DE JUEGOS DE REPARACION", "F3", 8, hCenter
                Posi = 120
                oDoc.WTextBox Posi, 10, 20, 120, "Juego de Rep.", "F2", 8, hCenter
                oDoc.WTextBox Posi, 130, 20, 120, "Producto", "F2", 8, hCenter
                oDoc.WTextBox Posi, 250, 20, 60, "Cantidad", "F2", 8, hCenter
                oDoc.WTextBox Posi, 310, 20, 60, "Costo", "F2", 8, hCenter
                oDoc.WTextBox Posi, 370, 20, 60, "Moneda", "F2", 8, hCenter
                oDoc.WTextBox Posi, 430, 20, 60, "Total", "F2", 8, hCenter
                oDoc.WTextBox Posi, 490, 20, 50, "P. Venta", "F3", 8, hCenter
                oDoc.WTextBox Posi, 540, 20, 60, "Ganancia", "F3", 8, hCenter
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
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 430, Posi
        oDoc.WLineTo 490, Posi
        oDoc.LineStroke
        oDoc.WTextBox Posi, 430, 20, 60, Format(Total, "$#,###,##0.00"), "F3", 7, hRight
        sBuscar = "SELECT PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView1.ListItems(Cont - 1) & "'"
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            oDoc.WTextBox Posi, 490, 20, 50, Format(tRs2.Fields("PRECIO_COSTO") * ((tRs2.Fields("GANANCIA") / 100) + 1), "$#,###,##0.00"), "F3", 7, hRight
            oDoc.WTextBox Posi, 540, 20, 50, Format(tRs2.Fields("PRECIO_COSTO") * ((tRs2.Fields("GANANCIA") / 100) + 1) - Total, "$#,###,##0.00"), "F3", 7, hRight
        End If
        Total = 0
        Posi = Posi + 24
        Posi = Posi + 30
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 760, Posi
        oDoc.LineStroke
        Posi = Posi + 16
        ' TEXTO ABAJO
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command9.Value = True
    End If
End Sub
