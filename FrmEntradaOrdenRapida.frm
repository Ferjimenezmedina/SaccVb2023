VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmEntradaOrdenRapida 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada de orden rapida"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reimprime"
      Height          =   1335
      Left            =   8760
      TabIndex        =   22
      Top             =   360
      Width           =   975
      Begin VB.CommandButton Command2 
         Caption         =   "Imp."
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
         Left            =   120
         Picture         =   "FrmEntradaOrdenRapida.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   840
         Width           =   690
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   8760
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   8760
      TabIndex        =   13
      Top             =   2010
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   2
      Top             =   6120
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmEntradaOrdenRapida.frx":29D2
         MousePointer    =   99  'Custom
         Picture         =   "FrmEntradaOrdenRapida.frx":2CDC
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmEntradaOrdenRapida.frx":4DBE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.CommandButton Command1 
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
         Left            =   7320
         Picture         =   "FrmEntradaOrdenRapida.frx":4DDA
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   6720
         Width           =   1050
      End
      Begin VB.Frame Frame1 
         Caption         =   "Recibido"
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   5160
         Width           =   8295
         Begin VB.OptionButton Option1 
            Caption         =   "No. Serie"
            Height          =   255
            Left            =   4680
            TabIndex        =   20
            Top             =   960
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   3600
            TabIndex        =   16
            Top             =   600
            Width           =   4455
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   1080
            TabIndex        =   14
            Top             =   600
            Width           =   1455
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
            Height          =   375
            Left            =   7080
            Picture         =   "FrmEntradaOrdenRapida.frx":77AC
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   960
            Width           =   1050
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   7080
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4440
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   7
            Top             =   240
            Width           =   2535
         End
         Begin VB.OptionButton Option2 
            Caption         =   "No Aplica"
            Height          =   255
            Left            =   5760
            TabIndex        =   21
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Comentario"
            Height          =   255
            Left            =   2640
            TabIndex        =   17
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Factura"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Cantidad Recibida"
            Height          =   255
            Left            =   5640
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   3720
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Producto"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4048
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
         Height          =   2295
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4048
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
Attribute VB_Name = "FrmEntradaOrdenRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim IdOCR As String
Dim IdDetalle As String
Private Sub Command1_Click()
    ListView1.Enabled = True
    RECIBIDO
End Sub
Private Sub Command2_Click()
    IdOCR = Text8.Text
    RECIBIDO
    Text8.Text = ""
    IdOCR = ""
End Sub
Private Sub Command3_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim Cont As Integer
    Cont = 1
    ListView1.Enabled = False
    If Text1.Text <> "" Then
        If CDbl(Text3.Text) = CDbl(Text2.Text) Or CDbl(Text3.Text) < CDbl(Text2.Text) Then
            If Text4.Text <> "" And Text6.Text <> "" Then
                'COMENTARIO_SALIDA
                sBuscar = "UPDATE ORDEN_RAPIDA_DETALLE SET COMENTARIO_SALID = '" & Text6.Text & "', FACTURA = '" & Text5.Text & "', CAN_RECIBIDA = CAN_RECIBIDA + " & CDbl(Text3.Text) & " WHERE ID_PRODUCTO = '" & Text1.Text & "' AND ID_ORDEN_RAPIDA = " & IdOCR & " AND ID_DETALLE = " & IdDetalle
                cnn.Execute (sBuscar)
                sBuscar = "SELECT CAN_RECIBIDA, CANTIDAD FROM ORDEN_RAPIDA_DETALLE WHERE ID_PRODUCTO = '" & Text1.Text & "' AND ID_ORDEN_RAPIDA = " & IdOCR & " AND ID_DETALLE = " & IdDetalle
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    If tRs.Fields("CAN_RECIBIDA") = tRs.Fields("CANTIDAD") Then
                        sBuscar = "UPDATE ORDEN_RAPIDA_DETALLE SET SURTIDO = 'S' WHERE ID_PRODUCTO = '" & Text1.Text & "' AND ID_ORDEN_RAPIDA = " & IdOCR & " AND ID_DETALLE = " & IdDetalle
                        cnn.Execute (sBuscar)
                    End If
                End If
                sBuscar = "SELECT ID_PRODUCTO FROM PRODUCTOS_CONSUMIBLES WHERE ID_PRODUCTO = '" & Text1.Text & "' AND CONTROLA_EXISTENCIA = 'S'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    If Option1.Value = True Then
                        Do While Cont <= CDbl(Text3.Text)
                            sBuscar = "INSERT INTO EXISTENCIA_FIJA (ID_PRODUCTO, CANTIDAD, ID_ORDEN_RAPIDA, NUMERO_SERIE) VALUES ('" & Text1.Text & "', " & CDbl(Text3.Text) & ", " & IdOCR & ", '" & InputBox("CAPTURA DE NUMERO DE SERIE", "SACC") & "')"
                            Cont = Cont + 1
                        Loop
                    Else
                         sBuscar = "INSERT INTO EXISTENCIA_FIJA (ID_PRODUCTO, CANTIDAD, ID_ORDEN_RAPIDA) VALUES ('" & Text1.Text & "', " & CDbl(Text3.Text) & ", " & IdOCR & ")"
                    End If
                End If
                sBuscar = "UPDATE ORDEN_RAPIDA_DETALLE SET CAN_RECIBIDA = CANTIDAD WHERE CAN_RECIBIDA > CANTIDAD"
                cnn.Execute (sBuscar)
                Text1.Text = ""
                Text2.Text = ""
                Text3.Text = ""
                Text5.Text = ""
                Text6.Text = ""
            Else
                MsgBox "INGRESE  UN NUMERO DE FACTURA  O COMENTARIO POR FAVOR!", vbExclamation, "SACC"
            End If
       Else
            MsgBox "LA CANTIDAD RECIBIDA NO PUEDE SER MAYOR A LA COMPRADA!", vbExclamation, "SACC"
       End If
    End If
End Sub
Private Sub RECIBIDO()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim loca As Integer
    Dim sBuscar As String
    Dim PosVer As Integer
    Dim Posabo As Integer
    Dim tRs  As ADODB.Recordset
    Set oDoc = New cPDF
    Dim sumdeuda As Double
    Dim sumIndi As Double
    Dim sNombre As String
    Dim sumabonos As Double
    Dim sumIndiabonos As Double
    Dim AcumDeudas As String
    Dim NoRe As Integer
    Dim ConPag As Integer
    ConPag = 1
    If Not oDoc.PDFCreate(App.Path & "\EntraOrdenRapida.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
' Encabezado del reporte
    'Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    'oDoc.LoadImage Image2, "Logo", False, False
    oDoc.NewPage A4_Vertical
   ' oDoc.WImage 80, 40, 43, 161, "Logo"
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
    oDoc.WTextBox 90, 200, 20, 250, "Comprobante de Recibido e Orden Rapida", "F2", 10, hCenter
    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
    oDoc.WTextBox 100, 20, 30, 100, "ORDEN RAPIDA ", "F2", 9, hLeft
    oDoc.WTextBox 100, 100, 30, 80, "PRODUCTO ", "F2", 9, hLeft
    oDoc.WTextBox 100, 250, 40, 200, "CANTIDAD ", "F2", 9, hLeft
    oDoc.WTextBox 100, 330, 40, 300, "CANTIDAD RECIBIDA ", "F2", 9, hLeft
    oDoc.WTextBox 100, 450, 40, 300, "COMENTARIO ", "F2", 9, hLeft
    Posi = 110 + 10
    sBuscar = "SELECT * FROM ORDEN_RAPIDA_DETALLE WHERE CAN_RECIBIDA <> 0 AND ID_ORDEN_RAPIDA = " & IdOCR
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        'oDoc.WTextBox Posi, 440, 40, 300, tRs.Fields("COMENTARIO"), "F2", 9, hLeft
        Do While Not tRs.EOF
            oDoc.WTextBox Posi, 20, 30, 40, IdOCR, "F2", 9, hLeft
            oDoc.WTextBox Posi, 100, 30, 150, tRs.Fields("ID_PRODUCTO"), "F2", 9, hLeft
            oDoc.WTextBox Posi, 250, 40, 200, tRs.Fields("CANTIDAD"), "F2", 9, hLeft
            oDoc.WTextBox Posi, 330, 40, 300, tRs.Fields("CAN_RECIBIDA"), "F2", 9, hLeft
            'oDoc.WTextBox Posi, 440, 40, 300, tRs.Fields("COMENTARIO"), "F2", 9, hLeft
            Posi = Posi + 20
            tRs.MoveNext
        Loop
    End If
' Encabezado de pagina
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 115
    oDoc.WLineTo 580, 115
    oDoc.LineStroke
    Posi = Posi + 30
    oDoc.WTextBox Posi, 50, 40, 200, " RECIBIO", "F2", 9, hLeft
    oDoc.WTextBox Posi, 400, 40, 300, " ENTREGO", "F2", 9, hLeft
    If Posi >= 760 Then
        oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
        ConPag = ConPag + 1
        oDoc.NewPage A4_Vertical
        ' Encabezado del reporte
        Posi = 110
        oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
        oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
        oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
        oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
        oDoc.WTextBox 90, 200, 20, 250, "Comprobante de Recibido e Orden Rapida", "F2", 10, hCenter
        oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
        oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
        ' Encabezado de pagina
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 100
        oDoc.WLineTo 580, 100
        oDoc.LineStroke
    End If
    Cont = Cont + 1
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
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
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Num. Orden", 1000
        .ColumnHeaders.Add , , "Proveedor", 4200
        .ColumnHeaders.Add , , "Fecha", 1000
        .ColumnHeaders.Add , , "Rentas", 1000
        .ColumnHeaders.Add , , "Fac. Nota", 1000
        .ColumnHeaders.Add , , "Total Prov.", 1000
        .ColumnHeaders.Add , , "Retención", 1000
        .ColumnHeaders.Add , , "Total Ret.", 1000
        .ColumnHeaders.Add , , "Comentario", 3000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave", 2000
        .ColumnHeaders.Add , , "Descripcion", 5200
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Precio", 1000
        .ColumnHeaders.Add , , "Surtido", 1000
        .ColumnHeaders.Add , , "Id Detalle", 0
    End With
    BuscaOrden
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub BuscaOrden()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT ID_ORDEN_RAPIDA, FECHA, NOMBRE, COMENTARIO, RENTAS, FAC_NOTA, TOT_PROV, RETENCION, TOTAL_RETENCION FROM VsRecibeOrdenRapida WHERE SURTIDO = 'N' GROUP BY ID_ORDEN_RAPIDA, FECHA, NOMBRE, COMENTARIO, RENTAS, FAC_NOTA, TOT_PROV, RETENCION, TOTAL_RETENCION ORDER BY ID_ORDEN_RAPIDA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_ORDEN_RAPIDA"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(2) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("RENTAS")) Then tLi.SubItems(3) = tRs.Fields("RENTAS")
            If Not IsNull(tRs.Fields("FAC_NOTA")) Then tLi.SubItems(4) = tRs.Fields("FAC_NOTA")
            If Not IsNull(tRs.Fields("TOT_PROV")) Then tLi.SubItems(5) = tRs.Fields("TOT_PROV")
            If Not IsNull(tRs.Fields("RETENCION")) Then tLi.SubItems(6) = tRs.Fields("RETENCION")
            If Not IsNull(tRs.Fields("TOTAL_RETENCION")) Then tLi.SubItems(7) = tRs.Fields("TOTAL_RETENCION")
            If Not IsNull(tRs.Fields("COMENTARIO")) Then tLi.SubItems(8) = tRs.Fields("COMENTARIO")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView2.ListItems.Clear
    IdOCR = Item
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = Item
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = Item.SubItems(1)
    sBuscar = "SELECT  ORDEN_RAPIDA.ID_ORDEN_RAPIDA, ORDEN_RAPIDA.ESTADO, ORDEN_RAPIDA.FECHA, PROVEEDOR_CONSUMO.NOMBRE, ORDEN_RAPIDA.MONEDA, ORDEN_RAPIDA.COMENTARIO, ORDEN_RAPIDA.RENTAS, ORDEN_RAPIDA.FAC_NOTA, ORDEN_RAPIDA.TOT_PROV, ORDEN_RAPIDA.RETENCION, ORDEN_RAPIDA.TOTAL_RETENCION, ORDEN_RAPIDA_DETALLE.ID_PRODUCTO, ORDEN_RAPIDA_DETALLE.CANTIDAD, ORDEN_RAPIDA_DETALLE.PRECIO, ORDEN_RAPIDA_DETALLE.Surtido , PRODUCTOS_CONSUMIBLES.Descripcion, IsNull(ORDEN_RAPIDA_DETALLE.CAN_RECIBIDA, 0) AS CAN_RECIBIDA, ORDEN_RAPIDA_DETALLE.ID_DETALLE FROM ORDEN_RAPIDA INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA INNER JOIN PROVEEDOR_CONSUMO ON ORDEN_RAPIDA.ID_PROVEEDOR = PROVEEDOR_CONSUMO.ID_PROVEEDOR INNER JOIN PRODUCTOS_CONSUMIBLES ON ORDEN_RAPIDA_DETALLE.ID_PRODUCTO = PRODUCTOS_CONSUMIBLES.ID_PRODUCTO " & _
    "WHERE (ORDEN_RAPIDA.ESTADO IN ('A', 'F')) AND (PRODUCTOS_CONSUMIBLES.CONTROLA_EXISTENCIA IN ('S', 'E')) AND ORDEN_RAPIDA.ID_ORDEN_RAPIDA = " & Item & " AND ORDEN_RAPIDA_DETALLE.SURTIDO = 'N'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD") - tRs.Fields("CAN_RECIBIDA")
            If Not IsNull(tRs.Fields("PRECIO")) Then tLi.SubItems(3) = tRs.Fields("PRECIO")
            If Not IsNull(tRs.Fields("SURTIDO")) Then tLi.SubItems(4) = tRs.Fields("SURTIDO")
            If Not IsNull(tRs.Fields("ID_DETALLE")) Then tLi.SubItems(5) = tRs.Fields("ID_DETALLE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item
    Text2.Text = Item.SubItems(2)
    Text3.Text = Item.SubItems(2)
    IdDetalle = Item.SubItems(5)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
