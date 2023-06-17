VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCheque 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de Cheque"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab2 
      Height          =   2175
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3836
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Nacionales"
      TabPicture(0)   =   "FrmCheque.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Internacionales"
      TabPicture(1)   =   "FrmCheque.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Ordeon Rapida"
      TabPicture(2)   =   "FrmCheque.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView3"
      Tab(2).ControlCount=   1
      Begin MSComctlLib.ListView ListView3 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   24
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2566
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   23
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2566
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2990
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7920
      TabIndex        =   10
      Top             =   2520
      Width           =   975
      Begin VB.Image Command2 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FrmCheque.frx":0054
         MousePointer    =   99  'Custom
         Picture         =   "FrmCheque.frx":035E
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Cheque"
      TabPicture(0)   =   "FrmCheque.frx":1F30
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtNum2Let(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtNum2Let(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtNum2Let(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TxtNOMBRE"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DTPicker1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Combo1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtNUM_CHEQUE"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtNUM_ORDEN"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtTIPO_ORDEN"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Check1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Command4"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtIdProv"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      Begin VB.TextBox TxtIdProv 
         Height          =   285
         Left            =   2280
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   195
         Left            =   7200
         TabIndex        =   25
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   ".."
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
         Left            =   3000
         TabIndex        =   20
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pago en Efectivo de Caja Chica"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox TxtTIPO_ORDEN 
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxtNUM_ORDEN 
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxtNUM_CHEQUE 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   16
         Top             =   2160
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   2160
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49872897
         CurrentDate     =   39385
      End
      Begin VB.TextBox TxtNOMBRE 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   600
         Width           =   6735
      End
      Begin VB.TextBox txtNum2Let 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   3
         Top             =   1800
         Width           =   5895
      End
      Begin VB.TextBox txtNum2Let 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   ".."
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
         Left            =   7320
         TabIndex        =   4
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox txtNum2Let 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Numero de Cheque :"
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Banco :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Total con Letra :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Total :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   495
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   7920
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FrmCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cALetra As New clsNum2Let
Private cnn As ADODB.Connection
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        TxtNUM_CHEQUE.Enabled = False
        TxtNUM_CHEQUE.Text = ""
        Combo1.Enabled = False
        Combo1.Text = ""
    Else
        TxtNUM_CHEQUE.Enabled = True
        TxtNUM_CHEQUE.Text = ""
        Combo1.Enabled = True
        Combo1.Text = ""
    End If
End Sub
Private Sub Command1_Click()
    txtNum2Let(2).Enabled = True
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    If (TxtNOMBRE.Text <> "" And txtNum2Let(0).Text <> "" And Combo1.Text <> "") Or Me.Check1.Value = 1 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        Dim Path As String
        If Check1.Value = 0 Then
            sBuscar = "INSERT INTO CHEQUES(NOMBRE, FECHA, TOTAL, TOTAL_LETRA, BANCO, NUM_CHEQUE, TIPO_ORDEN, NUM_ORDEN, FECHA_REALIZADO, ID_USUARIO) VALUES('" & TxtNOMBRE.Text & "', '" & DTPicker1.Value & "', '" & CDbl(Replace(txtNum2Let(1).Text, "$", "")) & "', '" & txtNum2Let(2).Text & "', '" & Combo1.Text & "', '" & TxtNUM_CHEQUE.Text & "', '" & TxtTIPO_ORDEN.Text & "', '" & TxtNUM_ORDEN.Text & "', '" & Format(Date, "dd/mm/yyyy") & "', '" & VarMen.Text1(0).Text & "')"
        Else
            sBuscar = "INSERT INTO CAJA_CHICA(NOMBRE, FECHA, TOTAL, TOTAL_LETRA, TIPO_ORDEN, NUM_ORDEN, FECHA_REALIZADO, ID_USUARIO) VALUES('" & TxtNOMBRE.Text & "', '" & DTPicker1.Value & "', '" & CDbl(Replace(txtNum2Let(1).Text, "$", "")) & "', '" & txtNum2Let(2).Text & "', '" & TxtTIPO_ORDEN.Text & "', '" & Mid(TxtNUM_ORDEN.Text, 1, Len(TxtNUM_ORDEN.Text) - 2) & "', '" & Format(Date, "dd/mm/yyyy") & "', " & VarMen.Text1(0).Text & ")"
        End If
        Set tRs = cnn.Execute(sBuscar)
        If Check1.Value = 0 Then
            cheque
        Else
            efectivo
        End If
        If TxtIdProv.Text <> "" Then
            If TxtTIPO_ORDEN.Text = "RAPIDA" Then
                sBuscar = "SELECT EMAIL FROM PROVEEDOR_CONSUMO WHERE ID_PROVEEDOR = " & TxtIdProv.Text & " AND EMAIL IS NOT NULL"
            Else
                sBuscar = "SELECT EMAIL FROM PROVEEDOR WHERE ID_PROVEEDOR = " & TxtIdProv.Text & " AND EMAIL IS NOT NULL"
            End If
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                EnviaCorreo (tRs.Fields("EMAIL"))
            End If
        End If
        Unload Me
    Else
        MsgBox "FALTA INFORMACIÓN NECESARIA PARA EL REGISTRO!", vbInformation, "SACC"
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command3_Click()
    txtNum2Let(1).Enabled = True
End Sub
Private Sub Command4_Click()
    cheque
End Sub
Private Sub Form_Load()
On Error GoTo ManejaError
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
           "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT NOMBRE FROM BANCOS ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
    'With ListView1
    '    .View = lvwReport
    '    .GridLines = True
    '    .LabelEdit = lvwManual
    '    .HideSelection = False
    '    .HotTracking = False
    '    .HoverSelection = False
    '    .ColumnHeaders.Add , , "ID_ORDEN", 0
    '    .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
    '    .ColumnHeaders.Add , , "FOLIO", 500
    '    .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
    '    .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
    '    .ColumnHeaders.Add , , "COMENTARIO", 1440
    'End With
    'With ListView2
    '    .View = lvwReport
    '    .GridLines = True
    '    .LabelEdit = lvwManual
    '    .HideSelection = False
    '    .HotTracking = False
    '    .HoverSelection = False
    '    .ColumnHeaders.Add , , "ID_ORDEN", 0
    '    .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
    '    .ColumnHeaders.Add , , "FOLIO", 500
    '    .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
    '    .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
    '    .ColumnHeaders.Add , , "COMENTARIO", 1440
    'End With
    'With ListView3
    '    .View = lvwReport
    '    .GridLines = True
    '    .LabelEdit = lvwManual
    '    .HideSelection = False
    '    .HotTracking = False
    '    .HoverSelection = False
    '    .ColumnHeaders.Add , , "ID_ORDEN", 0
    '    .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
    '    .ColumnHeaders.Add , , "FOLIO", 500
    '    .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
    '    .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
    '    .ColumnHeaders.Add , , "COMENTARIO", 1440
    'End With
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtNum2Let_Change(Index As Integer)
    If Index = 0 Then
        If Val(txtNum2Let(0).Text) > 0 Then
            txtNum2Let(1).Text = Format(Val(txtNum2Let(0).Text), "Currency")
            cALetra.Numero = Val(txtNum2Let(0).Text)
            txtNum2Let(2).Text = cALetra.ALetra
        Else
            txtNum2Let(1).Text = ""
            txtNum2Let(2).Text = ""
        End If
    End If
End Sub
Private Sub cheque()
On Error GoTo ManejaError
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim sBusca As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim totprod As Double
    Dim TotAbonos As Double
    Dim sqlQuery As String
    Dim TotOrden As Double
    Dim Restante As Double
    Dim sBuscar As String
    Dim PosX As Integer
    Dim NumerosOrdenes As String
    Dim sDolar As Double
    Dim Dolar As Double
    Dim TipoPago As String
    Set oDoc = New cPDF
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Not oDoc.PDFCreate(App.Path & "\Cheque.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    oDoc.LoadImage Image1, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 50, 40, 38, 161, "Logo"
    sBuscar = "SELECT TOP 1 VENTA FROM DOLAR ORDER BY ID_DOLAR DESC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        sDolar = tRs.Fields("VENTA")
    End If
    sqlQuery = "SELECT TOP 1 ID_CHEQUE FROM VSCHEQUES2 ORDER BY ID_CHEQUE DESC" '
    Set tRs = cnn.Execute(sqlQuery)
    sBusca = tRs.Fields("ID_CHEQUE")
    sqlQuery = "SELECT * FROM VSCHEQUES2 WHERE ID_CHEQUE='" & tRs.Fields("ID_CHEQUE") & "' "
    Set tRs2 = cnn.Execute(sqlQuery)
    sqlQuery = "SELECT TOP 1 TIPOPAGO FROM ABONOS_PAGO_OC ORDER BY ID_ABONO DESC" '
    Set tRs3 = cnn.Execute(sqlQuery)
    If Not (tRs3.EOF And tRs3.BOF) Then
        If Not IsNull(tRs3.Fields("TIPOPAGO")) Then
            TipoPago = tRs3.Fields("TIPOPAGO")
        Else
            TipoPago = "PAGO"
        End If
    Else
        TipoPago = "PAGO"
    End If
    'cuadros encabezado
    'Posi = 50
    oDoc.WTextBox 25, 300, 15, 300, "POLIZA DE " & TipoPago, "F2", 20, hLeft
    ''''primer  cuadro
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 50
    oDoc.WLineTo 580, 50
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 230
    oDoc.WLineTo 580, 230
    oDoc.LineStroke
    'Posi = Posi + 6
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 50
    oDoc.WLineTo 10, 230
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 580, 50
    oDoc.WLineTo 580, 230
    oDoc.LineStroke
    '''final del primer
    ''''segundo  cuadro
    oDoc.WTextBox 240, 30, 20, 300, "CONCEPTO " & TipoPago & " :", "F3", 8, hLeft
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 235
    oDoc.WLineTo 400, 235
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 250
    oDoc.WLineTo 400, 250
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 320
    oDoc.WLineTo 400, 320
    oDoc.LineStroke
    'Posi = Posi + 6
    'Posi = Posi + 6
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 235
    oDoc.WLineTo 10, 320
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 400, 235
    oDoc.WLineTo 400, 320
    oDoc.LineStroke
    '''final del segundo
    ''''  tercero
    oDoc.WTextBox 240, 415, 20, 200, "FIRMA DE " & TipoPago & " RECIBIDO  :", "F3", 8, hLeft
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 410, 235
    oDoc.WLineTo 580, 235
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 410, 250
    oDoc.WLineTo 580, 250
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 410, 320
    oDoc.WLineTo 580, 320
    oDoc.LineStroke
    'Posi = Posi + 6
    'Posi = Posi + 6
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 410, 235
    oDoc.WLineTo 410, 320
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 580, 235
    oDoc.WLineTo 580, 320
    oDoc.LineStroke
    '''final del tercero
    ''''  cuarto
    oDoc.WTextBox 328, 30, 20, 50, "CUENTA:", "F3", 8, hLeft
    oDoc.WTextBox 328, 90, 20, 70, "SUB-CUENTA:", "F3", 8, hLeft
    oDoc.WTextBox 328, 210, 20, 50, "NOMBRE:", "F3", 8, hLeft
    oDoc.WTextBox 328, 350, 20, 50, "PARCIAL:", "F3", 8, hLeft
    oDoc.WTextBox 328, 450, 20, 50, "DEBER:", "F3", 8, hLeft
    oDoc.WTextBox 328, 540, 20, 50, "HABER:", "F3", 8, hLeft
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 325
    oDoc.WLineTo 580, 325
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 340
    oDoc.WLineTo 580, 340
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 400
    oDoc.WLineTo 580, 400
    oDoc.LineStroke
    'Posi = Posi + 6
    'Posi = Posi + 6
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 325
    oDoc.WLineTo 10, 400
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 580, 325
    oDoc.WLineTo 580, 400
    oDoc.LineStroke
    '''lineas de
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 80, 325
    oDoc.WLineTo 80, 400
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 160, 325
    oDoc.WLineTo 160, 400
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 330, 325
    oDoc.WLineTo 330, 400
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 420, 325
    oDoc.WLineTo 420, 400
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 500, 325
    oDoc.WLineTo 500, 400
    oDoc.LineStroke
    '''final del cuarto
    ''''  quinto
    oDoc.WTextBox 407, 30, 20, 70, "HECHO POR:", "F3", 8, hLeft
    oDoc.WTextBox 407, 90, 20, 60, "REVISADO:", "F3", 8, hLeft
    oDoc.WTextBox 407, 210, 20, 70, "AUTORIZADO:", "F3", 8, hLeft
    oDoc.WTextBox 440, 200, 20, 150, VarMen.TxtEmp(11).Text & ":", "F3", 8, hLeft
    oDoc.WTextBox 407, 350, 20, 50, "AUXILIARES:", "F3", 8, hLeft
    oDoc.WTextBox 407, 450, 20, 50, "DIARIO:", "F3", 8, hLeft
    oDoc.WTextBox 407, 540, 20, 50, "HABER:", "F3", 8, hLeft
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 405
    oDoc.WLineTo 580, 405
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 420
    oDoc.WLineTo 580, 420
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 460
    oDoc.WLineTo 580, 460
    oDoc.LineStroke
    'Posi = Posi + 6
    'Posi = Posi + 6
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 405
    oDoc.WLineTo 10, 460
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 580, 405
    oDoc.WLineTo 580, 460
    oDoc.LineStroke
    '''lineas de
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 80, 405
    oDoc.WLineTo 80, 460
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 160, 405
    oDoc.WLineTo 160, 460
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 330, 405
    oDoc.WLineTo 330, 460
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 420, 405
    oDoc.WLineTo 420, 460
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 500, 405
    oDoc.WLineTo 500, 460
    oDoc.LineStroke
    '''final del quinto
    Posi = 230
    If Not (tRs2.EOF And tRs2.BOF) Then
        Do While Not (tRs2.EOF)
            oDoc.WTextBox 70, 20, 20, 300, tRs2.Fields("NOMBRE"), "F3", 8, hLeft
            oDoc.WTextBox 55, 400, 20, 50, tRs2.Fields("FECHA"), "F3", 8, hLeft
            oDoc.WTextBox 70, 420, 20, 100, Format(tRs2.Fields("TOTAL"), "$ ###,###,##0.00"), "F3", 8, hLeft
            oDoc.WTextBox 88, 20, 20, 300, tRs2.Fields("TOTAL_LETRA"), "F3", 8, hLeft
            oDoc.WTextBox 200, 130, 20, 80, TipoPago & ": " & tRs2.Fields("NUM_CHEQUE"), "F3", 8, hLeft
            oDoc.WTextBox 200, 320, 20, 50, tRs2.Fields("BANCO"), "F3", 8, hLeft
            oDoc.WTextBox 260, 20, 20, 80, "PAGO DE O.C No", "F3", 8, hLeft
            sBusca = "SELECT NUM_ORDEN FROM VSCHEQUES2 WHERE NUM_CHEQUE = '" & tRs2.Fields("NUM_CHEQUE") & "' AND BANCO = '" & tRs2.Fields("BANCO") & "' ORDER BY ID_CHEQUE"
            Set tRs3 = cnn.Execute(sBusca)
            If Not (tRs3.EOF And tRs3.BOF) Then
                Do While Not tRs3.EOF
                    NumerosOrdenes = Me.TxtNUM_ORDEN.Text 'NumerosOrdenes & tRs3.Fields("NUM_ORDEN") & ", "
                    tRs3.MoveNext
                Loop
            End If
            oDoc.WTextBox 260, 100, 20, 270, Mid(NumerosOrdenes, 1, Len(NumerosOrdenes) - 2), "F3", 8, hLeft
            oDoc.WTextBox 290, 120, 20, 300, tRs2.Fields("NOMBRE"), "F3", 8, hLeft
            oDoc.WTextBox 300, 250, 20, 80, TipoPago & ":" & tRs2.Fields("NUM_CHEQUE"), "F3", 8, hLeft
            'oDoc.WTextBox 300, 300, 20, 80, tRs2.Fields("NUM_CHEQUE"), "F3", 8, hLeft
            tRs2.MoveNext
        Loop
    End If
    Posi = 470
    TotAbonos = 0
    '////////////////////////////////// PAGOS ANTERIORES A LA MISMA ORDEN /////////////////////////////////////////////////
    oDoc.WTextBox Posi, 30, 20, 50, "ABONO", "F2", 8, hCenter
    oDoc.WTextBox Posi, 80, 20, 80, "RESTANTE", "F2", 8, hCenter
    PosX = 30
    If TxtTIPO_ORDEN.Text = "INTERNACIONAL" Then
        sBuscar = "SELECT TOTAL, DISCOUNT, FREIGHT, TAX, OTROS_CARGOS, MONEDA FROM ORDEN_COMPRA WHERE NUM_ORDEN IN (" & Mid(TxtNUM_ORDEN.Text, 1, Len(TxtNUM_ORDEN.Text) - 2) & ") AND TIPO = 'I'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                If tRs.Fields("MONEDA") = "DOLARES" Then
                    TotOrden = TotOrden + (tRs.Fields("TOTAL") + tRs.Fields("TAX") + tRs.Fields("FREIGHT") + tRs.Fields("OTROS_CARGOS") - tRs.Fields("DISCOUNT")) * sDolar
                    Restante = TotOrden
                Else
                    TotOrden = TotOrden + (tRs.Fields("TOTAL") + tRs.Fields("TAX") + tRs.Fields("FREIGHT") + tRs.Fields("OTROS_CARGOS") - tRs.Fields("DISCOUNT"))
                    Restante = TotOrden
                End If
                tRs.MoveNext
            Loop
        End If
        sBuscar = "SELECT FECHA, BANCO, CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN IN (" & Mid(NumerosOrdenes, 1, Len(NumerosOrdenes) - 2) & ") AND TIPO = 'I' ORDER BY ID_ABONO"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                If Restante >= 0.01 Or Restante = 0 Then
                    Posi = Posi + 10
                    TotAbonos = TotAbonos + tRs.Fields("CANT_ABONO")
                    oDoc.WTextBox Posi, PosX, 20, 50, Format(tRs.Fields("CANT_ABONO"), "$ ###,###,##0.00"), "F3", 8, hRight
                    Restante = Restante - tRs.Fields("CANT_ABONO")
                    oDoc.WTextBox Posi, PosX + 50, 20, 80, Format(Restante, "$ ###,###,##0.00"), "F3", 8, hRight
                    If Posi > 760 Then
                        PosX = PosX + 140
                        Posi = 470
                        oDoc.WTextBox Posi, PosX, 20, 50, "ABONO", "F2", 8, hCenter
                        oDoc.WTextBox Posi, PosX + 50, 20, 80, "RESTANTE", "F2", 8, hCenter
                    End If
                End If
                tRs.MoveNext
            Loop
        End If
        Posi = Posi + 20
        If Not (tRs.EOF And tRs.BOF) Then
            oDoc.WTextBox Posi, PosX, 20, 100, "Total Orden :", "F2", 8, hLeft
            oDoc.WTextBox Posi, PosX + 100, 20, 100, Format(TotOrden * sDolar, "$ ###,###,##0.00"), "F3", 8, hRight
            Posi = Posi + 10
            oDoc.WTextBox Posi, PosX, 20, 100, "Restante Estimado :", "F2", 8, hLeft
            oDoc.WTextBox Posi, PosX + 100, 20, 100, Format((TotOrden * sDolar) - TotAbonos, "$ ###,###,##0.00"), "F3", 8, hRight
        End If
    End If
    If TxtTIPO_ORDEN.Text = "NACIONAL" Then
        sBuscar = "SELECT TOTAL, DISCOUNT, FREIGHT, TAX, OTROS_CARGOS, MONEDA FROM ORDEN_COMPRA WHERE NUM_ORDEN IN (" & Mid(TxtNUM_ORDEN.Text, 1, Len(TxtNUM_ORDEN.Text) - 2) & ") AND TIPO = 'N'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                If tRs.Fields("MONEDA") = "DOLARES" Then
                    TotOrden = TotOrden + (tRs.Fields("TOTAL") + tRs.Fields("TAX") + tRs.Fields("FREIGHT") + tRs.Fields("OTROS_CARGOS") - tRs.Fields("DISCOUNT")) * sDolar
                    Restante = TotOrden
                Else
                    TotOrden = TotOrden + (tRs.Fields("TOTAL") + tRs.Fields("TAX") + tRs.Fields("FREIGHT") + tRs.Fields("OTROS_CARGOS") - tRs.Fields("DISCOUNT"))
                    Restante = TotOrden
                End If
                tRs.MoveNext
            Loop
        End If
        sBuscar = "SELECT FECHA, BANCO, CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN IN (" & Mid(TxtNUM_ORDEN.Text, 1, Len(TxtNUM_ORDEN.Text) - 2) & ") AND TIPO = 'N' ORDER BY ID_ABONO"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                If Restante >= 0.01 Then
                    Posi = Posi + 10
                    TotAbonos = TotAbonos + tRs.Fields("CANT_ABONO")
                    oDoc.WTextBox Posi, PosX, 20, 50, Format(tRs.Fields("CANT_ABONO"), "$ ###,###,##0.00"), "F3", 8, hRight
                    Restante = Restante - tRs.Fields("CANT_ABONO")
                    oDoc.WTextBox Posi, PosX + 50, 20, 80, Format(Restante, "$ ###,###,##0.00"), "F3", 8, hRight
                    If Posi > 760 Then
                        PosX = PosX + 140
                        Posi = 470
                        oDoc.WTextBox Posi, PosX, 20, 50, "ABONO", "F2", 8, hCenter
                        oDoc.WTextBox Posi, PosX + 50, 20, 80, "RESTANTE", "F2", 8, hCenter
                    End If
                End If
                tRs.MoveNext
            Loop
        End If
        Posi = Posi + 20
        If Not (tRs.EOF And tRs.BOF) Then
            oDoc.WTextBox Posi, PosX, 20, 100, "Total Orden :", "F2", 8, hLeft
            oDoc.WTextBox Posi, PosX + 100, 20, 100, Format(TotOrden, "$ ###,###,##0.00"), "F3", 8, hRight
            Posi = Posi + 10
            oDoc.WTextBox Posi, PosX, 20, 100, "Restante Estimado :", "F2", 8, hLeft
            oDoc.WTextBox Posi, PosX + 100, 20, 100, Format(TotOrden - TotAbonos, "$ ###,###,##0.00"), "F3", 8, hRight
        End If
    End If
    If TxtTIPO_ORDEN.Text = "RAPIDA" Then
        sBuscar = "SELECT SUM (TOTAL) AS TOTAL, MONEDA FROM VsOrdenRapida WHERE ID_ORDEN_RAPIDA IN (" & Mid(TxtNUM_ORDEN.Text, 1, Len(TxtNUM_ORDEN.Text) - 2) & ") GROUP BY MONEDA"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                If tRs.Fields("MONEDA") = "DOLARES" Then
                    TotOrden = TotOrden + tRs.Fields("TOTAL") * sDolar
                    Restante = TotOrden
                Else
                    TotOrden = TotOrden + tRs.Fields("TOTAL")
                    Restante = TotOrden
                End If
                tRs.MoveNext
            Loop
        End If
        sBuscar = "SELECT FECHA, BANCO, CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN IN (" & Mid(TxtNUM_ORDEN.Text, 1, Len(TxtNUM_ORDEN.Text) - 2) & ") AND TIPO = 'R' ORDER BY ID_ABONO"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                If Restante >= 0.01 Then
                    Posi = Posi + 10
                    TotAbonos = TotAbonos + tRs.Fields("CANT_ABONO")
                    oDoc.WTextBox Posi, PosX, 20, 50, Format(tRs.Fields("CANT_ABONO"), "$ ###,###,##0.00"), "F3", 8, hRight
                    Restante = Restante - tRs.Fields("CANT_ABONO")
                    oDoc.WTextBox Posi, PosX + 50, 20, 80, Format(Restante, "$ ###,###,##0.00"), "F3", 8, hRight
                    If Posi > 760 Then
                        PosX = PosX + 140
                        Posi = 470
                        oDoc.WTextBox Posi, PosX, 20, 50, "ABONO", "F2", 8, hCenter
                        oDoc.WTextBox Posi, PosX + 50, 20, 80, "RESTANTE", "F2", 8, hCenter
                    End If
                End If
                If Restante < 0 Then
                Restante = 0
                End If
                tRs.MoveNext
            Loop
        End If
        Posi = Posi + 20
        If Not (tRs.EOF And tRs.BOF) Then
            oDoc.WTextBox Posi, PosX, 20, 100, "Total Orden :", "F2", 8, hLeft
            oDoc.WTextBox Posi, PosX + 100, 20, 100, Format(TotOrden, "$ ###,###,##0.00"), "F3", 8, hRight
            Posi = Posi + 10
            oDoc.WTextBox Posi, PosX, 20, 100, "Restante Estimado :", "F2", 8, hLeft
            oDoc.WTextBox Posi, PosX + 100, 20, 100, Format((TotOrden - TotAbonos), "$ ###,###,##0.00"), "F3", 8, hRight
        End If
    End If
    '/////////////////////////////////////////////// PAGOS CANCELADOS ///////////////////////////////////////////////////////
    'Posi = 470
    'oDoc.WTextBox Posi, 320, 20, 100, "CANCELACION", "F2", 8, hCenter
    'oDoc.WTextBox Posi, 420, 20, 100, "IMPORTE", "F2", 8, hCenter
    'If TxtTIPO_ORDEN.Text <> "ALMACEN1" Then
    '    If TxtTIPO_ORDEN.Text = "NACIONAL" Then
    '        sBuscar = "SELECT FECHA, BANCO, CANT_ABONO FROM ABONOS_OC_CANCELADOS WHERE NUM_ORDEN IN (" & Mid(TxtNUM_ORDEN.Text, 1, Len(TxtNUM_ORDEN.Text) - 2) & ") AND TIPO = 'N'"
    '    End If
    '    If TxtTIPO_ORDEN.Text = "INTERNACIONAL" Then
    '        sBuscar = "SELECT FECHA, BANCO, CANT_ABONO FROM ABONOS_OC_CANCELADOS WHERE NUM_ORDEN IN (" & Mid(TxtNUM_ORDEN.Text, 1, Len(TxtNUM_ORDEN.Text) - 2) & ") AND TIPO = 'I'"
    '    End If
    '    If TxtTIPO_ORDEN.Text = "RAPIDA" Then
    '        sBuscar = "SELECT FECHA, BANCO, CANT_ABONO FROM ABONOS_OC_CANCELADOS WHERE NUM_ORDEN IN (" & Mid(TxtNUM_ORDEN.Text, 1, Len(TxtNUM_ORDEN.Text) - 2) & ") AND TIPO = 'R'"
    '    End If
    '    Set tRs = cnn.Execute(sBuscar)
    '    If Not (tRs.EOF And tRs.BOF) Then
    '        Do While Not tRs.EOF
    '            Posi = Posi + 10
    '            TotAbonos = TotAbonos + tRs.Fields("CANT_ABONO")
    '            oDoc.WTextBox Posi, 320, 20, 100, tRs.Fields("FECHA"), "F3", 8, hLeft
    '            oDoc.WTextBox Posi, 420, 20, 100, Format(tRs.Fields("CANT_ABONO"), "$ ###,###,##0.00"), "F3", 8, hRight
    '            tRs.MoveNext
    '        Loop
    '    End If
    'End If
    'cierre del reporte
    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub efectivo()
On Error GoTo ManejaError
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim sBusca As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim totprod As Double
    Dim TotAbonos As Double
    Dim sqlQuery As String
    Dim TotOrden As Double
    Dim Restante As Double
    Dim sBuscar As String
    Dim PosX As Integer
    Dim NumerosOrdenes As String
    Dim sDolar As Double
    Dim Dolar As Double
    Set oDoc = New cPDF
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Not oDoc.PDFCreate(App.Path & "\Efectivo.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    oDoc.LoadImage Image1, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 50, 40, 38, 161, "Logo"
    sBuscar = "SELECT TOP 1 VENTA FROM DOLAR ORDER BY ID_DOLAR DESC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        sDolar = tRs.Fields("VENTA")
    End If
    sqlQuery = "SELECT TOP 1 ID_CHEQUE FROM VSCHEQUES2 ORDER BY ID_CHEQUE DESC" '
    Set tRs = cnn.Execute(sqlQuery)
    'sBusca = tRs.Fields("ID_CHEQUE")
    'sqlQuery = "SELECT * FROM VSCHEQUES2 WHERE ID_CHEQUE='" & tRs.Fields("ID_CHEQUE") & "' "
    'Set tRs2 = cnn.Execute(sqlQuery)
    
    'cuadros encabezado
    oDoc.WTextBox 25, 380, 15, 200, "PAGO EN EFECTIVO", "F2", 20, hLeft
    Posi = 230
    oDoc.WTextBox 70, 20, 20, 300, TxtNOMBRE.Text, "F3", 8, hLeft
    oDoc.WTextBox 55, 400, 20, 50, DTPicker1.Value, "F3", 8, hLeft
    oDoc.WTextBox 70, 420, 20, 100, Format(txtNum2Let(1).Text, "$ ###,###,##0.00"), "F3", 8, hLeft
    oDoc.WTextBox 88, 20, 20, 300, txtNum2Let(2).Text, "F3", 8, hLeft
    oDoc.WTextBox 200, 190, 20, 80, "PAGO EN EFECTIVO ", "F3", 8, hLeft
    'oDoc.WTextBox 200, 320, 20, 50, tRs2.Fields("BANCO"), "F3", 8, hLeft
    oDoc.WTextBox 260, 20, 20, 80, "PAGO DE O.C No", "F3", 8, hLeft
    'sBusca = "SELECT NUM_ORDEN FROM VSCHEQUES2 WHERE NUM_CHEQUE = '" & tRs2.Fields("NUM_CHEQUE") & "' AND BANCO = '" & tRs2.Fields("BANCO") & "' ORDER BY ID_CHEQUE"
    'Set tRs3 = cnn.Execute(sBusca)
    'If Not (tRs3.EOF And tRs3.BOF) Then
    '    Do While Not tRs3.EOF
    '        NumerosOrdenes = Me.TxtNUM_ORDEN.Text 'NumerosOrdenes & tRs3.Fields("NUM_ORDEN") & ", "
    '        tRs3.MoveNext
    '    Loop
    'End If
    oDoc.WTextBox 260, 100, 20, 270, NumerosOrdenes, "F3", 8, hLeft
    oDoc.WTextBox 290, 120, 20, 300, TxtNOMBRE.Text, "F3", 8, hLeft
    'tRs2.MoveNext
    Posi = 470
    TotAbonos = 0
    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub EnviaCorreo(MAIL As String)
On Error GoTo ManejaError
    If GetSetting("APTONER", "ConfigSACC", "Correo", "") <> "" Then
        Dim email As CDO.Message
        Dim correo As String
        Dim passwd As String
        Dim destino As String
        Dim Asunto As String
        Dim cuerpo As String
        Set email = New CDO.Message
        correo = GetSetting("APTONER", "ConfigSACC", "Correo", "")  ' "sistemas2@aptoner.com.mx"
        passwd = GetSetting("APTONER", "ConfigSACC", "CorreoPass", "")  ' "@Pt171218." Contraseña Generada por Gmail
        destino = MAIL '"control.sistemas.aptoner@gmail.com"
        Asunto = "Programación de Pago"
        cuerpo = "La empresa " & VarMen.TxtEmp(0).Text & " acaba de programar el pago de la(s) orden(es) de compra " & TxtNUM_ORDEN.Text & " por un monto de " & txtNum2Let(1).Text & ", en las proximas horas podrá ver el pago reflejado en su banco"
        email.Configuration.Fields(cdoSMTPServer) = GetSetting("APTONER", "ConfigSACC", "SMTP", "")  '"aptoner.com.mx"
        email.Configuration.Fields(cdoSendUsingMethod) = 2
        With email.Configuration.Fields
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(GetSetting("APTONER", "ConfigSACC", "PuertoCorreo", ""))  ' 26
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = Abs(1)
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = correo
            .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = passwd
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
        End With
        With email
            .To = destino
            .From = correo
            .Subject = Asunto
            .TextBody = cuerpo
            .Configuration.Fields.Update
            .Send
        End With
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
