VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmPagosOrdenes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de compra pendientes de pago"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   12120
      TabIndex        =   42
      Top             =   3000
      Width           =   975
      Begin VB.Image Image3 
         Height          =   780
         Left            =   120
         MouseIcon       =   "FrmPagosOrdenes.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmPagosOrdenes.frx":030A
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label11 
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
         TabIndex        =   43
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   6375
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Ordenes"
      TabPicture(0)   =   "FrmPagosOrdenes.frx":208C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Nacionales"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lvwOCIndirectas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lvwOCNacionales"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lvwOCInternacionales"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Saldos"
      TabPicture(1)   =   "FrmPagosOrdenes.frx":20A8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwOrdenesNSurtidas"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView lvwOrdenesNSurtidas 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   39
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   9340
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
      Begin MSComctlLib.ListView lvwOCInternacionales 
         Height          =   2535
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwOCNacionales 
         Height          =   2535
         Left            =   240
         TabIndex        =   34
         Top             =   3600
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwOCIndirectas 
         Height          =   135
         Left            =   4320
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   238
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label13 
         Caption         =   "Indirectas :"
         Height          =   255
         Left            =   3480
         TabIndex        =   38
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Nacionales :"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Nacionales 
         Caption         =   "Internacionales :"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   12240
      TabIndex        =   31
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   12240
      TabIndex        =   30
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   12120
      TabIndex        =   28
      Top             =   240
      Visible         =   0   'False
      Width           =   975
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmPagosOrdenes.frx":20C4
         MousePointer    =   99  'Custom
         Picture         =   "FrmPagosOrdenes.frx":23CE
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   12120
      TabIndex        =   26
      Top             =   4200
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmPagosOrdenes.frx":3D90
         MousePointer    =   99  'Custom
         Picture         =   "FrmPagosOrdenes.frx":409A
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   12120
      TabIndex        =   24
      Top             =   5400
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmPagosOrdenes.frx":5A5C
         MousePointer    =   99  'Custom
         Picture         =   "FrmPagosOrdenes.frx":5D66
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
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   195
      Left            =   11160
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   8640
      TabIndex        =   11
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmPagosOrdenes.frx":7E48
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbldeuda"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblFolio"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "textsalpago"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtTotal"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "opnNacional"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "opnInternacional"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "opnIndirecta"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtSaldo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "0"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton opnIndirecta 
         Caption         =   "Indirecta"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton opnInternacional 
         Caption         =   "Internacional"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton opnNacional 
         Caption         =   "Nacional"
         Enabled         =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   0
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "0"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Height          =   3135
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   3135
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   1080
            Width           =   2895
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox txtTrans 
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   2280
            Width           =   2895
         End
         Begin VB.TextBox txtCheque 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   1680
            Width           =   2895
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Fact.Prove"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   13
            Top             =   2760
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label Label4 
            Caption         =   "BANCO"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "TIPO DE PAGO"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "NUMERO DE TRANSFERENCIA"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   2040
            Width           =   2415
         End
         Begin VB.Label Label7 
            Caption         =   "NUMERO DE CHEQUE"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1440
            Width           =   2055
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   195
         Left            =   3000
         TabIndex        =   5
         Top             =   2520
         Width           =   255
      End
      Begin VB.TextBox textsalpago 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "SALDO"
         Height          =   255
         Left            =   960
         TabIndex        =   41
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblFolio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label9 
         Caption         =   "FOLIO"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL"
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblID 
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lbldeuda 
         Caption         =   "PAGO A REALIZAR"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12240
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
End
Attribute VB_Name = "FrmPagosOrdenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim TipoOrden As String
Dim sPendiente As String
Dim tip As String
Dim TotPagar As String
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Command1_Click()
    FRORPEN.Show vbModal
End Sub
Private Sub Command2_Click()
    frmfactpro.Show vbModal
End Sub
Private Sub Command3_Click()
    textsalpago.Enabled = True
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
    With lvwOCInternacionales
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "FOLIO", 800
        .ColumnHeaders.Add , , "ID_ORDEN", 0
        .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
        .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
        .ColumnHeaders.Add , , "COMENTARIO", 500
        .ColumnHeaders.Add , , "DEUDA PENDIENTE", 1440
    End With
    With lvwOCNacionales
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "FOLIO", 800
        .ColumnHeaders.Add , , "ID_ORDEN", 0
        .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
        .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
        .ColumnHeaders.Add , , "COMENTARIO", 500
         .ColumnHeaders.Add , , "DEUDA PENDIENTE", 1440
    End With
    With lvwOCIndirectas
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "FOLIO", 800
        .ColumnHeaders.Add , , "ID_ORDEN", 0
        .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
        .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
        .ColumnHeaders.Add , , "COMENTARIO", 1440
        .ColumnHeaders.Add , , "DEUDA PENDIENTE", 1440
    End With
    With lvwOrdenesNSurtidas
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "PROVEEDOR", 1440
        .ColumnHeaders.Add , , "NUM.ORDEN", 1440
        .ColumnHeaders.Add , , "FECHA", 1000
        .ColumnHeaders.Add , , "IMPORTE", 1440
    End With
    If Hay_Ordenes_Compra Then
        Llenar_Lista_Compras "Internacionales"
        Llenar_Lista_Compras "Nacionales"
        Llenar_Lista_Compras "Indirectas"
    End If
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    sBuscar = "SELECT * FROM BANCOS"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Combo2.Clear
            Do While Not .EOF
                Combo2.AddItem (.Fields("NOMBRE"))
                .MoveNext
            Loop
        Else
            MsgBox "NO EXISTEN BANCOS REGISTRADOS, NO PUEDE REGISTRAR PAGOS", vbInformation, "SACC"
        End If
        .Close
    End With
    sBuscar = "SELECT * FROM TPAGOS_OC"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Combo1.Clear
            Do While Not .EOF
                Combo1.AddItem (.Fields("Descripcion"))
                .MoveNext
            Loop
        Else
            MsgBox "FALLO DE INFORMACION, FAVOR DE LLAMAR A SOPORTE", vbInformation, "SACC"
        End If
        .Close
    End With
    sBuscar = "SELECT * FROM ABONOS_PAGO_OC WHERE CANT_ABONO = 0"
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Do While Not tRs1.EOF
            If tRs1.Fields("TIPO") = "R" Then
                Tipo = "RAPIDA"
            End If
            If tRs1.Fields("TIPO") = "N" Then
                Tipo = "NACIONAL"
            End If
            If tRs1.Fields("TIPO") = "I" Then
                Tipo = "INTERNACIONAL"
            End If
            If tRs1.Fields("TIPO") = "X" Then
                Tipo = "INDIRECTA"
            End If
            If Tipo <> "" Then
                sBuscar = "SELECT * FROM CHEQUES WHERE NUM_ORDEN LIKE '%" & tRs1.Fields("NUM_ORDEN") & ",%' AND TIPO_ORDEN = '" & Tipo & "' AND FECHA_REALIZADO = '" & tRs1.Fields("FECHA") & "'"
            Else
                sBuscar = "SELECT * FROM CHEQUES WHERE NUM_ORDEN LIKE '%" & tRs1.Fields("NUM_ORDEN") & ",%' AND FECHA_REALIZADO = '" & tRs1.Fields("FECHA") & "'"
            End If
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                sBuscar = "UPDATE ABONOS_PAGO_OC SET NO_CHEQUE = '" & tRs.Fields("NUM_CHEQUE") & "', BANCO = '" & tRs.Fields("BANCO") & "', CANTIDAD = '" & Replace(tRs.Fields("TOTAL"), "$", "") & "', CANT_ABONO = '" & Replace(Replace(tRs.Fields("TOTAL"), "$", ""), ",", "") & "', NUMCHEQUE = '" & tRs.Fields("NUM_CHEQUE") & "' WHERE NUM_ORDEN IN (" & Mid(tRs.Fields("NUM_ORDEN"), 1, Len(tRs.Fields("NUM_ORDEN")) - 2) & ") AND CANT_ABONO = 0"
                cnn.Execute (sBuscar)
            End If
            tRs1.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Public Function Hay_Ordenes_Compra() As Boolean
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT  count(*) as Orden_Compra From ORDEN_COMPRA WHERE Confirmada = 'X'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If .Fields("ORDEN_COMPRA") <> 0 Then
            Hay_Ordenes_Compra = True
        Else
            Hay_Ordenes_Compra = False
        End If
        .Close
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Lista_Compras(Tipo As String)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim CompDolar As Double
    Dim NumOrden As Integer
    Dim tRs2 As ADODB.Recordset
    Dim sBusca As String
    'falta agregar el tipo de OC al update para evistar duplicidad
    sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN IN (SELECT ORDEN_COMPRA.NUM_ORDEN FROM ORDEN_COMPRA, ABONOS_PAGO_OC WHERE ORDEN_COMPRA.NUM_ORDEN = ABONOS_PAGO_OC.NUM_ORDEN AND ORDEN_COMPRA.TIPO = ABONOS_PAGO_OC.TIPO AND ORDEN_COMPRA.CONFIRMADA = 'X' AND ABONOS_PAGO_OC.CANT_ABONO >= (ORDEN_COMPRA.TOTAL + ORDEN_COMPRA.TAX + ORDEN_COMPRA.OTROS_CARGOS + ORDEN_COMPRA.FREIGHT - ORDEN_COMPRA.DISCOUNT) AND ORDEN_COMPRA.MONEDA = 'PESOS') AND CONFIRMADA = 'X' AND MONEDA = 'PESOS'"
    cnn.Execute (sBuscar)
    sBuscar = "SELECT COMPRA FROM DOLAR WHERE FECHA = '" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        If Not IsNull(tRs.Fields("COMPRA")) Then
            CompDolar = tRs.Fields("COMPRA")
        Else
            CompDolar = InputBox("DE EL PRECIO DE VENTA DEL DOLAR HOY!")
            sBuscar = "INSERT INTO DOLAR (FECHA, COMPRA, VENTA) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', " & CompDolar & ", " & InputBox("CON FIN DE ACTUALIZAR EL TIPO DE CAMBIO A LA FECHA, DE EL PRECIO DE COMPRA DEL DOLAR HOY!") & ");"
            cnn.Execute (sBuscar)
        End If
    Else
        CompDolar = InputBox("DE EL PRECIO DE VENTA DEL DOLAR HOY!")
        sBuscar = "INSERT INTO DOLAR (FECHA, COMPRA, VENTA) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', " & CompDolar & ", " & InputBox("CON FIN DE ACTUALIZAR EL TIPO DE CAMBIO A LA FECHA, DE EL PRECIO DE COMPRA DEL DOLAR HOY!") & ");"
        cnn.Execute (sBuscar)
    End If
    sBuscar = "SELECT OC.Id_Orden_Compra, OC.NUM_ORDEN, OC.Id_Proveedor, P.Nombre, ((OC.Total - OC.Discount) + OC.TAX + OC.Freight + OC.Otros_Cargos) AS Total_Pagar, OC.COMENTARIO, MONEDA FROM ORDEN_COMPRA AS OC JOIN PROVEEDOR AS P ON P.Id_Proveedor = OC.Id_Proveedor WHERE OC.Confirmada = 'X' AND OC.Tipo = '"
    Select Case Tipo
        Case "Internacionales":
            Me.lvwOCInternacionales.ListItems.Clear
            sBuscar = sBuscar & "I' ORDER BY NUM_ORDEN"
            Set tRs = cnn.Execute(sBuscar)
            With tRs
                While Not .EOF
                    Set ItMx = Me.lvwOCInternacionales.ListItems.Add(, , .Fields("NUM_ORDEN"))
                    If Not IsNull(.Fields("ID_ORDEN_COMPRA")) Then ItMx.SubItems(1) = Trim(.Fields("ID_ORDEN_COMPRA"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(2) = Trim(.Fields("ID_PROVEEDOR"))
                    'modificacion line de abajo
                    If Not IsNull(.Fields("NUM_ORDEN")) Then NumOrden = .Fields("NUM_ORDEN")
                    If Not IsNull(.Fields("NOMBRE")) Then ItMx.SubItems(3) = Trim(.Fields("NOMBRE"))
                    If .Fields("MONEDA") = "DOLARES" Then
                        If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Format(Trim(Format(CDbl(.Fields("Total_Pagar")) * CDbl(CompDolar), "###,###,##0.00")), "###,###,##0.00")
                    Else
                        If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(Format(CDbl(.Fields("Total_Pagar")), "###,###,##0.00"))
                    End If
                    If Not IsNull(.Fields("COMENTARIO")) Then ItMx.SubItems(5) = Trim(.Fields("COMENTARIO"))
                    .MoveNext
                Wend
            .Close
            End With
        Case "Nacionales":
            Me.lvwOCNacionales.ListItems.Clear
            sBuscar = sBuscar & "N' ORDER BY NUM_ORDEN"
            Set tRs = cnn.Execute(sBuscar)
            With tRs
                While Not .EOF
                    Set ItMx = Me.lvwOCNacionales.ListItems.Add(, , .Fields("NUM_ORDEN"))
                    If Not IsNull(.Fields("ID_ORDEN_COMPRA")) Then ItMx.SubItems(1) = Trim(.Fields("ID_ORDEN_COMPRA"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(2) = Trim(.Fields("ID_PROVEEDOR"))
                    If Not IsNull(.Fields("Nombre")) Then ItMx.SubItems(3) = Trim(.Fields("Nombre"))
                    If .Fields("MONEDA") = "DOLARES" Then
                        If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(Format(CDbl(.Fields("Total_Pagar")) * CDbl(CompDolar), "###,###,##0.00"))
                    Else
                        If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(Format(CDbl(.Fields("Total_Pagar")), "###,###,##0.00"))
                    End If
                    If Not IsNull(.Fields("COMENTARIO")) Then ItMx.SubItems(5) = Trim(.Fields("COMENTARIO"))
                    .MoveNext
                Wend
            .Close
            End With
        Case "Indirectas":
            Me.lvwOCIndirectas.ListItems.Clear
            sBuscar = sBuscar & "X' ORDER BY NUM_ORDEN"
            Set tRs = cnn.Execute(sBuscar)
            With tRs
                While Not .EOF
                    Set ItMx = Me.lvwOCIndirectas.ListItems.Add(, , .Fields("NUM_ORDEN"))
                    If Not IsNull(.Fields("ID_ORDEN_COMPRA")) Then ItMx.SubItems(1) = Trim(.Fields("ID_ORDEN_COMPRA"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(2) = Trim(.Fields("ID_PROVEEDOR"))
                    If Not IsNull(.Fields("Nombre")) Then ItMx.SubItems(3) = Trim(.Fields("Nombre"))
                    If .Fields("MONEDA") = "DOLARES" Then
                        If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(Format(CDbl(.Fields("Total_Pagar")) * CDbl(CompDolar), "###,###,##0.00"))
                    Else
                        If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(Format(CDbl(.Fields("Total_Pagar")), "###,###,##0.00"))
                    End If
                    If Not IsNull(.Fields("COMENTARIO")) Then ItMx.SubItems(5) = Trim(.Fields("COMENTARIO"))
                    .MoveNext
                Wend
            .Close
            End With
        Case Else:
            MsgBox "ERROR GRAVE. LA APLICACIÓN TERMINARA", vbCritical, "SACC"
            End
    End Select
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image1_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim sResta As String
    Dim sPago As String
    Dim subto As Double
    Dim Subtotal As Double
    subto = CDbl(sPendiente) - CDbl(textsalpago)
    Subtotal = CDbl(textsalpago.Text) - CDbl(textsalpago.Text)
    If IsNumeric(textsalpago.Text) Then
        sResta = textsalpago.Text
    Else
        sResta = 0
        MsgBox "Debe ingresar una cantidad para abonar", vbExclamation, "SACC"
        Exit Sub
    End If
    If MsgBox("ESTA POR REGISTRARA UN PAGO A LA ORDEN DE COMPRA SELECCIONADA,¿ESTA SEGURO QUE DESEA CONTINUAR?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
        FrmCheque.TxtNUM_ORDEN.Text = lblFolio.Caption 'numero de orden de compra
        FrmCheque.TxtTIPO_ORDEN.Text = TipoOrden 'tipo de orden de compra
        FrmCheque.txtNum2Let(0).Text = textsalpago
        FrmCheque.TxtNOMBRE.Text = Label5.Caption 'nombre del proveedor a recibir el pago
        FrmCheque.TxtNUM_CHEQUE.Text = txtCheque.Text 'numero de cheque
        FrmCheque.Combo1.Text = Combo2.Text 'banco
        FrmCheque.Show vbModal
        If subto = 0 Or Subtotal = 0 Then
            If TipoOrden = "NACIONAL" Then
                For Con = 1 To lvwOCNacionales.ListItems.Count
                    If lvwOCNacionales.ListItems(Con).Checked Then
                       If subto = 0 Or Subtotal = 0 Then
                            ' aqui esta cerrando la orden mas no checa si se debe cerrar
                            ' debe tener un ciclo que reste el total de cada orden al abono dado...
                            ' y solo si el importe restante
                            sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN IN (" & Mid(lblFolio.Caption, 1, Len(lblFolio.Caption) - 2) & ") AND TIPO='" & tip & "' "
                            cnn.Execute (sBuscar)
                       End If
                    End If
                Next Con
            End If
            If TipoOrden = "INTERNACIONAL" Then
                For Con = 1 To lvwOCInternacionales.ListItems.Count
                    If lvwOCInternacionales.ListItems(Con).Checked Then
                        If subto = 0 Or Subtotal = 0 Then
                            sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN  IN (" & Mid(lblFolio.Caption, 1, Len(lblFolio.Caption) - 2) & "') AND TIPO='" & tip & "' "
                            cnn.Execute (sBuscar)
                        End If
                    End If
                Next Con
            End If
             If TipoOrden = "INDIRECTA" Then
                For Con = 1 To lvwOCIndirectas.ListItems.Count
                    If lvwOCIndirectas.ListItems(Con).Checked Then
                        If subto = 0 Or Subtotal = 0 Then
                            sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN IN (" & Mid(lblFolio.Caption, 1, Len(lblFolio.Caption) - 2) & ") AND TIPO='" & tip & "' "
                            cnn.Execute (sBuscar)
                        End If
                    End If
                Next Con
            End If
        End If
        If textsalpago <> txtTotal.Text Then
            If lblID.Caption <> "" Then
                If TipoOrden = "INTERNACIONAL" Then
                    For Con = 1 To lvwOCInternacionales.ListItems.Count
                    ' se debe determinar el importe por cadaorden a pagar para
                    'restar solo el abono de esa orden y saber que cantidad queda pendiente por abonar
                        sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN = " & lvwOCInternacionales.ListItems(Con) & " AND TIPO='I'"
                        Set tRs = cnn.Execute(sBusca)
                        If CDbl(sResta) > CDbl(tRs.Fields("CANT_ABONO")) Then
                            sResta = CDbl(sResta) - CDbl(tRs.Fields("CANT_ABONO"))
                            sPago = tRs.Fields("CANT_ABONO")
                        Else
                            sPago = sResta
                            sResta = "0.00"
                        End If
                        If lvwOCInternacionales.ListItems(Con).Checked Then
                            sBuscar = "INSERT INTO ABONOS_PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANT_ABONO, FECHA, TIPO, NUM_ORDEN, ID_PROVEEDOR) VALUES (" & lvwOCInternacionales.ListItems(Con).SubItems(1) & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "','I','" & lvwOCInternacionales.ListItems(Con) & "', '" & lvwOCInternacionales.ListItems(Con).SubItems(2) & "');"
                            cnn.Execute (sBuscar)
                        End If
                    Next Con
             End If
                If TipoOrden = "NACIONAL" Then
                    For Con = 1 To lvwOCNacionales.ListItems.Count
                        sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN = " & lvwOCNacionales.ListItems(Con) & " AND TIPO='N'"
                        Set tRs = cnn.Execute(sBusca)
                        If CDbl(sResta) > CDbl(tRs.Fields("CANT_ABONO")) Then
                            sResta = CDbl(sResta) - CDbl(tRs.Fields("CANT_ABONO"))
                            sPago = tRs.Fields("CANT_ABONO")
                        Else
                            sPago = sResta
                            sResta = "0.00"
                        End If
                        If lvwOCNacionales.ListItems(Con).Checked Then
                            sBuscar = "INSERT INTO ABONOS_PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANT_ABONO, FECHA, TIPO, NUM_ORDEN, ID_PROVEEDOR) VALUES (" & lvwOCNacionales.ListItems(Con).SubItems(1) & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "','I','" & lvwOCNacionales.ListItems(Con) & "','" & lvwOCNacionales.ListItems(Con).SubItems(2) & "');"
                            cnn.Execute (sBuscar)
                        End If
                    Next Con
                End If
                If TipoOrden = "INDIRECTA" Then
                    For Con = 1 To lvwOCIndirectas.ListItems.Count
                        sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN = " & lvwOCIndirectas.ListItems(Con) & " AND TIPO='X'"
                        Set tRs = cnn.Execute(sBusca)
                        If CDbl(sResta) > CDbl(tRs.Fields("CANT_ABONO")) Then
                            sResta = CDbl(sResta) - CDbl(tRs.Fields("CANT_ABONO"))
                            sPago = tRs.Fields("CANT_ABONO")
                        Else
                            sPago = sResta
                            sResta = "0.00"
                        End If
                        If lvwOCIndirectas.ListItems(Con).Checked Then
                            sBuscar = "INSERT INTO ABONOS_PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANT_ABONO, FECHA, TIPO, NUM_ORDEN, ID_PROVEEDOR) VALUES (" & lvwOCIndirectas.ListItems(Con).SubItems(1) & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "', 'I', '" & lvwOCIndirectas.ListItems(Con) & "', '" & lvwOCIndirectas.ListItems(Con).SubItems(2) & "');"
                            cnn.Execute (sBuscar)
                        End If
                    Next Con
                End If
            End If
            ' si el pago escrito es igual al pendiente entonces lo marca como pagado
            ' insert que generaba el pago por el total de la factura (sin parecialidades)
            ' insert nuevo que permite el pago en parcialidades dado en "textsalpago"
            sBuscar = "INSERT INTO PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANTIDAD, FECHA) VALUES (" & lblID.Caption & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "');"
            cnn.Execute (sBuscar)
            'ImprimePago
            TipoOrden = ""
        Else
            MsgBox "NO PUEDE REGISTRAR PAGOS SIN LA INFORMACION COMPLETA", vbInformation, "SACC"
        End If
    Else
        MsgBox "DEBE SELECCIONAR UNA ORDEN DE COMPRA A PAGAR", vbInformation, "SACC"
    End If
End Sub
Private Sub Image3_Click()
    FRORPEN.Show vbModal
End Sub
Private Sub Image8_Click()
    If Not ArchivoEnUso(App.Path & "\Cheque.pdf") Then
        If Combo1.Text <> "" And Combo2.Text <> "" And (txtCheque.Text <> "" Or txtTrans.Text <> "") And textsalpago.Text <> "" Then
            Dim sBuscar As String
            Dim tRs As ADODB.Recordset
            Dim sResta As String
            Dim sPago As String
            Dim subto As Double
            Dim SaldoResta As Double
            SaldoResta = CDbl(textsalpago.Text)
            Dim Subtotal As Double
            Dim IdProv As String
            subto = CDbl(sPendiente) - CDbl(textsalpago)
            Subtotal = CDbl(txtTotal.Text) - CDbl(textsalpago.Text)
            If IsNumeric(textsalpago.Text) Then
                sResta = textsalpago.Text
            Else
                sResta = 0
                MsgBox "Debe ingresar una cantidad para abonar", vbExclamation, "SACC"
                Exit Sub
            End If
            If MsgBox("ESTA POR REGISTRARA UN PAGO A LA ORDEN DE COMPRA SELECCIONADA,¿ESTA SEGURO QUE DESEA CONTINUAR?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
                If subto = 0 Or Subtotal = 0 Then
                     If TipoOrden = "NACIONAL" Then
                        For Con = 1 To lvwOCNacionales.ListItems.Count
                            If lvwOCNacionales.ListItems(Con).Checked Then
                                IdProv = lvwOCNacionales.ListItems(Con).SubItems(2)
                                If subto = 0 Or Subtotal = 0 Then
                                    If CDbl(lvwOCNacionales.ListItems(Con).SubItems(4)) >= SaldoResta Then
                                        ' aqui esta cerrando la orden mas no checa si se debe cerrar
                                        ' debe tener un ciclo que reste el total de cada orden al abono dado...
                                        ' y solo si el importe restante
                                        If ((CDbl(textsalpago.Text) + 0.009) >= CDbl(TotPagar)) Or (CDbl(textsalpago.Text) = 0) Then
                                            sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN= '" & lvwOCNacionales.ListItems(Con) & "' AND TIPO='" & tip & "' "
                                            cnn.Execute (sBuscar)
                                            SaldoResta = SaldoResta - CDbl(lvwOCNacionales.ListItems(Con).SubItems(4))
                                        End If
                                    End If
                                End If
                            End If
                        Next Con
                    End If
                    If TipoOrden = "INTERNACIONAL" Then
                        For Con = 1 To lvwOCInternacionales.ListItems.Count
                            If lvwOCInternacionales.ListItems(Con).Checked Then
                                IdProv = lvwOCInternacionales.ListItems(Con).SubItems(2)
                                If CDbl(lvwOCInternacionales.ListItems(Con).SubItems(4)) >= SaldoResta Then
                                    If subto = 0 Or Subtotal = 0 Then
                                        If ((CDbl(textsalpago.Text) + 0.009) >= CDbl(TotPagar)) Or (CDbl(textsalpago.Text) = 0) Then
                                            sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN =' " & lvwOCInternacionales.ListItems(Con) & "' AND TIPO='" & tip & "' "
                                            cnn.Execute (sBuscar)
                                            SaldoResta = SaldoResta - CDbl(lvwOCNacionales.ListItems(Con).SubItems(4))
                                        End If
                                    End If
                                End If
                            End If
                        Next Con
                    End If
                     If TipoOrden = "INDIRECTA" Then
                        For Con = 1 To lvwOCIndirectas.ListItems.Count
                            If lvwOCIndirectas.ListItems(Con).Checked Then
                                IdProv = lvwOCIndirectas.ListItems(Con).SubItems(2)
                                If CDbl(lvwOCInternacionales.ListItems(Con).SubItems(4)) >= SaldoResta Then
                                    If subto = 0 Or Subtotal = 0 Then
                                        If ((CDbl(textsalpago.Text) + 0.009) >= CDbl(TotPagar)) Or (CDbl(textsalpago.Text) = 0) Then
                                            sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN =' " & lvwOCIndirectas.ListItems(Con) & "' AND TIPO='" & tip & "' "
                                            cnn.Execute (sBuscar)
                                            SaldoResta = SaldoResta - CDbl(lvwOCNacionales.ListItems(Con).SubItems(4))
                                        End If
                                    End If
                                End If
                            End If
                        Next Con
                    End If
                End If
                '''''''''''''LO ANTERIOR  ES PARA CERRAR  LA ORDEN  EL IF  SIGUIENTE  ES CUANDO HACE  ABONOS
                'OSEA  QUE  IF  EXTSALPAGO<>TXTTOTAL  POR ESO  NO CERRABA  LA ORDEN
                If lblID.Caption <> "" Then
                    If TipoOrden = "INTERNACIONAL" Then
                        For Con = 1 To lvwOCInternacionales.ListItems.Count
                        ' se debe determinar el importe por cadaorden a pagar para
                        'restar solo el abono de esa orden y saber que cantidad queda pendiente por abonar
                            If lvwOCInternacionales.ListItems(Con).Checked Then
                                sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN = " & lvwOCInternacionales.ListItems(Con) & " AND TIPO='I'"
                                Set tRs = cnn.Execute(sBusca)
                                If Not (tRs.EOF And tRs.BOF) Then
                                    If Not IsNull(tRs.Fields("CANT_ABONO")) Then
                                        If CDbl(sResta) > CDbl(tRs.Fields("CANT_ABONO")) Then
                                            sResta = CDbl(sResta) - CDbl(tRs.Fields("CANT_ABONO"))
                                            sPago = tRs.Fields("CANT_ABONO")
                                        Else
                                            sPago = sResta
                                            sResta = "0.00"
                                            sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN= '" & lvwOCInternacionales.ListItems(Con) & "' AND TIPO='" & tip & "' "
                                            cnn.Execute (sBuscar)
                                        End If
                                    Else
                                        sPago = sResta
                                        sResta = "0.00"
                                        'sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN= '" & lvwOCInternacionales.ListItems(Con) & "' AND TIPO='" & tip & "' "
                                        'cnn.Execute (sBuscar)
                                    End If
                                Else
                                    sPago = sResta
                                    sResta = "0.00"
                                    'sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN= '" & lvwOCInternacionales.ListItems(Con) & "' AND TIPO='" & tip & "' "
                                    'cnn.Execute (sBuscar)
                                End If
                                If lvwOCInternacionales.ListItems(Con).Checked Then
                                    sBuscar = "INSERT INTO ABONOS_PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANT_ABONO, FECHA, TIPO, NUM_ORDEN, ID_PROVEEDOR) VALUES (" & lvwOCInternacionales.ListItems(Con).SubItems(1) & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "','I','" & lvwOCInternacionales.ListItems(Con) & "','" & lvwOCInternacionales.ListItems(Con).SubItems(2) & "');"
                                    cnn.Execute (sBuscar)
                                    sBuscar = "INSERT INTO PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANTIDAD, FECHA) VALUES (" & lvwOCInternacionales.ListItems(Con).SubItems(1) & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "');"
                                    cnn.Execute (sBuscar)
                                End If
                            End If
                        Next Con
                    End If
                    If TipoOrden = "NACIONAL" Then
                        For Con = 1 To lvwOCNacionales.ListItems.Count
                            If lvwOCNacionales.ListItems(Con).Checked Then
                                sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN = " & lvwOCNacionales.ListItems(Con) & " AND TIPO='N'"
                                Set tRs = cnn.Execute(sBusca)
                                If Not IsNull(tRs.Fields("CANT_ABONO")) Then
                                    If CDbl(sResta) > CDbl(tRs.Fields("CANT_ABONO")) Then
                                        sResta = CDbl(sResta) - CDbl(tRs.Fields("CANT_ABONO"))
                                        sPago = tRs.Fields("CANT_ABONO")
                                    Else
                                        sPago = sResta
                                        sResta = "0.00"
                                        'sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN= '" & lvwOCNacionales.ListItems(Con) & "' AND TIPO='" & tip & "' "
                                        'cnn.Execute (sBuscar)
                                    End If
                                Else
                                    sPago = sResta
                                    sResta = "0.00"
                                    'sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN= '" & lvwOCNacionales.ListItems(Con) & "' AND TIPO='" & tip & "' "
                                    'cnn.Execute (sBuscar)
                                End If
                                If lvwOCNacionales.ListItems(Con).Checked Then
                                    sBuscar = "INSERT INTO ABONOS_PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANT_ABONO, FECHA, TIPO, NUM_ORDEN, ID_PROVEEDOR) VALUES (" & lvwOCNacionales.ListItems(Con).SubItems(1) & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "','N','" & lvwOCNacionales.ListItems(Con) & "', '" & lvwOCNacionales.ListItems(Con).SubItems(2) & "');"
                                    cnn.Execute (sBuscar)
                                    sBuscar = "INSERT INTO PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANTIDAD, FECHA) VALUES (" & lvwOCNacionales.ListItems(Con).SubItems(1) & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "');"
                                    cnn.Execute (sBuscar)
                                End If
                            End If
                        Next Con
                    End If
                    If TipoOrden = "INDIRECTA" Then
                        For Con = 1 To lvwOCIndirectas.ListItems.Count
                            If lvwOCIndirectas.ListItems(Con).Checked Then
                                sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN = " & lvwOCIndirectas.ListItems(Con) & " AND TIPO='X'"
                                Set tRs = cnn.Execute(sBusca)
                                If CDbl(sResta) > CDbl(tRs.Fields("CANT_ABONO")) Then
                                    sResta = CDbl(sResta) - CDbl(tRs.Fields("CANT_ABONO"))
                                    sPago = tRs.Fields("CANT_ABONO")
                                Else
                                    sPago = sResta
                                    sResta = "0.00"
                                    'sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN= '" & lvwOCIndirectas.ListItems(Con) & "' AND TIPO='" & tip & "' "
                                    'cnn.Execute (sBuscar)
                                End If
                                If lvwOCIndirectas.ListItems(Con).Checked Then
                                    sBuscar = "INSERT INTO ABONOS_PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANT_ABONO, FECHA, TIPO, NUM_ORDEN, ID_PROVEEDOR) VALUES (" & lvwOCIndirectas.ListItems(Con).SubItems(1) & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "', 'X', '" & lvwOCIndirectas.ListItems(Con) & "', '" & lvwOCIndirectas.ListItems(Con).SubItems(2) & "');"
                                    cnn.Execute (sBuscar)
                                    sBuscar = "INSERT INTO PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANTIDAD, FECHA) VALUES (" & lvwOCIndirectas.ListItems(Con).SubItems(1) & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "');"
                                    cnn.Execute (sBuscar)
                                End If
                            End If
                        Next Con
                    End If
                End If
                'ImprimePago
                FrmCheque.TxtNUM_ORDEN.Text = lblFolio.Caption 'numero de orden de compra
                FrmCheque.TxtTIPO_ORDEN.Text = TipoOrden 'tipo de orden de compra
                FrmCheque.txtNum2Let(0).Text = textsalpago
                FrmCheque.TxtNOMBRE.Text = Label5.Caption 'nombre del proveedor a recibir el pago
                FrmCheque.TxtIdProv.Text = IdProv
                If txtTrans.Text = "" Then
                    FrmCheque.TxtNUM_CHEQUE.Text = txtCheque.Text 'numero de cheque
                Else
                    FrmCheque.TxtNUM_CHEQUE.Text = txtTrans.Text
                End If
                FrmCheque.Combo1.Text = Combo2.Text 'banco
                FrmCheque.Show vbModal
                TipoOrden = ""
                If Hay_Ordenes_Compra Then
                    Llenar_Lista_Compras "Internacionales"
                    Llenar_Lista_Compras "Nacionales"
                    Llenar_Lista_Compras "Indirectas"
                End If
                'ASEGURA EL CIERRE DE ORDENES QUE YA FUERON PAGADAS
                sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE NUM_ORDEN IN (SELECT NUM_ORDEN FROM ABONOS_PAGO_OC WHERE NUM_ORDEN IN (SELECT NUM_ORDEN FROM ORDEN_COMPRA WHERE TIPO = 'N') AND TIPO='N'  AND CANT_ABONO = (SELECT (TOTAL + FREIGHT + TAX + OTROS_CARGOS - DISCOUNT) AS TOTAL FROM ORDEN_COMPRA WHERE NUM_ORDEN = ABONOS_PAGO_OC.NUM_ORDEN AND TIPO = 'N' AND CONFIRMADA = 'X' ) GROUP BY NUM_ORDEN) AND TIPO = 'N'"
                cnn.Execute (sBuscar)
            Else
                MsgBox "DEBE SELECCIONAR UNA ORDEN DE COMPRA A PAGAR", vbInformation, "SACC"
            End If
        Else
            MsgBox "Debe dar la información del pago antes de continuar", vbExclamation, "SACC"
        End If
    Else
        MsgBox "Cierre la poliza de cheque que tiene abierta para poder continuar", vbExclamation, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwOCIndirectas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim Con As Integer
    Dim IdProve As String
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    txtTotal.Text = "0"
    textsalpago.Text = "0"
    lblFolio.Caption = ""
    Label5.Caption = ""
    lblID.Caption = ""
    TipoOrden = "NACIONAL"
    tip = "N"
    opnIndirecta.Value = True
    For Con = 1 To lvwOCInternacionales.ListItems.Count
        lvwOCInternacionales.ListItems(Con).Checked = False
    Next
    For Con = 1 To lvwOCNacionales.ListItems.Count
        lvwOCNacionales.ListItems(Con).Checked = False
    Next
    For Con = 1 To lvwOCIndirectas.ListItems.Count
        If lvwOCIndirectas.ListItems(Con).Checked Then
            If IdProve = "" Then
                IdProve = lvwOCIndirectas.ListItems(Con).SubItems(2)
                sBusca = "SELECT ORDEN_COMPRA.ID_PROVEEDOR AS PROVEE, ORDEN_COMPRA.FECHA AS FECHA, " & _
                         "       ORDENES_NO_SURTIDAS.NUM_ORDEN AS NUMORDEN, (ORDENES_NO_SURTIDAS.PRECIO * ORDENES_NO_SURTIDAS.CANTIDAD) AS IMPORTE  " & _
                         "FROM ORDENES_NO_SURTIDAS INNER JOIN " & _
                         "     ORDEN_COMPRA ON ORDENES_NO_SURTIDAS.NUM_ORDEN = ORDEN_COMPRA.NUM_ORDEN AND " & _
                         "     ORDENES_NO_SURTIDAS.Tipo = ORDEN_COMPRA.Tipo " & _
                         "WHERE ORDEN_COMPRA.ID_PROVEEDOR = " & IdProve
                Set tRs = cnn.Execute(sBusca)
                Me.lvwOrdenesNSurtidas.ListItems.Clear
                txtSaldo.Text = "0"
                With tRs
                    While Not .EOF
                        Set tLi = lvwOrdenesNSurtidas.ListItems.Add(, , .Fields("PROVEE"))
                        If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(1) = .Fields("FECHA")
                        If Not IsNull(.Fields("NUMORDEN")) Then tLi.SubItems(2) = .Fields("NUMORDEN")
                        If Not IsNull(.Fields("IMPORTE")) Then
                            tLi.SubItems(2) = .Fields("IMPORTE")
                            txtSaldo.Text = CDbl(txtSaldo.Text) + CDbl(.Fields("IMPORTE"))
                        End If
                        .MoveNext
                    Wend
                    .Close
                End With
            End If
            If IdProve = lvwOCIndirectas.ListItems(Con).SubItems(2) Then
                lblFolio.Caption = lblFolio.Caption & lvwOCIndirectas.ListItems(Con) & ", "
                Label5.Caption = lvwOCIndirectas.ListItems(Con).SubItems(3)
                txtTotal.Text = CDbl(txtTotal.Text) + CDbl(lvwOCIndirectas.ListItems(Con).SubItems(4))
                sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN IN (" & Mid(lblFolio.Caption, 1, Len(lblFolio.Caption) - 2) & ") AND TIPO='X'"
                Set tRs = cnn.Execute(sBusca)
                If Not (tRs.EOF And tRs.BOF) Then
                    If Not IsNull(tRs.Fields("CANT_ABONO")) Then
                        textsalpago.Text = CDbl(txtTotal.Text) - CDbl(tRs.Fields("CANT_ABONO"))
                        TotPagar = textsalpago.Text
                    Else
                        textsalpago.Text = txtTotal.Text
                        TotPagar = textsalpago.Text
                    End If
                End If
                textsalpago.Text = CDbl(textsalpago.Text) - CDbl(txtSaldo.Text)
                sPendiente = textsalpago.Text
                lblID.Caption = lblID.Caption & lvwOCIndirectas.ListItems(Con).SubItems(1) & ", "
            Else
                lvwOCIndirectas.ListItems(Con).Checked = False
                MsgBox "TODAS LAS ORDENES DEBEN SER DEL MISMO PROVEEDOR !!", vbExclamation, "SACC"
            End If
        End If
    Next
End Sub
Private Sub lvwOCInternacionales_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim Con As Integer
    Dim IdProve As String
    Dim sBuscar As String
    Dim ConSel As Integer
    Dim tRs As ADODB.Recordset
    ConSel = 0
    txtTotal.Text = "0"
    textsalpago.Text = "0"
    lblFolio.Caption = ""
    Label5.Caption = ""
    lblID.Caption = ""
    TipoOrden = "INTERNACIONAL"
    tip = "I"
    opnInternacional.Value = True
    For Con = 1 To lvwOCNacionales.ListItems.Count
        lvwOCNacionales.ListItems(Con).Checked = False
    Next
    For Con = 1 To lvwOCIndirectas.ListItems.Count
        lvwOCIndirectas.ListItems(Con).Checked = False
    Next
    For Con = 1 To lvwOCInternacionales.ListItems.Count
        If lvwOCInternacionales.ListItems(Con).Checked Then
            If IdProve = "" Then
                IdProve = lvwOCInternacionales.ListItems(Con).SubItems(2)
               'cv
                sBusca = "SELECT ORDEN_COMPRA.ID_PROVEEDOR AS PROVEE, ORDEN_COMPRA.FECHA AS FECHA, " & _
                         "       ORDENES_NO_SURTIDAS.NUM_ORDEN AS NUMORDEN, (ORDENES_NO_SURTIDAS.PRECIO * ORDENES_NO_SURTIDAS.CANTIDAD) AS IMPORTE  " & _
                         "FROM ORDENES_NO_SURTIDAS INNER JOIN " & _
                         "     ORDEN_COMPRA ON ORDENES_NO_SURTIDAS.NUM_ORDEN = ORDEN_COMPRA.NUM_ORDEN AND " & _
                         "     ORDENES_NO_SURTIDAS.Tipo = ORDEN_COMPRA.Tipo " & _
                         "WHERE ORDEN_COMPRA.ID_PROVEEDOR = " & IdProve
                Set tRs = cnn.Execute(sBusca)
                Me.lvwOrdenesNSurtidas.ListItems.Clear
                txtSaldo.Text = "0"
                With tRs
                    While Not .EOF
                        Set tLi = lvwOrdenesNSurtidas.ListItems.Add(, , .Fields("PROVEE"))
                        If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(1) = .Fields("FECHA")
                        If Not IsNull(.Fields("NUMORDEN")) Then tLi.SubItems(2) = .Fields("NUMORDEN")
                        If Not IsNull(.Fields("IMPORTE")) Then
                            tLi.SubItems(3) = .Fields("IMPORTE")
                            txtSaldo.Text = CDbl(txtSaldo.Text) + CDbl(.Fields("IMPORTE"))
                        End If
                        .MoveNext
                    Wend
                    .Close
                End With
            End If
            If IdProve = lvwOCInternacionales.ListItems(Con).SubItems(2) Then
                lblFolio.Caption = lblFolio.Caption & lvwOCInternacionales.ListItems(Con) & ", "
                Label5.Caption = lvwOCInternacionales.ListItems(Con).SubItems(3)
                txtTotal.Text = CDbl(txtTotal.Text) + CDbl(lvwOCInternacionales.ListItems(Con).SubItems(4))
                sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN IN (" & Mid(lblFolio.Caption, 1, Len(lblFolio.Caption) - 2) & ") AND TIPO='I'"
                Set tRs = cnn.Execute(sBusca)
                If Not (tRs.EOF And tRs.BOF) Then
                    If Not IsNull(tRs.Fields("CANT_ABONO")) Then
                        textsalpago.Text = CDbl(txtTotal.Text) - CDbl(tRs.Fields("CANT_ABONO"))
                        TotPagar = textsalpago.Text
                    Else
                        textsalpago.Text = txtTotal.Text
                        TotPagar = textsalpago.Text
                    End If
                End If
                textsalpago.Text = CDbl(textsalpago.Text) - CDbl(txtSaldo.Text)
                sPendiente = textsalpago.Text
                lblID.Caption = lblID.Caption & lvwOCInternacionales.ListItems(Con).SubItems(1) & ", "
            Else
                lvwOCInternacionales.ListItems(Con).Checked = False
                MsgBox "TODAS LAS ORDENES DEBEN SER DEL MISMO PROVEEDOR !!", vbExclamation, "SACC"
            End If
            ConSel = ConSel + 1
            If CDbl(ConSel) > 1 Then
                Command3.Enabled = False
                textsalpago.Enabled = False
            Else
                Command3.Enabled = True
            End If
        End If
    Next
End Sub
Private Sub ImprimePago()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim POSY As Integer
    POSY = 4400
    sBuscar = "SELECT * FROM VsPagoOrden WHERE ID_ORDEN_COMPRA IN (" & Mid(lblID.Caption, 1, Len(lblID.Caption) - 2) & ")"
    Set tRs = cnn.Execute(sBuscar)
    CommonDialog1.Flags = 64
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
    Printer.Print VarMen.Text5(0).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print ""
    ' DATOS OC
    Printer.Print "     Numero de Orden : " & tRs.Fields("NUM_ORDEN")
    Printer.Print "     Fecha de la OC : " & tRs.Fields("FECHA")
    Printer.Print "     Moneda : " & tRs.Fields("MONEDA")
    Printer.Print "     Tipo de Orden: " & tRs.Fields("TIPO")
    'DATOS PROVEEDOR
    Printer.Print ""
    Printer.Print ""
    Printer.Print "     Proveedor : " & tRs.Fields("NOMBRE")
    Printer.Print "----- BANCO ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "     Banco : " & tRs.Fields("TRANS_BANCO")
    Printer.Print "     Dirección : " & tRs.Fields("TRANS_DIRECCION")
    Printer.Print "     Ciudad : " & tRs.Fields("TRANS_CIUDAD")
    Printer.Print "     Routing : " & tRs.Fields("TRANS_ROUTING")
    Printer.Print "     Cuenta : " & tRs.Fields("TRANS_CUENTA")
    Printer.Print ""
    Printer.Print ""
    'DATOS PAGO
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "PRODUCTO"
    Printer.CurrentY = POSY
    Printer.CurrentX = 1600
    Printer.Print "Descripcion"
    Printer.CurrentY = POSY
    Printer.CurrentX = 9130
    Printer.Print "CANTIDAD"
    Printer.CurrentY = POSY
    Printer.CurrentX = 10200
    Printer.Print "PRECIO"
    Printer.Print "----- DETALLE ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Do While Not tRs.EOF
        POSY = POSY + 200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print tRs.Fields("ID_PRODUCTO")
        Printer.CurrentY = POSY
        Printer.CurrentX = 1600
        Printer.Print tRs.Fields("Descripcion")
        Printer.CurrentY = POSY
        Printer.CurrentX = 9130
        Printer.Print tRs.Fields("CANTIDAD")
        Printer.CurrentY = POSY
        Printer.CurrentX = 10200
        Printer.Print tRs.Fields("PRECIO")
        If POSY > 13600 Then
            POSY = 200
            Printer.EndDoc
        End If
        tRs.MoveNext
    Loop
    tRs.MoveFirst
    If POSY > 13600 Then
        Printer.EndDoc
    End If
    POSY = POSY + 200
    Printer.CurrentY = POSY + 200
    Printer.Print "----- PAGO -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    If POSY > 13600 Then
        Printer.EndDoc
    End If
    POSY = POSY + 200
    Printer.CurrentY = POSY
    Printer.Print "     Fecha del Pago : " & tRs.Fields("FECHA_PAGO")
    If POSY > 13600 Then
        Printer.EndDoc
    End If
    POSY = POSY + 200
    Printer.CurrentY = POSY
    Printer.Print "     No. de Orden : " & tRs.Fields("FREIGHT")
    If POSY > 13600 Then
        Printer.EndDoc
    End If
    POSY = POSY + 200
    Printer.CurrentY = POSY
    Printer.Print "     No. de Orden : " & tRs.Fields("TAX")
    If POSY > 13600 Then
        Printer.EndDoc
    End If
    POSY = POSY + 200
    Printer.CurrentY = POSY
    Printer.Print "     No. de Orden : " & tRs.Fields("TOTAL")
    If POSY > 13600 Then
        Printer.EndDoc
    End If
    POSY = POSY + 200
    Printer.CurrentY = POSY
    Printer.Print "     No. de Orden : " & tRs.Fields("TIPOPAGO")
    If POSY > 13600 Then
        Printer.EndDoc
    End If
    POSY = POSY + 200
    Printer.CurrentY = POSY
    Printer.Print "     No. de Orden : " & tRs.Fields("BANCO")
    If POSY > 13600 Then
        Printer.EndDoc
    End If
    POSY = POSY + 200
    Printer.CurrentY = POSY
    Printer.Print "     Numedo de Transferencia : " & tRs.Fields("NUMTRANS")
    If POSY > 13600 Then
        Printer.EndDoc
    End If
    POSY = POSY + 200
    Printer.CurrentY = POSY
    Printer.Print "     No. de Orden : " & tRs.Fields("NUMCHEQUE")
    If POSY > 13600 Then
        Printer.EndDoc
    End If
    POSY = POSY + 200
    Printer.CurrentY = POSY
    Printer.Print "     No. de Orden : " & tRs.Fields("CANTIDAD_OC")
    sBuscar = "SELECT  FECHA, BANCO, CANT_ABONO FROM ABONOS_PAGO_OC WHERE ID_ORDEN_COMPRA IN (" & Mid(lblID.Caption, 1, Len(lblID.Caption) - 2) & ")"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            POSY = POSY + 200
            Printer.Print tRs.Fields("FECHA") & Chr(32) & tRs.Fields("BANCO") & Chr(32) & tRs.Fields("CANT_ABONO")
            tRs.MoveNext
        Loop
    End If
    CommonDialog1.Copies = 1
    ListView1.ListItems.Clear
    Printer.EndDoc
    Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub lvwOCNacionales_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim Con As Integer
    Dim IdProve As String
    Dim sBuscar As String
    Dim ConSel As String
    Dim tRs As ADODB.Recordset
    ConSel = 0
    txtTotal.Text = "0"
    textsalpago.Text = "0"
    lblFolio.Caption = ""
    Label5.Caption = ""
    lblID.Caption = ""
    TipoOrden = "NACIONAL"
    tip = "N"
    opnNacional.Value = True
    For Con = 1 To lvwOCInternacionales.ListItems.Count
        lvwOCInternacionales.ListItems(Con).Checked = False
    Next
    For Con = 1 To lvwOCIndirectas.ListItems.Count
        lvwOCIndirectas.ListItems(Con).Checked = False
    Next
    For Con = 1 To lvwOCNacionales.ListItems.Count
        If lvwOCNacionales.ListItems(Con).Checked Then
            If IdProve = "" Then
                IdProve = lvwOCNacionales.ListItems(Con).SubItems(2)
                sBusca = "SELECT ORDEN_COMPRA.ID_PROVEEDOR AS PROVEE, ORDEN_COMPRA.FECHA AS FECHA, " & _
                         "       ORDENES_NO_SURTIDAS.NUM_ORDEN AS NUMORDEN, (ORDENES_NO_SURTIDAS.PRECIO * ORDENES_NO_SURTIDAS.CANTIDAD) AS IMPORTE " & _
                         "FROM ORDENES_NO_SURTIDAS INNER JOIN " & _
                         "     ORDEN_COMPRA ON ORDENES_NO_SURTIDAS.NUM_ORDEN = ORDEN_COMPRA.NUM_ORDEN AND " & _
                         "     ORDENES_NO_SURTIDAS.Tipo = ORDEN_COMPRA.Tipo " & _
                         "WHERE ORDEN_COMPRA.ID_PROVEEDOR = " & IdProve
                Set tRs = cnn.Execute(sBusca)
                Me.lvwOrdenesNSurtidas.ListItems.Clear
                txtSaldo.Text = "0"
                With tRs
                    While Not .EOF
                        Set tLi = lvwOrdenesNSurtidas.ListItems.Add(, , .Fields("PROVEE"))
                        If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(1) = .Fields("FECHA")
                        If Not IsNull(.Fields("NUMORDEN")) Then tLi.SubItems(2) = .Fields("NUMORDEN")
                        If Not IsNull(.Fields("IMPORTE")) Then
                            tLi.SubItems(2) = .Fields("IMPORTE")
                            txtSaldo.Text = CDbl(txtSaldo.Text) + CDbl(.Fields("IMPORTE"))
                        End If
                        .MoveNext
                    Wend
                    .Close
                End With
            End If
            If IdProve = lvwOCNacionales.ListItems(Con).SubItems(2) Then
                lblFolio.Caption = lblFolio.Caption & lvwOCNacionales.ListItems(Con) & ", "
                Label5.Caption = lvwOCNacionales.ListItems(Con).SubItems(3)
                txtTotal.Text = CDbl(txtTotal.Text) + CDbl(lvwOCNacionales.ListItems(Con).SubItems(4))
                sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN IN (" & Mid(lblFolio.Caption, 1, Len(lblFolio.Caption) - 2) & ") AND TIPO='N'"
                Set tRs = cnn.Execute(sBusca)
                If Not (tRs.EOF And tRs.BOF) Then
                    If Not IsNull(tRs.Fields("CANT_ABONO")) Then
                        textsalpago.Text = CDbl(txtTotal.Text) - CDbl(tRs.Fields("CANT_ABONO"))
                        TotPagar = textsalpago.Text
                    Else
                        textsalpago.Text = txtTotal.Text
                        TotPagar = textsalpago.Text
                    End If
                End If
                textsalpago.Text = CDbl(textsalpago.Text) - CDbl(txtSaldo.Text)
                sPendiente = textsalpago.Text
                lblID.Caption = lblID.Caption & lvwOCNacionales.ListItems(Con).SubItems(1) & ", "
            Else
                lvwOCNacionales.ListItems(Con).Checked = False
                MsgBox "TODAS LAS ORDENES DEBEN SER DEL MISMO PROVEEDOR !!", vbExclamation, "SACC"
            End If
            ConSel = ConSel + 1
            If CDbl(ConSel) > 1 Then
                Command3.Enabled = False
                textsalpago.Enabled = False
            Else
                Command3.Enabled = True
            End If
        End If
    Next
End Sub
Public Function ArchivoEnUso(ByVal sFileName As String) As Boolean
    Dim filenum As Integer, errnum As Integer
    On Error Resume Next
    filenum = FreeFile()
    Open sFileName For Input Lock Read As #filenum
    Close filenum
    errnum = Err
    On Error GoTo 0
    Select Case errnum
        Case 0
            ArchivoEnUso = False
        Case 70
            ArchivoEnUso = True
        Case Else
            Error errnum
    End Select
End Function
