VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmEmpataFolios 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Folios del SAT"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5520
      TabIndex        =   1
      Top             =   1680
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmEmpataFolios.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmEmpataFolios.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmEmpataFolios.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DTPicker1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DTPicker2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAgregar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Ejecutar"
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
         Left            =   2040
         Picture         =   "FrmEmpataFolios.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   3000
         TabIndex        =   6
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   50855937
         CurrentDate     =   44729
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   50855937
         CurrentDate     =   44729
      End
      Begin VB.Label Label2 
         Caption         =   "Al :"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Del :"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmEmpataFolios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private cnn1 As ADODB.Connection
Private Sub cmdAgregar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT tblEmpresa.RFC_empresa, tblSerie.Codigo_serie, tblFactura.Num_factura, tblFactura.Folio_SAT, tblFactura.Fecha_factura FROM tblEmpresa INNER JOIN tblFactura ON tblEmpresa.Sec_empresa = tblFactura.Sec_empresa INNER JOIN tblSerie ON tblFactura.Sec_empresa = tblSerie.Sec_empresa AND tblFactura.Sec_sucursal = tblSerie.Sec_sucursal WHERE tblEmpresa.RFC_empresa = '" & VarMen.TxtEmp(8).Text & "' AND tblFactura.Fecha_factura BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
    Set tRs = cnn1.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            sBuscar = "UPDATE VENTAS SET UUID = '" & tRs.Fields("Folio_SAT") & "' WHERE FOLIO = '" & tRs.Fields("Codigo_serie") & tRs.Fields("Num_factura") & "'"
            cnn.Execute (sBuscar)
            tRs.MoveNext
        Loop
    End If
    MsgBox "Proceso finalizado con éxito!", vbInformation, "SACC"
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Set cnn1 = New ADODB.Connection
    With cnn1
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "BDFacturacion", "FacturaGlobal") & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
