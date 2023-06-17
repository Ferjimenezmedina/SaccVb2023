VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRepCXC2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Cuentas CxC"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10920
      TabIndex        =   2
      Top             =   3960
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepCXC2.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepCXC2.frx":030A
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10920
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Detalles de Cilentes CxC"
      TabPicture(0)   =   "FrmRepCXC2.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "textid"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtnom"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtnom 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox textid 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   6588
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
         Caption         =   "Num Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre Cliente :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmRepCXC2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Form_Load()
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
        .ColumnHeaders.Add , , "FACTURA", 1200
        .ColumnHeaders.Add , , "SUBTOTAL", 1200
        .ColumnHeaders.Add , , "IVA", 1300
        .ColumnHeaders.Add , , "TOTAL", 1300
        .ColumnHeaders.Add , , "FECHA", 1300
        .ColumnHeaders.Add , , "FECHA DE VENCIMIENTO", 4000
        .ColumnHeaders.Add , , "DIAS VENCIDOS", 1300
        .ColumnHeaders.Add , , "PAGADA", 1300
    End With
    Busca
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Busca()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT * FROM  VsRepCXC  WHERE  ID_CLIENTE='" & textid.Text & "' "
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("FOLIO"))
            If Not IsNull(tRs.Fields("SUBTOTAL")) Then tLi.SubItems(1) = tRs.Fields("SUBTOTAL")
            If Not IsNull(tRs.Fields("IVA")) Then tLi.SubItems(2) = tRs.Fields("IVA")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(3) = tRs.Fields("TOTAL")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(4) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("FECHA_PAGARE")) Then tLi.SubItems(5) = tRs.Fields("FECHA_PAGARE")
            If Not IsNull(tRs.Fields("FECHA_PAGARE")) Then tLi.SubItems(6) = tRs.Fields("FECHA_PAGARE")
            If tRs.Fields("PAGADA") = "S" Then
                tLi.SubItems(7) = "PAGADA"
            Else
                tLi.SubItems(7) = "PENDIENTE"
            End If
            tRs.MoveNext
        Loop
    End If
End Sub
