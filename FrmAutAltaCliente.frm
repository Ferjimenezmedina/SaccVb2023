VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmAutAltaCliente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorizar alta de cliente"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   6
      Top             =   3240
      Width           =   975
      Begin VB.Image Image6 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmAutAltaCliente.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmAutAltaCliente.frx":030A
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
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
   Begin VB.Frame Frame21 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   4
      Top             =   2040
      Width           =   975
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Aceptar"
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
      Begin VB.Image imgLeer 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmAutAltaCliente.frx":1DBC
         MousePointer    =   99  'Custom
         Picture         =   "FrmAutAltaCliente.frx":20C6
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   2
      Top             =   4440
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmAutAltaCliente.frx":3B78
         MousePointer    =   99  'Custom
         Picture         =   "FrmAutAltaCliente.frx":3E82
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmAutAltaCliente.frx":5F64
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.ListView ListView1 
         Height          =   4935
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8705
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
Attribute VB_Name = "FrmAutAltaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Form_Load()
    On Error GoTo ManejaError
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
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
        .Checkboxes = True
        .ColumnHeaders.Add , , "ID", 300
        .ColumnHeaders.Add , , "NOMBRE", 5500
        .ColumnHeaders.Add , , "RFC", 1500
        .ColumnHeaders.Add , , "Direccion", 1500
        .ColumnHeaders.Add , , "Num. Ext - Int", 2000
        .ColumnHeaders.Add , , "CP", 1000
        .ColumnHeaders.Add , , "Ciudad", 2000
        .ColumnHeaders.Add , , "Estado", 2000
        .ColumnHeaders.Add , , "Pais", 2000
        .ColumnHeaders.Add , , "Telefono", 2000
        .ColumnHeaders.Add , , "e-Mail", 2000
        .ColumnHeaders.Add , , "Fecha Alta", 2000
    End With
    Buscar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub

Private Sub Image6_Click()
    Dim Con As Integer
    Dim sBuscar As String
    For Con = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(Con).Checked Then
            sBuscar = "UPDATE CLIENTE SET VALORACION = 'E' WHERE ID_CLIENTE = " & ListView1.ListItems(Con)
            cnn.Execute (sBuscar)
        End If
    Next Con
    Buscar
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Buscar()
    Dim tRs As ADODB.Recordset
    Dim sBsucar As String
    Dim tLi As ListItem
    sBuscar = "SELECT * FROM CLIENTE WHERE VALORACION = 'R' ORDER BY ID_CLIENTE"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("RFC")) Then tLi.SubItems(2) = tRs.Fields("RFC")
            If Not IsNull(tRs.Fields("DIRECCION")) Then tLi.SubItems(3) = tRs.Fields("DIRECCION")
            If Not IsNull(tRs.Fields("NUMERO_EXTERIOR")) Then tLi.SubItems(4) = tRs.Fields("NUMERO_EXTERIOR") & " - " & tRs.Fields("NUMERO_INTERIOR")
            If Not IsNull(tRs.Fields("CP")) Then tLi.SubItems(5) = tRs.Fields("CP")
            If Not IsNull(tRs.Fields("CIUDAD")) Then tLi.SubItems(6) = tRs.Fields("CIUDAD")
            If Not IsNull(tRs.Fields("ESTADO")) Then tLi.SubItems(7) = tRs.Fields("ESTADO")
            If Not IsNull(tRs.Fields("PAIS")) Then tLi.SubItems(8) = tRs.Fields("PAIS")
            If Not IsNull(tRs.Fields("TELEFONO_CASA")) Then tLi.SubItems(9) = tRs.Fields("TELEFONO_CASA")
            If Not IsNull(tRs.Fields("EMAIL")) Then tLi.SubItems(10) = tRs.Fields("EMAIL")
            If Not IsNull(tRs.Fields("FECHA_ALTA")) Then tLi.SubItems(11) = tRs.Fields("FECHA_ALTA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub imgLeer_Click()
    Dim Con As Integer
    Dim sBuscar As String
    For Con = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(Con).Checked Then
            sBuscar = "UPDATE CLIENTE SET VALORACION = 'A' WHERE ID_CLIENTE = " & ListView1.ListItems(Con)
            cnn.Execute (sBuscar)
        End If
    Next Con
    Buscar
End Sub

