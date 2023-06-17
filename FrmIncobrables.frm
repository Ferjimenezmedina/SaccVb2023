VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmIncobrables 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas Incobrables"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame26 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9120
      TabIndex        =   4
      Top             =   3840
      Width           =   975
      Begin VB.Image Image24 
         Height          =   765
         Left            =   240
         MouseIcon       =   "FrmIncobrables.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmIncobrables.frx":030A
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Regresar"
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
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9120
      TabIndex        =   2
      Top             =   5040
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmIncobrables.frx":1D98
         MousePointer    =   99  'Custom
         Picture         =   "FrmIncobrables.frx":20A2
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmIncobrables.frx":4184
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.ListView ListView1 
         Height          =   5655
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   9975
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
Attribute VB_Name = "FrmIncobrables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.txtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Cuenta", 1000
        .ColumnHeaders.Add , , "Venta", 1000
        .ColumnHeaders.Add , , "Folio", 1000
        .ColumnHeaders.Add , , "Cliente", 3500
        .ColumnHeaders.Add , , "Subtotal", 1500
        .ColumnHeaders.Add , , "IVA", 1500
        .ColumnHeaders.Add , , "Total", 1500
        .ColumnHeaders.Add , , "Sucursal", 1500
        .ColumnHeaders.Add , , "Vencimiento", 1500
    End With
    Actualiza
End Sub
Private Sub Actualiza()
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim tLi As ListItem
    sBuscar = "SELECT CUENTAS.ID_CUENTA, VENTAS.ID_VENTA, VENTAS.FOLIO, VENTAS.NOMBRE, VENTAS.SUBTOTAL, VENTAS.IVA, VENTAS.TOTAL, Ventas.Sucursal, cuentas.FECHA_VENCE FROM CUENTAS INNER JOIN CUENTA_VENTA ON CUENTAS.ID_CUENTA = CUENTA_VENTA.ID_CUENTA INNER JOIN VENTAS ON CUENTA_VENTA.ID_VENTA = VENTAS.ID_VENTA WHERE (CUENTAS.PAGADA = 'I')"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CUENTA"))
            If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(1) = tRs.Fields("ID_VENTA")
            If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(2) = tRs.Fields("FOLIO")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(3) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("SUBTOTAL")) Then tLi.SubItems(4) = tRs.Fields("SUBTOTAL")
            If Not IsNull(tRs.Fields("IVA")) Then tLi.SubItems(5) = tRs.Fields("IVA")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(6) = tRs.Fields("TOTAL")
            If Not IsNull(tRs.Fields("Sucursal")) Then tLi.SubItems(7) = tRs.Fields("Sucursal")
            If Not IsNull(tRs.Fields("FECHA_VENCE")) Then tLi.SubItems(8) = tRs.Fields("FECHA_VENCE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image24_Click()
    Dim Con As Integer
    Dim sBuscar As String
    For Con = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(Con).Checked Then
            sBuscar = "UPDATE CUENTAS SET PAGADA = 'N' WHERE ID_CUENTA = " & ListView1.ListItems(Con)
            cnn.Execute (sBuscar)
        End If
    Next Con
    Actualiza
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
