VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmModiVenta 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar tipo de venta"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Forma de Pago"
      TabPicture(0)   =   "FrmModiVenta.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Option5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Option4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.OptionButton Option8 
         Caption         =   "No Aplica"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   2280
         Width           =   1695
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Tarjeta de Debito"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Transf. electrónica"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Efectivo"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Tarjeta de Credito"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Cheque"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdBuscar2 
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
      Left            =   2040
      Picture         =   "FrmModiVenta.frx":001C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2990
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Factura"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nota"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5880
      TabIndex        =   3
      Top             =   360
      Width           =   975
      Begin VB.Label Label8 
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmModiVenta.frx":29EE
         MousePointer    =   99  'Custom
         Picture         =   "FrmModiVenta.frx":2CF8
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5880
      TabIndex        =   1
      Top             =   1560
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmModiVenta.frx":46BA
         MousePointer    =   99  'Custom
         Picture         =   "FrmModiVenta.frx":49C4
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "FrmModiVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim NumVenta As String
Private Sub cmdBuscar2_Click()
    If Text1.Text <> "" Then
       ListView1.ListItems.Clear
        Buscar
    End If
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
        .ColumnHeaders.Add , , "ID_VENTA", 0
        .ColumnHeaders.Add , , "NOMBRE", 1440
        .ColumnHeaders.Add , , "TOTAL", 1000
        .ColumnHeaders.Add , , "FOLIO", 1000
        .ColumnHeaders.Add , , "FORMA PAGO", 0
    End With
End Sub
Private Sub Image8_Click()
    Dim sBuscar As String
    Dim FormPago As String
    Dim tRs As ADODB.Recordset
        Dim tRs2 As ADODB.Recordset
    If Option3.Value = True Then
        FormPago = "C"
    End If
    If Option4.Value = True Then
        FormPago = "T"
    End If
    If Option5.Value = True Then
        FormPago = "H"
    End If
    If Option6.Value = True Then
        FormPago = "E"
    End If
    If Option7.Value = True Then
        FormPago = "D"
    End If
    If Option8.Value = True Then
        FormPago = "N"
    End If
    If Option1.Value = True Then
        sBuscar = "UPDATE VENTAS SET TIPO_PAGO = '" & FormPago & "' WHERE ID_VENTA =' " & Text1.Text & "'"
        cnn.Execute (sBuscar)
        MsgBox "LOS CAMBIOS SE REALIZARON CON EXITO!", vbInformation, "SACC"
    End If
    If Option2.Value = True Then
        sBuscar = "SELECT * FROM VENTAS WHERE FOLIO = '" & Text1.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                sBuscar = "UPDATE VENTAS SET TIPO_PAGO = '" & FormPago & "' WHERE ID_VENTA =' " & tRs.Fields("ID_VENTA") & "'"
                Set tRs2 = cnn.Execute(sBuscar)
                tRs.MoveNext
            Loop
            MsgBox "LOS CAMBIOS SE REALIZARON CON EXITO!", vbInformation, "SACC"
        End If
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    NumVenta = Item
    If Item.SubItems(4) = "C" Then
        Option3.Value = True
    End If
    If Item.SubItems(4) = "H" Then
        Option5.Value = True
    End If
    If Item.SubItems(4) = "T" Then
        Option4.Value = True
    End If
    If Item.SubItems(4) = "E" Then
        Option6.Value = True
    End If
    If Item.SubItems(4) = "D" Then
        Option7.Value = True
    End If
    If Item.SubItems(4) = "N" Then
        Option8.Value = True
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    ListView1.ListItems.Clear
    If Text1.Text <> "" And KeyAscii = 13 Then
        Buscar
    End If
    If Option1.Value = True Then
        Valido = "1234567890"
    Else
        Valido = "1234567890ABCDEFGHIJKLMNÑOPQRSTUVWXYZ"
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Buscar()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If Option1.Value = True Then
        sBuscar = "SELECT NOMBRE, ID_VENTA, TOTAL, FOLIO, TIPO_PAGO FROM VENTAS WHERE ID_VENTA = " & Text1.Text
    Else
        sBuscar = "SELECT NOMBRE, ID_VENTA, TOTAL, FOLIO, TIPO_PAGO FROM VENTAS WHERE FOLIO = '" & Text1.Text & "'"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(2) = tRs.Fields("TOTAL")
            If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(3) = tRs.Fields("FOLIO")
            If Not IsNull(tRs.Fields("TIPO_PAGO")) Then tLi.SubItems(4) = tRs.Fields("TIPO_PAGO")
            tRs.MoveNext
        Loop
    Else
        MsgBox "LA VENTA NO FUE ENCONTRADA, REVISE SI LA VENTA ES DE CREDITO", vbInformation, "SACC"
    End If
End Sub
