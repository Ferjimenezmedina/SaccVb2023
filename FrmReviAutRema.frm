VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmReviAutRema 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revisar Autorizacion de Remanufactura"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Comanda"
      Height          =   1335
      Left            =   9360
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
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
         Left            =   120
         Picture         =   "FrmReviAutRema.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9600
      TabIndex        =   6
      Top             =   5400
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmReviAutRema.frx":29D2
         MousePointer    =   99  'Custom
         Picture         =   "FrmReviAutRema.frx":2CDC
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11456
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Autorizadas"
      TabPicture(0)   =   "FrmReviAutRema.frx":4DBE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Rechazada/Cancelada"
      TabPicture(1)   =   "FrmReviAutRema.frx":4DDA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Pendientes"
      TabPicture(2)   =   "FrmReviAutRema.frx":4DF6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSComctlLib.ListView ListView3 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   2
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9763
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   1
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9763
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
         Height          =   5535
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9763
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
Attribute VB_Name = "FrmReviAutRema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    Buscar
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    Dim tRs As ADODB.Recordset
    Dim sBusqueda As String
    Dim tLi As ListItem
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "COMANDA", 800
        .ColumnHeaders.Add , , "CANTIDAD", 1200
        .ColumnHeaders.Add , , "ID_PROD", 2000
        .ColumnHeaders.Add , , "CLIENTE", 4000
        .ColumnHeaders.Add , , "ARTICULO", 0
        .ColumnHeaders.Add , , "SUCURSAL", 2000
    End With
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "COMANDA", 800
        .ColumnHeaders.Add , , "CANTIDAD", 1200
        .ColumnHeaders.Add , , "ID_PROD", 2000
        .ColumnHeaders.Add , , "CLIENTE", 4000
        .ColumnHeaders.Add , , "ARTICULO", 0
        .ColumnHeaders.Add , , "SUCURSAL", 2000
    End With
    With ListView3
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "COMANDA", 800
        .ColumnHeaders.Add , , "CANTIDAD", 1200
        .ColumnHeaders.Add , , "ID_PROD", 2000
        .ColumnHeaders.Add , , "CLIENTE", 4000
        .ColumnHeaders.Add , , "ARTICULO", 0
        .ColumnHeaders.Add , , "SUCURSAL", 2000
    End With
    Buscar
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Buscar()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If Text1.Text = "" Then
        sBuscar = "SELECT * FROM VSCOMREMA WHERE  ESTADO_ACTUAL = 'Z'"
    Else
        sBuscar = "SELECT * FROM VSCOMREMA WHERE  ID_COMANDA = '" & Text1.Text & "' AND ESTADO_ACTUAL = 'Z'"
    End If
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Set tLi = ListView3.ListItems.Add(, , .Fields("ID_COMANDA") & "")
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = .Fields("ID_PRODUCTO") & ""
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(3) = .Fields("NOMBRE") & ""
                If Not IsNull(.Fields("ARTICULO")) Then tLi.SubItems(4) = .Fields("ARTICULO") & ""
                If Not IsNull(.Fields("SUCURSAL")) Then tLi.SubItems(5) = .Fields("SUCURSAL") & ""
                .MoveNext
            Loop
        End If
    End With
    If Text1.Text = "" Then
        sBuscar = "SELECT * FROM VSCOMREMA WHERE ESTADO_ACTUAL = 'A'"
    Else
        sBuscar = "SELECT * FROM VSCOMREMA WHERE ESTADO_ACTUAL = 'A' AND ID_COMANDA = " & Text1.Text
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_COMANDA") & "")
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = .Fields("ID_PRODUCTO") & ""
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(3) = .Fields("NOMBRE") & ""
                If Not IsNull(.Fields("ARTICULO")) Then tLi.SubItems(4) = .Fields("ARTICULO") & ""
                If Not IsNull(.Fields("SUCURSAL")) Then tLi.SubItems(5) = .Fields("SUCURSAL") & ""
                .MoveNext
            Loop
        End If
    End With
    If Text1.Text = "" Then
        sBuscar = "SELECT * FROM VSCOMREMA WHERE ESTADO_ACTUAL = 'C'"
    Else
        sBuscar = "SELECT * FROM VSCOMREMA WHERE ESTADO_ACTUAL = 'C' AND ID_COMANDA = " & Text1.Text
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_COMANDA") & "")
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = .Fields("ID_PRODUCTO") & ""
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(3) = .Fields("NOMBRE") & ""
                If Not IsNull(.Fields("ARTICULO")) Then tLi.SubItems(4) = .Fields("ARTICULO") & ""
                If Not IsNull(.Fields("SUCURSAL")) Then tLi.SubItems(5) = .Fields("SUCURSAL") & ""
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
