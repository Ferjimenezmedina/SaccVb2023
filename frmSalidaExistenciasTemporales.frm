VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSalidaExistenciasTemporales 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Liberar existencias temporales"
   ClientHeight    =   5655
   ClientLeft      =   4155
   ClientTop       =   2100
   ClientWidth     =   9870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   9870
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Existencias temporales "
      TabPicture(0)   =   "frmSalidaExistenciasTemporales.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Existencia"
         Height          =   1935
         Left            =   240
         TabIndex        =   17
         Top             =   3240
         Width           =   8175
         Begin VB.TextBox txtexdetalle 
            Enabled         =   0   'False
            Height          =   195
            Left            =   7680
            TabIndex        =   24
            Top             =   1080
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   195
            Left            =   7440
            TabIndex        =   18
            Top             =   1080
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            TabIndex        =   2
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4080
            TabIndex        =   3
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   4
            Top             =   720
            Width           =   6855
         End
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   5
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            TabIndex        =   6
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Clave del producto"
            Height          =   255
            Left            =   2640
            TabIndex        =   22
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Descripcion"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Sucursal"
            Height          =   255
            Left            =   2160
            TabIndex        =   19
            Top             =   1080
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2535
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4471
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
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   8
      Top             =   3120
      Width           =   975
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   13
         Top             =   1320
         Width           =   975
         Begin VB.Image Image15 
            Height          =   720
            Left            =   120
            MouseIcon       =   "frmSalidaExistenciasTemporales.frx":001C
            MousePointer    =   99  'Custom
            Picture         =   "frmSalidaExistenciasTemporales.frx":0326
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label16 
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
            TabIndex        =   14
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   11
         Top             =   1320
         Width           =   975
         Begin VB.Label Label17 
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
            TabIndex        =   12
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "frmSalidaExistenciasTemporales.frx":1CE8
            MousePointer    =   99  'Custom
            Picture         =   "frmSalidaExistenciasTemporales.frx":1FF2
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   9
         Top             =   0
         Width           =   975
         Begin VB.Label Label18 
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
            TabIndex        =   10
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image17 
            Height          =   705
            Left            =   120
            MouseIcon       =   "frmSalidaExistenciasTemporales.frx":381C
            MousePointer    =   99  'Custom
            Picture         =   "frmSalidaExistenciasTemporales.frx":3B26
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Eliminar"
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
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image18 
         Enabled         =   0   'False
         Height          =   750
         Left            =   120
         MouseIcon       =   "frmSalidaExistenciasTemporales.frx":55D8
         MousePointer    =   99  'Custom
         Picture         =   "frmSalidaExistenciasTemporales.frx":58E2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   0
      Top             =   4320
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmSalidaExistenciasTemporales.frx":760C
         MousePointer    =   99  'Custom
         Picture         =   "frmSalidaExistenciasTemporales.frx":7916
         Top             =   120
         Width           =   720
      End
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
   End
End
Attribute VB_Name = "frmSalidaExistenciasTemporales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    On Error GoTo ManejaError
    Dim sBuscar As String
    Dim sBuscar2 As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tLi As ListItem
    Dim i As Integer
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
        .ColumnHeaders.Add , , "N-MOV", 0
        .ColumnHeaders.Add , , "FECHA", 1600
        .ColumnHeaders.Add , , "Clave del Producto", 1800
        .ColumnHeaders.Add , , "Descripcion", 5400
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "SUCURSAL", 1600
        .ColumnHeaders.Add , , "ID-MOVDETALLE", 0
    End With
    sBuscar = "SELECT M.ID_MOVEXISTENCIA, M.FECHA, M.ID_CLIENTE, S.CANTIDAD, S.ID_PRODUCTO, S.SUCURSAL, S.ID_EXDETALLE FROM EXISTENCIAS_TEMPORAL AS M JOIN EXISTENCIAS_TEMPORAL_DETALLES AS S ON M.ID_MOVEXISTENCIA = S.ID_MOVEXISTENCIA  ORDER BY FECHA"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        ListView1.ListItems.Clear
        Do While Not .EOF
            Set tLi = ListView1.ListItems.Add(, , .Fields("ID_MOVEXISTENCIA") & "")
            If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(1) = .Fields("FECHA") & ""
            If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = .Fields("ID_PRODUCTO") & ""
            If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(4) = .Fields("CANTIDAD") & ""
            If Not IsNull(.Fields("SUCURSAL")) Then tLi.SubItems(5) = .Fields("SUCURSAL") & ""
            If Not IsNull(.Fields("ID_EXDETALLE")) Then tLi.SubItems(6) = .Fields("ID_EXDETALLE") & ""
            sBuscar2 = "SELECT ID_PRODUCTO,Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' "
            Set tRs2 = cnn.Execute(sBuscar2)
            If Not IsNull(tRs2.Fields("Descripcion")) Then tLi.SubItems(3) = tRs2.Fields("Descripcion") & ""
            .MoveNext
        Loop
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image18_Click()
    Dim sBuscar As String
    Dim sBuscar2 As String
    Dim sBuscar3 As String
    Dim sBuscar4 As String
    Dim sBuscar5 As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tRs5 As ADODB.Recordset
    Dim operacionCambioExistencias As Double
    Dim cant As Double
    cant = Text5.Text
        sBuscar = "DELETE FROM EXISTENCIAS_TEMPORAL_DETALLES WHERE ID_EXDETALLE = '" & txtexdetalle & "'"
        Set tRs = cnn.Execute(sBuscar)
        sBuscar2 = "SELECT * FROM EXISTENCIAS_TEMPORAL_DETALLES WHERE ID_MOVEXISTENCIA = '" & Text1.Text & "'"
       Set tRs2 = cnn.Execute(sBuscar2)
    If (tRs2.BOF And tRs2.EOF) Then
        sBuscar3 = "DELETE FROM EXISTENCIAS_TEMPORAL WHERE ID_MOVEXISTENCIA = '" & Text1.Text & "'"
        Set tRs3 = cnn.Execute(sBuscar3)
    End If
    sBuscar4 = "SELECT ID_PRODUCTO,CANTIDAD,SUCURSAL FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Text3.Text & "' AND SUCURSAL = '" & Text6.Text & "'"
    Set tRs4 = cnn.Execute(sBuscar4)
    If Not (tRs4.EOF And tRs4.BOF) Then
        If Not IsNull(tRs4.Fields("CANTIDAD")) Then
            operacionCambioExistencias = CDbl(tRs4.Fields("CANTIDAD")) - CDbl(cant)
            sBuscar5 = "UPDATE EXISTENCIAS SET CANTIDAD = " & operacionCambioExistencias & " WHERE SUCURSAL = '" & Text6.Text & " ' AND ID_PRODUCTO = '" & Text3.Text & "'"
            Set tRs5 = cnn.Execute(sBuscar5)
        End If
    End If
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    txtexdetalle = ""
    Image18.Enabled = False
    ActualizarListView
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item
    Text2.Text = Item.SubItems(1)
    Text3.Text = Item.SubItems(2)
    Text4.Text = Item.SubItems(3)
    Text5.Text = Item.SubItems(4)
    Text6.Text = Item.SubItems(5)
    txtexdetalle = Item.SubItems(6)
    Image18.Enabled = True
End Sub
Public Sub ActualizarListView()
    Dim sBuscar As String
    Dim sBuscar2 As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT M.ID_MOVEXISTENCIA, M.FECHA, M.ID_CLIENTE, S.CANTIDAD,S.ID_PRODUCTO,S.SUCURSAL, S.ID_EXDETALLE FROM EXISTENCIAS_TEMPORAL AS M JOIN EXISTENCIAS_TEMPORAL_DETALLES AS S ON M.ID_MOVEXISTENCIA = S.ID_MOVEXISTENCIA  ORDER BY FECHA"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        ListView1.ListItems.Clear
        Do While Not .EOF
            Set tLi = ListView1.ListItems.Add(, , .Fields("ID_MOVEXISTENCIA") & "")
            If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(1) = .Fields("FECHA") & ""
            If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = .Fields("ID_PRODUCTO") & ""
            If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(4) = .Fields("CANTIDAD") & ""
            If Not IsNull(.Fields("SUCURSAL")) Then tLi.SubItems(5) = .Fields("SUCURSAL") & ""
            If Not IsNull(.Fields("ID_EXDETALLE")) Then tLi.SubItems(6) = .Fields("ID_EXDETALLE") & ""
            sBuscar2 = "SELECT ID_PRODUCTO,Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' "
            Set tRs2 = cnn.Execute(sBuscar2)
            If Not IsNull(tRs2.Fields("Descripcion")) Then tLi.SubItems(3) = tRs2.Fields("Descripcion") & ""
            .MoveNext
        Loop
    End With
End Sub
