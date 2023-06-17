VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmBuscaComanda 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Comandas"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8880
      TabIndex        =   17
      Top             =   5640
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
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmBuscaComanda.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmBuscaComanda.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Rastrear"
      TabPicture(0)   =   "FrmBuscaComanda.frx":23EC
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
      Tab(0).Control(5)=   "ListView1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListView2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ListView3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   120
         TabIndex        =   20
         Top             =   5100
         Width           =   2655
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Garantia"
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
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   2055
         End
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1575
         Left            =   3000
         TabIndex        =   19
         Top             =   4740
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2778
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   300
         Width           =   4575
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
         Left            =   5760
         Picture         =   "FrmBuscaComanda.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1695
         Left            =   3000
         TabIndex        =   11
         Top             =   2820
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2990
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame1 
         Caption         =   " Filtrar Por "
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   2775
         Begin VB.CheckBox Check2 
            Caption         =   "Solo de mi sucursal"
            Height          =   255
            Left            =   480
            TabIndex        =   16
            Top             =   2760
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   600
            TabIndex        =   8
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50987009
            CurrentDate     =   39122
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Por Fecha"
            Height          =   255
            Left            =   720
            TabIndex        =   7
            Top             =   1560
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Clave del Producto"
            Height          =   195
            Left            =   600
            TabIndex        =   6
            Top             =   960
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            Caption         =   "No. Comanda"
            Height          =   195
            Left            =   600
            TabIndex        =   5
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Nombre del Cliente"
            Height          =   195
            Left            =   600
            TabIndex        =   4
            Top             =   480
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   3000
         TabIndex        =   2
         Top             =   900
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2778
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label5 
         Caption         =   "Comandas"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Detalle"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   2580
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "LAS COMANDAS COBRADAS NO APARECEN EN ESTE LISTADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   4380
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "NOTA :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   4020
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar :"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   300
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmBuscaComanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        DTPicker1.Enabled = True
    Else
        DTPicker1.Enabled = False
    End If
End Sub
Private Sub Command1_Click()
    BuscaComanda
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Enabled = False
    DTPicker1.Value = Format(Date, "dd/mm/yyyy")
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
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. Comanda", 1200
        .ColumnHeaders.Add , , "Cliente", 5500
        .ColumnHeaders.Add , , "Telefono", 1850
        .ColumnHeaders.Add , , "Fecha", 1850
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 1850
        .ColumnHeaders.Add , , "Descripcion", 1850
        .ColumnHeaders.Add , , "Estado Actual", 1400
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Funcionaron", 1200
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Id_Garantia", 1000
        .ColumnHeaders.Add , , "Fecha", 1000
        .ColumnHeaders.Add , , "Nombre", 3000
        .ColumnHeaders.Add , , "Id_Venta", 1000
        .ColumnHeaders.Add , , "Id_producto", 2000
        .ColumnHeaders.Add , , "Cantidad", 1000
         .ColumnHeaders.Add , , "Cant_Aceptada", 1000
        .ColumnHeaders.Add , , "Estado Actual", 1400
        .ColumnHeaders.Add , , "Comentario", 3000
        .ColumnHeaders.Add , , "Llegaron", 1200
        .ColumnHeaders.Add , , "Llegaron", 1200
          .ColumnHeaders.Add , , "Produccion", 3000
    End With
End Sub
Private Sub BuscaComanda()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    If Option1.Value = True Then
        sBuscar = "SELECT NOMBRE, TELEFONO, ID_COMANDA, FECHA_INICIO FROM VsRevComanda WHERE NOMBRE LIKE '%" & Text1.Text & "%'"
    Else
        If Option2.Value = True Then
            sBuscar = "SELECT NOMBRE, TELEFONO, ID_COMANDA, FECHA_INICIO FROM VsRevComanda WHERE ID_COMANDA = " & Text1.Text
        Else
            sBuscar = "SELECT NOMBRE, TELEFONO, ID_COMANDA, FECHA_INICIO FROM VsRevComanda WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%'"
        End If
    End If
    If Check1.Value = 1 Then
        sBuscar = sBuscar & " AND FECHA_INICIO = '" & DTPicker1.Value & "'"
    End If
    If Check2.Value = 1 Then
        sBuscar = sBuscar & " AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
    End If
    sBuscar = sBuscar & " AND ESTADO_ACTUAL <> 'I' GROUP BY ID_COMANDA, NOMBRE, TELEFONO, FECHA_INICIO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_COMANDA"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("TELEFONO")) Then tLi.SubItems(2) = tRs.Fields("TELEFONO")
            If Not IsNull(tRs.Fields("FECHA_INICIO")) Then tLi.SubItems(3) = tRs.Fields("FECHA_INICIO")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView2.ListItems.Clear
    sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, ESTADO_ACTUAL, CANTIDAD, CANT_FUNCIONO FROM VsRevComanda WHERE ID_COMANDA = " & Item
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("ESTADO_ACTUAL")) Then
                If tRs.Fields("ESTADO_ACTUAL") = "A" Then
                    tLi.SubItems(2) = "Nueva"
                Else
                    If tRs.Fields("ESTADO_ACTUAL") = "R" Or tRs.Fields("ESTADO_ACTUAL") = "S" Then
                        tLi.SubItems(2) = "En Producción"
                    Else
                        If tRs.Fields("ESTADO_ACTUAL") = "P" Then
                            tLi.SubItems(2) = "Probando en Calidad"
                        Else
                            If tRs.Fields("ESTADO_ACTUAL") = "N" Or tRs.Fields("ESTADO_ACTUAL") = "M" Then
                                tLi.SubItems(2) = "Cartuchos Dañados"
                            Else
                                If tRs.Fields("ESTADO_ACTUAL") = "L" Then
                                    tLi.SubItems(2) = "Terminado"
                                Else
                                    If tRs.Fields("ESTADO_ACTUAL") = "Z" Then
                                        tLi.SubItems(2) = "Aprobar Rema"
                                    Else
                                        If tRs.Fields("ESTADO_ACTUAL") = "C" Or tRs.Fields("ESTADO_ACTUAL") = "0" Then
                                            tLi.SubItems(2) = "CANCELADA"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(3) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("CANT_FUNCIONO")) Then tLi.SubItems(4) = tRs.Fields("CANT_FUNCIONO")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Option2_Click()
    Text1.Text = ""
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BuscaComanda
    End If
    If Option2.Value = True Then
        Dim Valido As String
        Valido = "1234567890."
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView3.ListItems.Clear
    If KeyAscii = 13 Then
        sBuscar = "SELECT * FROM vsgarantias WHERE ID_VENTA= '" & Text2.Text & "' ORDER BY FECHA DESC"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_GARANTIA"))
                tLi.SubItems(1) = tRs.Fields("FECHA")
                tLi.SubItems(2) = tRs.Fields("NOMBRE")
                tLi.SubItems(3) = tRs.Fields("ID_VENTA")
                tLi.SubItems(4) = tRs.Fields("ID_PRODUCTO")
                If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(5) = tRs.Fields("CANTIDAD")
                If Not IsNull(tRs.Fields("CANT_ACEP")) Then tLi.SubItems(6) = tRs.Fields("CANT_ACEP")
                If tRs.Fields("ESTADO") = "A" Then
                    tLi.SubItems(7) = "ACEPTADA"
                End If
                If tRs.Fields("ESTADO") = "N" Then
                    tLi.SubItems(7) = "RECHAZADA"
                End If
                If tRs.Fields("ESTADO") = "P" Then
                    tLi.SubItems(7) = "PENDIENTE"
                End If
                If Not IsNull(tRs.Fields("COMENTARIO")) Then tLi.SubItems(8) = tRs.Fields("COMENTARIO")
                If Not IsNull(tRs.Fields("REVUNO")) Then tLi.SubItems(9) = tRs.Fields("REVUNO")
                If Not IsNull(tRs.Fields("REVDOS")) Then tLi.SubItems(10) = tRs.Fields("REVDOS")
                If Not IsNull(tRs.Fields("COMEN")) Then tLi.SubItems(11) = tRs.Fields("COMEN")
                tRs.MoveNext
            Loop
        Else
            MsgBox "NO EXISTE DATOS DE LA VENTA", vbInformation, "SACC"
        End If
    End If
End Sub
