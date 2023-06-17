VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCompAlm1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capturar compra de producto (Almacen 1)"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9600
      TabIndex        =   32
      Top             =   5520
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCompAlm1.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCompAlm1.frx":030A
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
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9600
      TabIndex        =   29
      Top             =   1920
      Width           =   975
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Productos"
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
         TabIndex        =   30
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image12 
         Height          =   780
         Left            =   120
         MouseIcon       =   "FrmCompAlm1.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmCompAlm1.frx":26F6
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9600
      TabIndex        =   19
      Top             =   3120
      Width           =   975
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
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
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmCompAlm1.frx":46E8
         MousePointer    =   99  'Custom
         Picture         =   "FrmCompAlm1.frx":49F2
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9600
      TabIndex        =   15
      Top             =   4320
      Width           =   975
      Begin VB.Label Label5 
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
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FrmCompAlm1.frx":64A4
         MousePointer    =   99  'Custom
         Picture         =   "FrmCompAlm1.frx":67AE
         Top             =   240
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Enviar a revisión"
      TabPicture(0)   =   "FrmCompAlm1.frx":8380
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListView1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ListView2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Option1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Option2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text5"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtFolio"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtFCorrecto"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Combo1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "FrmCompAlm1.frx":839C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text6"
      Tab(1).Control(1)=   "Command2"
      Tab(1).Control(2)=   "txtInd"
      Tab(1).Control(3)=   "ListView3"
      Tab(1).Control(4)=   "Label8"
      Tab(1).ControlCount=   5
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6360
         TabIndex        =   31
         Text            =   "BODEGA"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtFCorrecto 
         Height          =   285
         Left            =   8640
         TabIndex        =   27
         Top             =   3360
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtFolio 
         Height          =   285
         Left            =   8280
         TabIndex        =   25
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -74160
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   6120
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Quitar"
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
         Left            =   -67080
         Picture         =   "FrmCompAlm1.frx":83B8
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox txtInd 
         Height          =   285
         Left            =   -68040
         TabIndex        =   21
         Top             =   6120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   4320
         TabIndex        =   18
         Top             =   6000
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   14
         Top             =   600
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   9551
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         Enabled         =   0   'False
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
         Left            =   8040
         Picture         =   "FrmCompAlm1.frx":AD8A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6000
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6600
         TabIndex        =   12
         Top             =   6000
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   6000
         Width           =   2895
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   6360
         TabIndex        =   4
         Top             =   3480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   6360
         TabIndex        =   3
         Top             =   3240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   3360
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   5055
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3836
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
         Height          =   2175
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3836
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label11 
         Caption         =   "Sucursal :"
         Height          =   255
         Left            =   6360
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Folio :"
         Height          =   255
         Left            =   8280
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Clave :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Precio :"
         Height          =   255
         Left            =   3720
         TabIndex        =   17
         Top             =   6000
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   5760
         TabIndex        =   11
         Top             =   6000
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Clave :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   9840
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "FrmCompAlm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim IdProv As String
Dim DesProd As String
Private Sub Combo1_DropDown()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Combo1.Clear
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
     If Not (tRs.EOF And tRs.BOF) Then
         Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo1_LostFocus()
    If Combo1.Text = "" Then
        Combo1.Text = "BODEGA"
    End If
End Sub
Private Sub Command1_Click()
    Dim sBuscar As String
    'sBuscar = "UPDATE ALMACEN1 SET PRECIO_COSTO = " & Text5.Text & " WHERE ID_PRODUCTO = '" & Text3.Text & "'"
    'cnn.Execute (sBuscar)
    If IdProv <> "" And Text3.Text <> "" And DesProd <> "" And Text4.Text <> "" And Text5.Text <> "" And Combo1.Text <> "" Then
        Dim tLi As ListItem
        Set tLi = ListView3.ListItems.Add(, , IdProv & "")
            tLi.SubItems(1) = Text3.Text & ""
            tLi.SubItems(2) = DesProd & ""
            tLi.SubItems(3) = Text4.Text & ""
            tLi.SubItems(4) = Text5.Text & ""
            tLi.SubItems(5) = Combo1.Text & ""
        Text3.Text = ""
        DesProd = ""
        Text4.Text = ""
        Text5.Text = ""
        Combo1.Enabled = False
    Else
        Dim CADENA As String
        If IdProv = "" Then
            CADENA = "PROVEEDOR"
        End If
        If Text3.Text = "" Then
            If CADENA <> "" Then
                CADENA = CADENA & ", "
            End If
            CADENA = CADENA & "PRODUCTO"
        End If
        If Text4.Text = "" Then
            If CADENA <> "" Then
                CADENA = CADENA & ", "
            End If
            CADENA = CADENA & "CANTIDAD"
        End If
        If Text5.Text = "" Then
            If CADENA <> "" Then
                CADENA = CADENA & ", "
            End If
            CADENA = CADENA & "PRECIO"
        End If
        If Combo1.Text = "" Then
            If CADENA <> "" Then
                CADENA = CADENA & " Y "
            End If
            CADENA = CADENA & "SUCURSAL"
        End If
        MsgBox "FALTA INFORMACION NECESARIA PARA EL REGISTRO (" & CADENA & ")", vbInformation, "SACC"
        CADENA = ""
    End If
    If ListView3.ListItems.COUNT > 0 Then
        Command2.Enabled = True
    Else
        Command2.Enabled = False
    End If
End Sub
Private Sub Command2_Click()
    ListView3.ListItems.Remove (CDbl(txtInd.Text))
    If ListView3.ListItems.COUNT > 0 Then
        Command2.Enabled = True
    Else
        Command2.Enabled = False
    End If
    txtInd.Text = ""
    Text6.Text = ""
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    Dim sBuscar As String
    Dim nGrupo As Integer
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
        .ColumnHeaders.Add , , "Id Proveedor", 0
        .ColumnHeaders.Add , , "Nombre", 7200
        .ColumnHeaders.Add , , "TELEFONO", 1850
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
        .ColumnHeaders.Add , , "Id Producto", 1500
        .ColumnHeaders.Add , , "Descripcion", 4550
        .ColumnHeaders.Add , , "Existencia", 1000
        .ColumnHeaders.Add , , "Maximo", 1000
        .ColumnHeaders.Add , , "Precio de Compra", 1000
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
        .ColumnHeaders.Add , , "Id Proveedor", 0
        .ColumnHeaders.Add , , "Id Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 4650
        .ColumnHeaders.Add , , "Cantidad", 1200
        .ColumnHeaders.Add , , "Precio de Compra", 1000
        .ColumnHeaders.Add , , "Sucursal", 1000
    End With
    sBuscar = "SELECT TOP 1 GRUPO FROM REV_COMPRA_ALMACEN1 ORDER BY ID_REVISION DESC"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If .EOF And .BOF Then
            txtFolio.Text = "1"
            txtFCorrecto.Text = "1"
        Else
            txtFolio.Text = Val(.Fields("GRUPO")) + 1
            txtFCorrecto.Text = Val(.Fields("GRUPO")) + 1
        End If
    End With
    If VarMen.Text1(21).Text = "N" Then
        Frame13.Visible = False
    End If
End Sub
Private Sub Image1_Click()
        Dim sBuscar As String
        Dim NRegistros As Integer
        Dim Con As Integer
        Dim Aux As String
        Dim nGrupo As Integer 'Para agrupar reporte reportes
        Dim Path As String
        Path = App.Path
        sBuscar = "SELECT GRUPO FROM REV_COMPRA_ALMACEN1 WHERE GRUPO = " & txtFolio.Text & " ORDER BY ID_REVISION DESC"
        Set tRs = cnn.Execute(sBuscar)
        If ListView3.ListItems.COUNT <> 0 Then
            With tRs
                If .EOF And .BOF Then
                    If txtFolio.Text = txtFCorrecto.Text Then
                        nGrupo = txtFolio.Text
                        NRegistros = ListView3.ListItems.COUNT
                        For Con = 1 To NRegistros
                            Aux = Replace(ListView3.ListItems(Con).SubItems(3), ",", "")
                            sBuscar = "INSERT INTO REV_COMPRA_ALMACEN1 (ID_PRODUCTO, ID_PROVEEDOR, CANTIDAD, FECHA, APROVADO, PRECIO_COMPRA, GRUPO, SUCURSAL, CANTIDAD_APROVADA) VALUES ('" & ListView3.ListItems(Con).SubItems(1) & "', " & ListView3.ListItems(Con) & ", " & Aux & ", '" & Format(Date, "dd/mm/yyyy") & "', 'P', " & ListView3.ListItems(Con).SubItems(4) & ", " & nGrupo & ", '" & Combo1.Text & "', " & Aux & ");"
                            cnn.Execute (sBuscar)
                        Next Con
                    Else
                        MsgBox "EL FOLIO DEBE SER: " & txtFCorrecto.Text & ", NO ES POSIBLE SALTEARSE FOLIOS"
                    End If
                Else
                    nGrupo = txtFolio.Text
                End If
            End With
        End If
        ImpRecep
        ImpRecep
        ListView3.ListItems.Clear
        Combo1.Text = ""
        Combo1.Enabled = True
End Sub
Private Sub Image12_Click()
    FrmAltaProdAlm1y2.Show vbModal
End Sub
Private Sub Image2_Click()
    FrmProvAlmace1.Show vbModal
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item.SubItems(1)
    IdProv = Item
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text3.Text = Item
    DesProd = Item.SubItems(1)
    Text5.Text = Item.SubItems(4)
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView3.ListItems.COUNT > 0 Then
        Text6.Text = Item.SubItems(1)
        txtInd.Text = Item.Index
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text1.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        ListView1.ListItems.Clear
        sBuscar = "SELECT ID_PROVEEDOR, NOMBRE, TELEFONO FROM PROVEEDOR_ALMACEN1 WHERE NOMBRE LIKE '%" & Text1.Text & "%' ORDER BY NOMBRE"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR") & "")
                    If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE") & ""
                    If Not IsNull(tRs.Fields("TELEFONO")) Then tLi.SubItems(2) = tRs.Fields("TELEFONO") & ""
                tRs.MoveNext
            Loop
        End If
    End If
    If Combo1.Text <> "" Then
        Command1.Enabled = True
    Else
        MsgBox ("ES NECESARIO SELECCIONAR UNA SUCURSAL")
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text2.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        ListView2.ListItems.Clear
        If Option1.value = True Then
            sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD, C_MAXIMA, PRECIO_COSTO, PRECIO_COSTO2 FROM vsInvAlm333 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%' AND SUCURSAL ='BODEGA' ORDER BY ID_PRODUCTO"
        Else
            sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD, C_MAXIMA, PRECIO_COSTO, PRECIO_COSTO2 FROM vsInvAlm333 WHERE Descripcion LIKE '%" & Text2.Text & "%' AND SUCURSAL ='BODEGA' ORDER BY ID_PRODUCTO"
        End If
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO") & "")
                    If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion") & ""
                    If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD") & ""
                    If Not IsNull(tRs.Fields("C_MAXIMA")) Then tLi.SubItems(3) = tRs.Fields("C_MAXIMA") & ""
                    If Not IsNull(tRs.Fields("CANTIDAD")) And Not IsNull(tRs.Fields("C_MAXIMA")) Then
                        If tRs.Fields("CANTIDAD") >= tRs.Fields("C_MAXIMA") Then
                            If Not IsNull(tRs.Fields("PRECIO_COSTO2")) Then tLi.SubItems(4) = tRs.Fields("PRECIO_COSTO2") & ""
                        Else
                            If Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(4) = tRs.Fields("PRECIO_COSTO") & ""
                        End If
                    Else
                        If Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(4) = tRs.Fields("PRECIO_COSTO") & ""
                    End If
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtFolio_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    If KeyAscii = 13 And txtFolio.Text <> "" Then
        Dim sBuscar As String
        Dim nGrupo As Integer 'Para agrupar reporte reportes
        Dim tLi As ListItem
        Dim tRs As ADODB.Recordset
        sBuscar = "SELECT R.GRUPO, R.ID_PRODUCTO, R.ID_PROVEEDOR, V.Descripcion, R.CANTIDAD, R.PRECIO_COMPRA, R.SUCURSAL FROM REV_COMPRA_ALMACEN1 AS R JOIN vsInvAlm333 AS V ON R.ID_PRODUCTO = V.ID_PRODUCTO WHERE GRUPO = " & txtFolio.Text & " AND V.SUCURSAL = '" & Combo1.Text & "' ORDER BY ID_REVISION DESC"
        Set tRs = cnn.Execute(sBuscar)
        ListView3.ListItems.Clear
        With tRs
            If .EOF And .BOF Then
                MsgBox "NO EXISTEN REGISTROS CON ESE FOLIO"
            Else
                Do While Not .EOF
                    Set tLi = ListView3.ListItems.Add(, , .Fields("ID_PROVEEDOR") & "")
                        tLi.SubItems(1) = .Fields("ID_PRODUCTO")
                        tLi.SubItems(2) = .Fields("Descripcion")
                        tLi.SubItems(3) = .Fields("CANTIDAD")
                        tLi.SubItems(4) = .Fields("PRECIO_COMPRA")
                        tLi.SubItems(5) = .Fields("SUCURSAL")
                    .MoveNext
                Loop
            End If
        End With
        SSTab1.Tab = 1
    End If
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub ImpRecep()
    Dim oDoc  As cPDF
    Dim Posi As Integer
    Dim sBuscar As String
    Dim tRs  As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Set oDoc = New cPDF
    Posi = 185
    Image4.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Not oDoc.PDFCreate(App.Path & "\CartVac.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Courier_Bold, MacRomanEncoding
' Encabezado del reporte
    oDoc.LoadImage Image4, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 80, 40, 43, 161, "Logo"
    sBuscar = "SELECT * FROM REV_COMPRA_ALMACEN1 WHERE GRUPO = " & txtFolio.Text & " ORDER BY ID_REVISION DESC"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F1", 7, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F1", 7, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F1", 7, hCenter
    oDoc.WTextBox 90, 200, 20, 250, "RECEPCION DE PRODUCTOS PARA REVISION", "F2", 10, hCenter
    oDoc.WTextBox 60, 400, 20, 250, "FECHA DE RECEPCION", "F2", 8, hCenter
    oDoc.WTextBox 70, 510, 20, 250, Format(Date, "dd/mm/yyyy"), "F2", 8, hLeft
' Encabezado de pagina
    sBuscar = "SELECT * FROM PROVEEDOR_ALMACEN1 WHERE ID_PROVEEDOR = " & tRs.Fields("ID_PROVEEDOR")
    Set tRs1 = cnn.Execute(sBuscar)
    oDoc.WTextBox 100, 10, 30, 400, "PROVEEDOR", "F2", 10, hLeft
    oDoc.WTextBox 110, 10, 30, 400, tRs1.Fields("NOMBRE"), "F2", 10, hLeft
    oDoc.WTextBox 130, 10, 30, 400, "TELEFONO", "F2", 10, hLeft
    oDoc.WTextBox 140, 10, 30, 400, tRs1.Fields("TELEFONO"), "F2", 10, hLeft
    oDoc.WTextBox 100, 450, 50, 400, "FOLIO: " & tRs.Fields("GRUPO"), "F2", 10, hLeft
    oDoc.WTextBox 130, 400, 40, 400, "SUCURSAL", "F2", 10, hLeft
    oDoc.WTextBox 140, 400, 40, 400, tRs.Fields("SUCURSAL"), "F2", 10, hLeft
' Cuerpo del reporte
    oDoc.WTextBox 170, 5, 40, 145, "CLAVE", "F2", 10, hLeft
    oDoc.WTextBox 170, 100, 40, 300, "DESCRIPCION", "F2", 10, hLeft
    oDoc.WTextBox 170, 350, 40, 70, "CANTIDAD", "F2", 10, hRight
    oDoc.WTextBox 170, 370, 40, 100, "P. UNIT.", "F2", 10, hRight
    oDoc.WTextBox 170, 470, 40, 80, "SUBTOTAL", "F2", 10, hRight
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 185
    oDoc.WLineTo 580, 185
    Do While Not tRs.EOF
        sBuscar = "SELECT Descripcion FROM ALMACEN1 WHERE ID_PRODUCTO = '" & Trim(tRs.Fields("ID_PRODUCTO")) & "'"
        Set tRs2 = cnn.Execute(sBuscar)
        If (tRs2.EOF And tRs2.BOF) Then
            sBuscar = "SELECT Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO = '" & Trim(tRs.Fields("ID_PRODUCTO")) & "'"
            Set tRs2 = cnn.Execute(sBuscar)
            If (tRs2.EOF And tRs2.BOF) Then
                sBuscar = "SELECT Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO = '" & Trim(tRs.Fields("ID_PRODUCTO")) & "'"
                Set tRs2 = cnn.Execute(sBuscar)
            End If
        End If
        oDoc.WTextBox Posi, 5, 40, 145, tRs.Fields("ID_PRODUCTO"), "F1", 9, hLeft
        oDoc.WTextBox Posi, 100, 9, 300, Mid(tRs2.Fields("Descripcion"), 1, 50), "F1", 9, hLeft
        oDoc.WTextBox Posi, 350, 40, 70, tRs.Fields("CANTIDAD"), "F1", 9, hRight
        oDoc.WTextBox Posi, 370, 40, 100, tRs.Fields("PRECIO_COMPRA"), "F1", 9, hRight
        oDoc.WTextBox Posi, 470, 40, 80, CDbl(tRs.Fields("CANTIDAD")) * CDbl(tRs.Fields("PRECIO_COMPRA")), "F1", 9, hRight
        If Posi > 700 Then
            Posi = 185
            oDoc.NewPage A4_Vertical
            oDoc.WImage 80, 40, 43, 161, "Logo"
            oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
            oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F1", 7, hCenter
            oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F1", 7, hCenter
            oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F1", 7, hCenter
            oDoc.WTextBox 90, 200, 20, 250, "RECEPCION DE PRODUCTOS PARA REVISION", "F2", 10, hCenter
            oDoc.WTextBox 60, 400, 20, 250, "FECHA DE RECEPCION", "F2", 8, hCenter
            oDoc.WTextBox 70, 510, 20, 250, Format(Date, "dd/mm/yyyy"), "F2", 8, hLeft
        ' Encabezado de pagina
            sBuscar = "SELECT * FROM PROVEEDOR_ALMACEN1 WHERE ID_PROVEEDOR = " & tRs.Fields("ID_PROVEEDOR")
            Set tRs1 = cnn.Execute(sBuscar)
            oDoc.WTextBox 100, 10, 30, 400, "PROVEEDOR", "F2", 10, hLeft
            oDoc.WTextBox 110, 10, 30, 400, tRs1.Fields("NOMBRE"), "F2", 10, hLeft
            oDoc.WTextBox 130, 10, 30, 400, "TELEFONO", "F2", 10, hLeft
            oDoc.WTextBox 140, 10, 30, 400, tRs1.Fields("TELEFONO"), "F2", 10, hLeft
            oDoc.WTextBox 100, 450, 50, 400, "FOLIO: " & tRs.Fields("GRUPO"), "F2", 10, hLeft
            oDoc.WTextBox 130, 400, 40, 400, "SUCURSAL", "F2", 10, hLeft
            oDoc.WTextBox 140, 400, 40, 400, tRs.Fields("SUCURSAL"), "F2", 10, hLeft
        ' Cuerpo del reporte
            oDoc.WTextBox 170, 5, 40, 145, "CLAVE", "F2", 10, hLeft
            oDoc.WTextBox 170, 100, 40, 300, "DESCRIPCION", "F2", 10, hLeft
            oDoc.WTextBox 170, 350, 40, 70, "CANTIDAD", "F2", 10, hRight
            oDoc.WTextBox 170, 370, 40, 100, "P. UNIT.", "F2", 10, hRight
            oDoc.WTextBox 170, 470, 40, 80, "SUBTOTAL", "F2", 10, hRight
            oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 185
            oDoc.WLineTo 580, 185
        End If
        tRs.MoveNext
        Posi = Posi + 10
    Loop
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, Posi + 20
    oDoc.WLineTo 580, Posi + 20
    oDoc.WTextBox 690, 5, 60, 580, "EL SR(A). " & tRs1.Fields("NOMBRE") & " CON DOMICILIO " & tRs1.Fields("DIRECCION") & " QUIEN SE IDENTIFICA CON " & tRs1.Fields("IDENTIFICACION") & " CON FOTOGRAFIA NO. " & tRs1.Fields("NUMERO_ID") & " VENDE A " & VarMen.TxtEmp(0).Text & " LOS ARTICULOS ARRIBA MENCIONADOS, MANIFESTANDO BAJO PROTESTA DE DECIR LA VERDAD QUE LA MERCANCIA ANTES DESCRITA SON DE SU EXCLUSIVA CONCESION Y DOMINIO, QUE LAS MISMAS FUERON ADQUIRIDAS CON SU PROPIO PECULIO DE MANERA LEGAL SIN PROCEDER DE HECHO ILICITO, POR LO QUE LIBERO DE CUALQUIER RESPONSABIIDAD PENAL O CIVIL AL COMPRADOR.", "F1", 10, hLeft
    oDoc.WTextBox 750, 5, 10, 580, "CLIENTE " & tRs1.Fields("NOMBRE") & " " & VarMen.Text4(3).Text & ", " & VarMen.Text4(4).Text & " A " & Format(Date, "dd/mm/yyyy"), "F1", 10, hLeft
    oDoc.WTextBox 770, 5, 10, 580, "ESTE DOCUMENTO AMPARA LA RECEPCION  DE LO CARTUCHOS ANTERIORMENTE DESCRITOS, Y NO GARANTIZA LA COMPRA TOTAL O PARCIAL DE LOS MISMOS. EL PRECIO MOSTRADO ES UN ESTIMADO DEL PRECIO MAXIMO DE COMPRA, Y ESTA SUJETO A CAMBIOS POR MOTIVOS DE FUNCIONAMIENTO O ESTADO DE LOS CARTUCHOS.", "F1", 10, hLeft
    oDoc.WTextBox 790, 5, 10, 580, "SE LE RECUERDA A LOS PROVEEDORES QUE EL HORARIO DE RECEPCION DE CARTUCHOS VACIOS ES DE LUNES A VIERNES DE 8:00 A 12:00 Y DE 4:00 A 5:30.", "F1", 10, hLeft
    oDoc.LineStroke
    oDoc.PDFClose
    oDoc.Show
End Sub
