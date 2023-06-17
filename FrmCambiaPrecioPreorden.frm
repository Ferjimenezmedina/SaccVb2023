VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCambiaPrecioPreorden 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar Pre-Orden"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5160
      TabIndex        =   9
      Top             =   3240
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCambiaPrecioPreorden.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCambiaPrecioPreorden.frx":030A
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
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5160
      TabIndex        =   7
      Top             =   2040
      Width           =   975
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Modificar"
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
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image7 
         Height          =   810
         Left            =   120
         MouseIcon       =   "FrmCambiaPrecioPreorden.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmCambiaPrecioPreorden.frx":26F6
         Top             =   120
         Width           =   765
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Precio"
      TabPicture(0)   =   "FrmCambiaPrecioPreorden.frx":4820
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
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Proveedor"
      TabPicture(1)   =   "FrmCambiaPrecioPreorden.frx":483C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdBuscar"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin MSComctlLib.ListView ListView1 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   17
         Top             =   1080
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4895
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdBuscar 
         Cancel          =   -1  'True
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
         Left            =   -71280
         Picture         =   "FrmCambiaPrecioPreorden.frx":4858
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -74880
         TabIndex        =   15
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   11
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   " "
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
         Left            =   -74880
         TabIndex        =   18
         Top             =   3960
         Width           =   4695
      End
      Begin VB.Label Label8 
         Caption         =   "Nuevo Precio :"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label11 
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Precio :"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Clave del Producto :"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   2160
         TabIndex        =   1
         Top             =   1200
         Width           =   2055
      End
   End
   Begin VB.Label LblTipoOrden 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "FrmCambiaPrecioPreorden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim IdProveedor As String
Private Sub cmdBuscar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_PROVEEDOR, NOMBRE FROM PROVEEDOR WHERE NOMBRE LIKE '%" & Text2.Text & "%'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
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
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "Proveedor", 3500
    End With
End Sub
Private Sub Image7_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If SSTab1.Tab = 0 Then
        If Text1.Text <> "" Then
            sBuscar = "UPDATE COTIZA_REQUI SET PRECIO = " & CDbl(Replace(Text1.Text, ",", "")) & " WHERE ID_COTIZACION IN (" & Label10.Caption & ") AND ID_PROVEEDOR = " & Label11.Caption & " AND ID_PRODUCTO = '" & Label1.Caption & "'"
            Set tRs = cnn.Execute(sBuscar)
            Unload Me
        Else
            MsgBox "NO SE HA DADO PRECIO NUEVO!", vbInformation, "SACC"
        End If
    Else
        If IdProveedor <> "" And Label13.Caption <> "" Then
            sBuscar = "SELECT ID_ORDEN_COMPRA FROM ORDEN_COMPRA WHERE NUM_ORDEN  = " & Label13.Caption & " AND TIPO  = '" & LblTipoOrden & "'"
            Set tRs = cnn.Execute(sBuscar)
            sBuscar = "UPDATE COTIZA_REQUI SET ID_PROVEEDOR = " & IdProveedor & " WHERE NUMOC = " & tRs.Fields("ID_ORDEN_COMPRA")
            cnn.Execute (sBuscar)
            sBuscar = "UPDATE ORDEN_COMPRA SET ID_PROVEEDOR = " & IdProveedor & " WHERE ID_ORDEN_COMPRA = " & tRs.Fields("ID_ORDEN_COMPRA")
            cnn.Execute (sBuscar)
            Unload Me
        Else
            If Label13.Caption <> "" Then
                MsgBox "LA PRE-ORDEN DEBE TENER UN FOLIO PARA MODIFICAR ESTA OPCION!", vbInformation, "SACC"
            Else
                MsgBox "NO SE HA DADO UN NUEVO PROVEEDOR!", vbInformation, "SACC"
            End If
        End If
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdProveedor = Item
    Label12.Caption = Item.SubItems(1)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
