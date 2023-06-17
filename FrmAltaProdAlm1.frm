VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmAltaProdAlm1y2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alta de Productos Almacén 1 y 2"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   22
      Top             =   1200
      Width           =   975
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   49
         Top             =   0
         Width           =   975
         Begin VB.Image Image17 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmAltaProdAlm1.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "FrmAltaProdAlm1.frx":030A
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label20 
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
            TabIndex        =   50
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   47
         Top             =   1320
         Width           =   975
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "FrmAltaProdAlm1.frx":1DBC
            MousePointer    =   99  'Custom
            Picture         =   "FrmAltaProdAlm1.frx":20C6
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label19 
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
            TabIndex        =   48
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   45
         Top             =   1320
         Width           =   975
         Begin VB.Label Label14 
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
            TabIndex        =   46
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image15 
            Height          =   720
            Left            =   120
            MouseIcon       =   "FrmAltaProdAlm1.frx":38F0
            MousePointer    =   99  'Custom
            Picture         =   "FrmAltaProdAlm1.frx":3BFA
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Image Image18 
         Height          =   750
         Left            =   120
         MouseIcon       =   "FrmAltaProdAlm1.frx":55BC
         MousePointer    =   99  'Custom
         Picture         =   "FrmAltaProdAlm1.frx":58C6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label21 
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
         TabIndex        =   51
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   24
      Top             =   3600
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmAltaProdAlm1.frx":75F0
         MousePointer    =   99  'Custom
         Picture         =   "FrmAltaProdAlm1.frx":78FA
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
         TabIndex        =   41
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Almacén 2"
      Height          =   255
      Left            =   6120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Almacén 1"
      Height          =   255
      Left            =   5040
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   23
      Top             =   2400
      Width           =   975
      Begin VB.Label Label15 
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
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmAltaProdAlm1.frx":99DC
         MousePointer    =   99  'Custom
         Picture         =   "FrmAltaProdAlm1.frx":9CE6
         Top             =   240
         Width           =   675
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Buscar"
      TabPicture(0)   =   "FrmAltaProdAlm1.frx":B6A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Información"
      TabPicture(1)   =   "FrmAltaProdAlm1.frx":B6C4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label18"
      Tab(1).Control(1)=   "Label17"
      Tab(1).Control(2)=   "Label16"
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(7)=   "Label9"
      Tab(1).Control(8)=   "Label10"
      Tab(1).Control(9)=   "Label11"
      Tab(1).Control(10)=   "Label12"
      Tab(1).Control(11)=   "Label13"
      Tab(1).Control(12)=   "Precio_Venta"
      Tab(1).Control(13)=   "Label7"
      Tab(1).Control(14)=   "Label22"
      Tab(1).Control(15)=   "Label4"
      Tab(1).Control(16)=   "Combo3"
      Tab(1).Control(17)=   "Combo2"
      Tab(1).Control(18)=   "Text6"
      Tab(1).Control(19)=   "Combo1"
      Tab(1).Control(20)=   "Text3"
      Tab(1).Control(21)=   "Text8"
      Tab(1).Control(22)=   "Text9"
      Tab(1).Control(23)=   "Text10"
      Tab(1).Control(24)=   "Text11"
      Tab(1).Control(25)=   "Text12"
      Tab(1).Control(26)=   "Text13"
      Tab(1).Control(27)=   "cmdEnviar"
      Tab(1).Control(28)=   "Text4"
      Tab(1).Control(29)=   "Combo4"
      Tab(1).Control(30)=   "Combo5"
      Tab(1).Control(31)=   "Check1"
      Tab(1).ControlCount=   32
      Begin VB.CheckBox Check1 
         Caption         =   "Se puede hacer pedido manual de este producto"
         Height          =   255
         Left            =   -74760
         TabIndex        =   53
         Top             =   3960
         Width           =   4695
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "FrmAltaProdAlm1.frx":B6E0
         Left            =   -71280
         List            =   "FrmAltaProdAlm1.frx":B6E2
         TabIndex        =   19
         Top             =   3480
         Width           =   3255
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "FrmAltaProdAlm1.frx":B6E4
         Left            =   -74040
         List            =   "FrmAltaProdAlm1.frx":B6EE
         TabIndex        =   18
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   -70440
         TabIndex        =   13
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdEnviar 
         Caption         =   "Limpiar"
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
         Left            =   -69240
         Picture         =   "FrmAltaProdAlm1.frx":B702
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   -72120
         MaxLength       =   300
         TabIndex        =   12
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   -73920
         MaxLength       =   20
         TabIndex        =   11
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   -74040
         MaxLength       =   6
         TabIndex        =   16
         Text            =   "0"
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   -70440
         MaxLength       =   7
         TabIndex        =   17
         Text            =   "0"
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -70560
         MaxLength       =   20
         TabIndex        =   15
         Text            =   "<NINGUNO>"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -74040
         MaxLength       =   20
         TabIndex        =   14
         Text            =   "<NINGUNO>"
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -73800
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1080
         Width           =   5775
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   4200
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   4335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   5400
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   5400
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -69840
         TabIndex        =   10
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -73320
         MaxLength       =   20
         TabIndex        =   6
         Top             =   600
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -71280
         TabIndex        =   9
         Text            =   "S"
         Top             =   1560
         Width           =   735
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   -74280
         TabIndex        =   8
         Text            =   "SIMPLE"
         Top             =   1560
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5530
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
         Caption         =   "Especie"
         Height          =   255
         Left            =   -72120
         TabIndex        =   52
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label22 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69480
         TabIndex        =   44
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Precio en"
         Height          =   255
         Left            =   -74880
         TabIndex        =   43
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Precio_Venta 
         Caption         =   "P. Venta"
         Height          =   255
         Left            =   -71160
         TabIndex        =   42
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Color"
         Height          =   255
         Left            =   -71160
         TabIndex        =   40
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Material"
         Height          =   255
         Left            =   -74880
         TabIndex        =   39
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "P. Compra"
         Height          =   255
         Left            =   -72960
         TabIndex        =   38
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "% Ganancia"
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "C. Minima"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "C. Maxima"
         Height          =   255
         Left            =   -71280
         TabIndex        =   35
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "* Marca"
         Height          =   255
         Left            =   -70440
         TabIndex        =   34
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "* Venta WEB"
         Height          =   255
         Left            =   -72360
         TabIndex        =   33
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "* Descripción"
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccionado :"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "* Clave del Producto"
         Height          =   255
         Left            =   -74880
         TabIndex        =   29
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "* Tipo"
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Marca"
         Height          =   255
         Left            =   -70440
         TabIndex        =   27
         Top             =   1560
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmAltaProdAlm1y2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim IdProv As String
Dim AlmEli As String
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Sub cmdEnviar_Click()
    IdProv = ""
    Text2.Text = ""
    Text3.Text = ""
    Text11.Text = ""
    Text10.Text = ""
    Combo3.Text = ""
    Combo2.Text = ""
    Combo1.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text6.Text = ""
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Dim Resp As Long
    Resp = SendMessageLong(Combo1.hWnd, &H14F, True, 0)
    If KeyAscii = 13 Then
        Text12.SetFocus
    End If
End Sub
Private Sub Combo2_DropDown()
    Combo2.Clear
    Combo2.AddItem "S"
    Combo2.AddItem "N"
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo3_DropDown()
    Combo3.Clear
    Combo3.AddItem "SIMPLE"
    Combo3.AddItem "COMPUESTO"
End Sub
Private Sub Combo3_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo4_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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
        .ColumnHeaders.Add , , "Id Proveedor", 2500
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Cantidad Minima", 1500
        .ColumnHeaders.Add , , "Cantidad Maxima", 1500
        .ColumnHeaders.Add , , "Tipo", 0
        .ColumnHeaders.Add , , "Venta WEB", 0
        .ColumnHeaders.Add , , "Marca", 0
        .ColumnHeaders.Add , , "Material", 0
        .ColumnHeaders.Add , , "Color", 0
        .ColumnHeaders.Add , , "Ganancia", 0
        .ColumnHeaders.Add , , "Precio de Costo", 0
        .ColumnHeaders.Add , , "Precio de Venta", 0
        .ColumnHeaders.Add , , "Precio en", 0
        .ColumnHeaders.Add , , "Especie", 0
        .ColumnHeaders.Add , , "Pedido Sucursal", 0
    End With
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT MARCA FROM MARCA ORDER BY MARCA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("MARCA")) Then Combo1.AddItem tRs.Fields("MARCA")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT ESPECIE FROM ALMACEN1 GROUP BY ESPECIE ORDER BY ESPECIE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("ESPECIE")) Then Combo5.AddItem tRs.Fields("ESPECIE")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT ESPECIE FROM ALMACEN2 GROUP BY ESPECIE ORDER BY ESPECIE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("ESPECIE")) Then Combo5.AddItem tRs.Fields("ESPECIE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image18_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim MensajeSist As String
    sBuscar = "SELECT ID_REPARACION FROM JUEGO_REPARACION WHERE ID_PRODUCTO = '" & Text6.Text & "'"
    Set tRs = cnn.Execute(sBuscar)
    If tRs.EOF And tRs.BOF Then
        If AlmEli = "1" Then
            sBuscar = "DELETE FROM ALMACEN1 WHERE ID_PRODUCTO = '" & Trim(Text6.Text) & "'"
            cnn.Execute (sBuscar)
        Else
            sBuscar = "DELETE FROM ALMACEN2 WHERE ID_PRODUCTO = '" & Trim(Text6.Text) & "'"
            cnn.Execute (sBuscar)
        End If
        sBuscar = "DELETE FROM JUEGO_REPARACION WHERE ID_PRODUCTO = '" & Trim(Text6.Text) & "'"
        cnn.Execute (sBuscar)
        sBuscar = "DELETE FROM JR_TEMPORALES WHERE ID_PRODUCTO = '" & Trim(Text6.Text) & "'"
        cnn.Execute (sBuscar)
        sBuscar = "DELETE FROM JR_ALTERNOS WHERE ID_PRODUCTO1 = '" & Trim(Text6.Text) & "'"
        cnn.Execute (sBuscar)
        sBuscar = "DELETE FROM JR_ALTERNOS WHERE ID_PRODUCTO2 = '" & Trim(Text6.Text) & "'"
        cnn.Execute (sBuscar)
        Text6.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text11.Text = "0"
        Text10.Text = "0"
        Combo3.Text = "SIMPLE"
        Combo2.Text = "S"
        Combo1.Text = ""
        'Combo6.Text = ""
        Text8.Text = "<NINGUNO>"
        Text9.Text = "<NINGUNO>"
        Text12.Text = ""
        Text13.Text = ""
        'Text7.Text = ""
        'Combo6.Text = "ORIGINAL"
        'Combo5.Text = ""
    Else
        Do While Not tRs.EOF
            MensajeSist = tRs.Fields("ID_REPARACION") & ", "
            tRs.MoveNext
        Loop
        If MsgBox("EL PRODUCTO SE ENCONTRO EN LOS JUEGOS DE REPARACIÓN CON CLAVE " & MensajeSist & " ¿DESEA ELIMINAR EL PRODUCTO?, SE ELIMINARA TAMBIEN DEL JUEGO DE REPARACIÓN", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
            If AlmEli = "1" Then
                sBuscar = "DELETE FROM ALMACEN1 WHERE ID_PRODUCTO = '" & Trim(Text6.Text) & "'"
                cnn.Execute (sBuscar)
            Else
                sBuscar = "DELETE FROM ALMACEN2 WHERE ID_PRODUCTO = '" & Trim(Text6.Text) & "'"
                cnn.Execute (sBuscar)
            End If
            sBuscar = "DELETE FROM JUEGO_REPARACION WHERE ID_PRODUCTO = '" & Trim(Text6.Text) & "'"
            cnn.Execute (sBuscar)
            sBuscar = "DELETE FROM JR_TEMPORALES WHERE ID_PRODUCTO = '" & Trim(Text6.Text) & "'"
            cnn.Execute (sBuscar)
            sBuscar = "DELETE FROM JR_ALTERNOS WHERE ID_PRODUCTO1 = '" & Trim(Text6.Text) & "'"
            cnn.Execute (sBuscar)
            sBuscar = "DELETE FROM JR_ALTERNOS WHERE ID_PRODUCTO2 = '" & Trim(Text6.Text) & "'"
            cnn.Execute (sBuscar)
            Text6.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text11.Text = "0"
            Text10.Text = "0"
            Combo3.Text = "SIMPLE"
            Combo2.Text = "S"
            Combo1.Text = ""
            'Combo6.Text = ""
            Text8.Text = "<NINGUNO>"
            Text9.Text = "<NINGUNO>"
            Text12.Text = ""
            Text13.Text = ""
            'Text7.Text = ""
            'Combo6.Text = "ORIGINAL"
            'Combo5.Text = ""
        End If
    End If
    BusProd
End Sub
Private Sub Image8_Click()
    Dim sPedido As String
    If Text6.Text <> "" And Text3.Text <> "" Then
        If Check1.Value = 1 Then
            sPedido = "S"
        Else
            sPedido = "N"
        End If
        If Text4.Text <> "" Then
            If CDbl(Text4.Text) = 0 Then
                Text4.Text = ""
            End If
        End If
        If Text13.Text <> "" Then
            If CDbl(Text13.Text) = 0 Then
                Text13.Text = ""
            End If
        End If
        If Text12.Text <> "" Then
            If CDbl(Text12.Text) = 0 Then
                Text12.Text = ""
            End If
        End If
        If Text4.Text <> "" And Text13.Text <> "" And Text12.Text = "" Then
            Text12.Text = ((CDbl(Text4.Text) / CDbl(Text13.Text)) - 1) * 100
        End If
        If Text12.Text = "" Then
            If Option3.Value = True Then
                Text12.Text = "50"
            Else
                Text12.Text = "43"
            End If
        End If
        If Not (Text4.Text = "" And Text13.Text = "") Then
            If Text4.Text = "" Then
                Text4.Text = ((CDbl(Text12.Text) / 100) + 1) * CDbl(Text13.Text)
            Else
                Text13.Text = CDbl(Text4.Text) / ((CDbl(Text12.Text) / 100) + 1)
            End If
        Else
            MsgBox "DEBE DAR UN PRECIO DE COSTO O PRECIO DE VENTA PARA EL REGISTRO!", vbInformation, "SACC"
            Exit Sub
        End If
        Text12.Text = CDbl(Text12.Text) / 100
        If Combo4.Text = "" Then
            Combo4.Text = "PESOS"
        End If
        Dim tRs As ADODB.Recordset
        Text2.Text = Replace(Text2.Text, ",", "")
        Text3.Text = Replace(Text3.Text, ",", "")
        Text11.Text = Replace(Text11.Text, ",", "")
        Text10.Text = Replace(Text10.Text, ",", "")
        Combo3.Text = Replace(Combo3.Text, ",", "")
        Combo2.Text = Replace(Combo2.Text, ",", "")
        Combo1.Text = Replace(Combo1.Text, ",", "")
        Text8.Text = Replace(Text8.Text, ",", "")
        Text9.Text = Replace(Text9.Text, ",", "")
        Text12.Text = Replace(Format(Text12.Text, "0.00000"), ",", "")
        Text13.Text = Replace(Format(Text13.Text, "0.00000"), ",", "")
        Text6.Text = Replace(Text6.Text, ",", "")
        Dim sBuscar As String
        If Text10.Text = "" Then
            Text10.Text = "0"
        End If
        If Text11.Text = "" Then
            Text11.Text = "0"
        End If
        If Text12.Text = "" Then
            Text12.Text = "0.000001"
        End If
        If Option3.Value = True Then
            sBuscar = "SELECT ID_PRODUCTO FROM ALMACEN1 WHERE ID_PRODUCTO = '" & Trim(Text6.Text) & "'"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                If MsgBox("YA EXISTE UN PRODUCTO CON LA CLAVE " & Trim(IdProv) & "!, ¿DESEA GUARDAR LOS CAMBIOS?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
                    sBuscar = "UPDATE ALMACEN1 SET ID_PRODUCTO = '" & Text6.Text & "', Descripcion = '" & Text3.Text & "', TIPO = '" & Combo3.Text & "', VENTA_WEB = '" & Trim(Combo2.Text) & "', MARCA = '" & Combo1.Text & "', GANANCIA = " & Text12.Text & ", PRECIO_COSTO = '" & Text13.Text & "', MATERIAL = '" & Trim(Text8.Text) & "', COLOR = '" & Trim(Text9.Text) & "', C_MINIMA = '" & Text11.Text & "', C_MAXIMA = " & Text10.Text & ", PRECIO_VENTA = " & Text4.Text & ", PRECIO_EN = '" & Combo4.Text & "', USR_MOD = " & VarMen.Text1(0).Text & ", FECHA_MOD = '" & Format(Date, "dd/mm/yyyy") & "', ESPECIE = '" & Combo5.Text & "', PEDIDO_SUCURSAL = '" & sPedido & "' WHERE ID_PRODUCTO = '" & Trim(IdProv) & "'"
                    Set tRs = cnn.Execute(sBuscar)
                End If
            Else
                sBuscar = "INSERT INTO ALMACEN1 (ID_PRODUCTO, DESCRIPCION, TIPO, VENTA_WEB, MARCA, GANANCIA, PRECIO_COSTO, MATERIAL, COLOR, C_MINIMA, C_MAXIMA, PRECIO_VENTA, PRECIO_EN, USR_ALTA, FECHA_ALTA, ESPECIE, PEDIDO_SUCURSAL) VALUES ('" & Text6.Text & "', '" & Text3.Text & "', '" & Combo3.Text & "', '" & Combo2.Text & "', '" & Combo1.Text & "', '" & Text12.Text & "', '" & Text13.Text & "', '" & Text8.Text & "', '" & Text9.Text & "', '" & Text11.Text & "', '" & Text10.Text & "', " & Text4.Text & ", '" & Combo4.Text & "', " & VarMen.Text1(0).Text & ", '" & Format(Date, "dd/mm/yyyy") & "', '" & Combo5.Text & "', '" & sPedido & "');"
                cnn.Execute (sBuscar)
            End If
        Else
            sBuscar = "SELECT ID_PRODUCTO FROM ALMACEN2 WHERE ID_PRODUCTO = '" & Trim(Text6.Text) & "'"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                sBuscar = "UPDATE ALMACEN2 SET ID_PRODUCTO = '" & Text6.Text & "', Descripcion = '" & Text3.Text & "', TIPO = '" & Combo3.Text & "', VENTA_WEB = '" & Combo2.Text & "', MARCA = '" & Combo1.Text & "', GANANCIA = " & Text12.Text & ", PRECIO_COSTO = '" & Text13.Text & "', MATERIAL = '" & Text8.Text & "', COLOR = '" & Text9.Text & "', C_MINIMA = '" & Text11.Text & "', C_MAXIMA = " & Text10.Text & ", PRECIO_VENTA = " & Text4.Text & ", PRECIO_EN = '" & Combo4.Text & "', USR_MOD='" & VarMen.Text1(0).Text & "', FECHA_MOD = '" & Format(Date, "dd/mm/yyyy") & "', ESPECIE = '" & Combo5.Text & "', PEDIDO_SUCURSAL = '" & sPedido & "' WHERE ID_PRODUCTO = '" & Trim(IdProv) & "'"
                Set tRs = cnn.Execute(sBuscar)
            Else
                sBuscar = "INSERT INTO ALMACEN2 (ID_PRODUCTO, DESCRIPCION, TIPO, VENTA_WEB, MARCA, GANANCIA, PRECIO_COSTO, MATERIAL, COLOR, C_MINIMA, C_MAXIMA, PRECIO_VENTA, PRECIO_EN, USR_ALTA, FECHA_ALTA, ESPECIE, PEDIDO_SUCURSAL) VALUES ('" & Text6.Text & "', '" & Text3.Text & "', '" & Combo3.Text & "', '" & Combo2.Text & "', '" & Combo1.Text & "', '" & Text12.Text & "', '" & Text13.Text & "', '" & Text8.Text & "', '" & Text9.Text & "', '" & Text11.Text & "', '" & Text10.Text & "', " & Text4.Text & ", '" & Combo4.Text & "', " & VarMen.Text1(0).Text & ", '" & Format(Date, "dd/mm/yyyy") & "', '" & Combo5.Text & "', '" & sPedido & "');"
                cnn.Execute (sBuscar)
            End If
        End If
        IdProv = ""
        Text2.Text = ""
        Text3.Text = ""
        Text11.Text = "0"
        Text10.Text = "0"
        Combo3.Text = "SIMPLE"
        Combo2.Text = "S"
        Combo4.Text = ""
        Combo1.Text = ""
        Text8.Text = "<NINGUNO>"
        Text9.Text = "<NINGUNO>"
        Text12.Text = ""
        Text13.Text = ""
        Text4.Text = ""
        Text6.Text = ""
        Check1.Value = 1
        If Text1.Text <> "" Then
            BusProd
        End If
    Else
        MsgBox "DEBE DAR UNA Clave del Producto, Descripcion E INFORMACION DE PRECIOS PARA EL REGISTRO DEL PRODUCTO!", vbInformation, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdProv = Item
    Label22.Caption = ""
    Text2.Text = Item.SubItems(1)
    Text3.Text = Item.SubItems(1)
    Text11.Text = Item.SubItems(2)
    Text10.Text = Item.SubItems(3)
    Combo3.Text = Item.SubItems(4)
    Combo2.Text = Item.SubItems(5)
    Combo1.Text = Item.SubItems(6)
    Text8.Text = Item.SubItems(7)
    Text9.Text = Item.SubItems(8)
    Text12.Text = CDbl(Item.SubItems(9)) * 100
    Text13.Text = Item.SubItems(10)
    Text6.Text = Item
    Text4.Text = Item.SubItems(11)
    Combo4.Text = Item.SubItems(12)
    AlmEli = Item.SubItems(12)
    Combo5 = Item.SubItems(13)
    If Item.SubItems(14) = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    If Combo4.Text = "DOLARES" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        sBuscar = "SELECT COMPRA FROM DOLAR ORDER BY ID_DOLAR DESC"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Label22.Caption = Format(CDbl(Text4.Text) * CDbl(tRs.Fields("COMPRA")), "###,###,##0.00")
            Label22.Caption = "$ " & Label22.Caption & " M.N."
        Else
            MsgBox "NO SE TIENE PRECIO DEL DOLAR PARA CALCULAR EL PRECIO DE VENTA EN PESOS!", vbExclamation, "SACC"
        End If
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 And Text1.Text <> "" Then
        BusProd
    End If
End Sub
Private Sub BusProd()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If Option3.Value = True Then
        If Option1.Value = True Then
            sBuscar = "SELECT * FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' ORDER BY ID_PRODUCTO"
        Else
            sBuscar = "SELECT * FROM ALMACEN1 WHERE Descripcion LIKE '%" & Text1.Text & "%' ORDER BY Descripcion"
        End If
    Else
        If Option1.Value = True Then
            sBuscar = "SELECT * FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' ORDER BY ID_PRODUCTO"
        Else
            sBuscar = "SELECT * FROM ALMACEN2 WHERE Descripcion LIKE '%" & Text1.Text & "%' ORDER BY Descripcion"
        End If
    End If
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
                If Not IsNull(tRs.Fields("C_MINIMA")) Then tLi.SubItems(2) = tRs.Fields("C_MINIMA")
                If Not IsNull(tRs.Fields("C_MAXIMA")) Then tLi.SubItems(3) = tRs.Fields("C_MAXIMA")
                If Not IsNull(tRs.Fields("TIPO")) Then tLi.SubItems(4) = tRs.Fields("TIPO")
                If Not IsNull(tRs.Fields("VENTA_WEB")) Then tLi.SubItems(5) = tRs.Fields("VENTA_WEB")
                If Not IsNull(tRs.Fields("MARCA")) Then tLi.SubItems(6) = tRs.Fields("MARCA")
                If Not IsNull(tRs.Fields("MATERIAL")) Then tLi.SubItems(7) = tRs.Fields("MATERIAL")
                If Not IsNull(tRs.Fields("COLOR")) Then tLi.SubItems(8) = tRs.Fields("COLOR")
                If Not IsNull(tRs.Fields("GANANCIA")) Then tLi.SubItems(9) = tRs.Fields("GANANCIA")
                If Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(10) = tRs.Fields("PRECIO_COSTO")
                If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(11) = tRs.Fields("PRECIO_VENTA")
                If Not IsNull(tRs.Fields("PRECIO_EN")) Then tLi.SubItems(12) = tRs.Fields("PRECIO_EN")
                If Not IsNull(tRs.Fields("ESPECIE")) Then tLi.SubItems(13) = tRs.Fields("ESPECIE")
                If Not IsNull(tRs.Fields("PEDIDO_SUCURSAL")) Then tLi.SubItems(14) = tRs.Fields("PEDIDO_SUCURSAL")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text10_LostFocus()
    If Text10.Text = "" Then
        Text10.Text = "0"
    End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text11_LostFocus()
    If Text11.Text = "" Then
        Text11.Text = "0"
    End If
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNOPQRSTUVWXYZ -/1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNOPQRSTUVWXYZ -/1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNOPQRSTUVWXYZ -/1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text8_LostFocus()
    If Text8.Text = "" Then
        Text8.Text = "<NINGUNO>"
    End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNOPQRSTUVWXYZ -/1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text9_LostFocus()
    If Text9.Text = "" Then
        Text9.Text = "<NINGUNO>"
    End If
End Sub
