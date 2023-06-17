VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmAltaProdAlm2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alta de Productos de Almacen 1"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   2
      Top             =   2280
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmAltaProdAlm2.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmAltaProdAlm2.frx":030A
         Top             =   240
         Width           =   675
      End
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   0
      Top             =   3480
      Width           =   975
      Begin VB.Image cmdCancelar 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmAltaProdAlm2.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "FrmAltaProdAlm2.frx":1FD6
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label14 
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Buscar"
      TabPicture(0)   =   "FrmAltaProdAlm2.frx":3A88
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Informacion"
      TabPicture(1)   =   "FrmAltaProdAlm2.frx":3AA4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdEnviar"
      Tab(1).Control(1)=   "Text13"
      Tab(1).Control(2)=   "Text12"
      Tab(1).Control(3)=   "Text11"
      Tab(1).Control(4)=   "Text10"
      Tab(1).Control(5)=   "Text9"
      Tab(1).Control(6)=   "Text8"
      Tab(1).Control(7)=   "Text7"
      Tab(1).Control(8)=   "Text3"
      Tab(1).Control(9)=   "Combo1"
      Tab(1).Control(10)=   "Text6"
      Tab(1).Control(11)=   "Combo2"
      Tab(1).Control(12)=   "Combo3"
      Tab(1).Control(13)=   "Label13"
      Tab(1).Control(14)=   "Label12"
      Tab(1).Control(15)=   "Label11"
      Tab(1).Control(16)=   "Label10"
      Tab(1).Control(17)=   "Label9"
      Tab(1).Control(18)=   "Label8"
      Tab(1).Control(19)=   "Label7"
      Tab(1).Control(20)=   "Label6"
      Tab(1).Control(21)=   "Label5"
      Tab(1).Control(22)=   "Label4"
      Tab(1).Control(23)=   "Label3"
      Tab(1).Control(24)=   "Label16"
      Tab(1).Control(25)=   "Label17"
      Tab(1).Control(26)=   "Label18"
      Tab(1).ControlCount=   27
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   -74400
         TabIndex        =   21
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -71400
         TabIndex        =   20
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -73440
         MaxLength       =   20
         TabIndex        =   19
         Top             =   600
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -69840
         TabIndex        =   18
         Top             =   1560
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   5400
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   5400
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4200
         Width           =   5535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -73920
         MaxLength       =   100
         TabIndex        =   13
         Top             =   1080
         Width           =   5895
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   -69240
         MaxLength       =   6
         TabIndex        =   12
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -74040
         MaxLength       =   20
         TabIndex        =   11
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -70560
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   -70440
         MaxLength       =   7
         TabIndex        =   9
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   -74040
         MaxLength       =   6
         TabIndex        =   8
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   -73920
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   -71640
         MaxLength       =   300
         TabIndex        =   6
         Top             =   2040
         Width           =   1215
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
         Left            =   -72000
         Picture         =   "FrmAltaProdAlm2.frx":3AC0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3960
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   22
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
      Begin VB.Label Label18 
         Caption         =   "Marca"
         Height          =   255
         Left            =   -70440
         TabIndex        =   38
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Clave del Producto"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccionado :"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Venta WEB"
         Height          =   255
         Left            =   -72360
         TabIndex        =   31
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Marca"
         Height          =   255
         Left            =   -70440
         TabIndex        =   30
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "P. de Venta"
         Height          =   255
         Left            =   -70200
         TabIndex        =   29
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "C. Maxima"
         Height          =   255
         Left            =   -71280
         TabIndex        =   28
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "C. Minima"
         Height          =   255
         Left            =   -74880
         TabIndex        =   27
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "% Ganancia"
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "P. Compra"
         Height          =   255
         Left            =   -72480
         TabIndex        =   25
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Material"
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Color"
         Height          =   255
         Left            =   -71160
         TabIndex        =   23
         Top             =   2520
         Width           =   375
      End
   End
End
Attribute VB_Name = "FrmAltaProdAlm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim IdProv As String
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
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
    Text7.Text = ""
    Text6.Text = ""
End Sub
Private Sub Combo1_DropDown()
    Combo1.Clear
    Dim sBuscar As String
    Dim tRs As Recordset
    sBuscar = "SELECT MARCA FROM MARCA ORDER BY MARCA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("MARCA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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
Private Sub Form_Load()
    Const sPathBase As String = "LINUX"
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Id Proveedor", 2500
        .ColumnHeaders.Add , , "Descripción", 5500
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
    End With
End Sub
Private Sub Image8_Click()
    Dim Ganan As String
    Dim PVent As String
    PVent = Format(CDbl(Text7.Text) / (1 + (CDbl(Text12.Text) / 100)), "0.00")
    Ganan = Format(CDbl(Text12.Text) / 100, "0.00")
    Text2.Text = Replace(Text2.Text, ",", ".")
    Text3.Text = Replace(Text3.Text, ",", ".")
    Text11.Text = Replace(Text11.Text, ",", ".")
    Text10.Text = Replace(Text10.Text, ",", ".")
    Combo3.Text = Replace(Combo3.Text, ",", ".")
    Combo2.Text = Replace(Combo2.Text, ",", ".")
    Combo1.Text = Replace(Combo1.Text, ",", ".")
    Text8.Text = Replace(Text8.Text, ",", ".")
    Text9.Text = Replace(Text9.Text, ",", ".")
    Text12.Text = Replace(Text12.Text, ",", ".")
    Text13.Text = Replace(Text13.Text, ",", ".")
    Text7.Text = Replace(Text7.Text, ",", ".")
    Text6.Text = Replace(Text6.Text, ",", ".")
    Ganan = Replace(Ganan, ",", ".")
    PVent = Replace(PVent, ",", ".")
    Dim sBuscar As String
    If IdProv <> "" Then
        Dim tRs As Recordset
        sBuscar = "UPDATE ALMACEN2 SET ID_PRODUCTO = '" & Text6.Text & "', DESCRIPCION = '" & Text3.Text & "', TIPO = '" & Combo3.Text & "', VENTA_WEB = '" & Combo2.Text & "', MARCA = '" & Combo1.Text & "', GANANCIA = " & Ganan & ", PRECIO_COSTO = '" & PVent & "', MATERIAL = '" & Text8.Text & "', COLOR = '" & Text9.Text & "', C_MINIMA = '" & Text11.Text & "', C_MAXIMA = " & Text10.Text & " WHERE ID_PRODUCTO = '" & IdProv & "'"
        Set tRs = cnn.Execute(sBuscar)
    Else
        sBuscar = "INSERT INTO ALMACEN2 (ID_PRODUCTO, DESCRIPCION, TIPO, VENTA_WEB, MARCA, GANANCIA, PRECIO_COSTO, MATERIAL, COLOR, C_MINIMA, C_MAXIMA) VALUES ('" & Text6.Text & "', '" & Text3.Text & "', '" & Combo3.Text & "', '" & Combo2.Text & "', '" & Combo1.Text & "', '" & Ganan & "', '" & PVent & "', '" & Text8.Text & "', '" & Text9.Text & "', '" & Text11.Text & "', '" & Text10.Text & "' );"
        cnn.Execute (sBuscar)
    End If
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
    Text7.Text = ""
    Text6.Text = ""
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdProv = Item
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
    Text7.Text = Item.SubItems(11)
    Text6.Text = Item
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 And Text1.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As Recordset
        Dim tLi As ListItem
        If Option1.Value = True Then
            sBuscar = "SELECT * FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' ORDER BY ID_PRODUCTO"
        Else
            sBuscar = "SELECT * FROM ALMACEN2 WHERE DESCRIPCION LIKE '%" & Text1.Text & "%' ORDER BY DESRIPCION"
        End If
        Set tRs = cnn.Execute(sBuscar)
        ListView1.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                    Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                    If Not IsNull(tRs.Fields("DESCRIPCION")) Then tLi.SubItems(1) = tRs.Fields("DESCRIPCION")
                    If Not IsNull(tRs.Fields("C_MINIMA")) Then tLi.SubItems(2) = tRs.Fields("C_MINIMA")
                    If Not IsNull(tRs.Fields("C_MAXIMA")) Then tLi.SubItems(3) = tRs.Fields("C_MAXIMA")
                    If Not IsNull(tRs.Fields("TIPO")) Then tLi.SubItems(4) = tRs.Fields("TIPO")
                    If Not IsNull(tRs.Fields("VENTA_WEB")) Then tLi.SubItems(5) = tRs.Fields("VENTA_WEB")
                    If Not IsNull(tRs.Fields("MARCA")) Then tLi.SubItems(6) = tRs.Fields("MARCA")
                    If Not IsNull(tRs.Fields("MATERIAL")) Then tLi.SubItems(7) = tRs.Fields("MATERIAL")
                    If Not IsNull(tRs.Fields("COLOR")) Then tLi.SubItems(8) = tRs.Fields("COLOR")
                    If Not IsNull(tRs.Fields("GANANCIA")) Then tLi.SubItems(9) = tRs.Fields("GANANCIA")
                    If Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(10) = tRs.Fields("PRECIO_COSTO")
                    If Not IsNull(tRs.Fields("GANANCIA")) And Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(11) = Format((1 + CDbl(tRs.Fields("GANANCIA"))) * CDbl(tRs.Fields("PRECIO_COSTO")), "0.00")
                tRs.MoveNext
            Loop
        End If
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
Private Sub Text11_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
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
Private Sub Text7_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub


