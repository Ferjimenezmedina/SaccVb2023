VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmEliProdAlm1y2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminar Productos Almacen 1 y 2"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   19
      Top             =   2280
      Width           =   975
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   24
         Top             =   1320
         Width           =   975
         Begin VB.Image Image15 
            Height          =   720
            Left            =   120
            MouseIcon       =   "FrmEliProdAlm1y2.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "FrmEliProdAlm1y2.frx":030A
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
            TabIndex        =   25
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   22
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
            TabIndex        =   23
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "FrmEliProdAlm1y2.frx":1CCC
            MousePointer    =   99  'Custom
            Picture         =   "FrmEliProdAlm1y2.frx":1FD6
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   20
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
            TabIndex        =   21
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image17 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmEliProdAlm1y2.frx":3800
            MousePointer    =   99  'Custom
            Picture         =   "FrmEliProdAlm1y2.frx":3B0A
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
         TabIndex        =   26
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image18 
         Height          =   750
         Left            =   120
         MouseIcon       =   "FrmEliProdAlm1y2.frx":55BC
         MousePointer    =   99  'Custom
         Picture         =   "FrmEliProdAlm1y2.frx":58C6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   17
      Top             =   3480
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmEliProdAlm1y2.frx":75F0
         MousePointer    =   99  'Custom
         Picture         =   "FrmEliProdAlm1y2.frx":78FA
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
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Buscar"
      TabPicture(0)   =   "FrmEliProdAlm1y2.frx":99DC
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
      TabCaption(1)   =   "Información"
      TabPicture(1)   =   "FrmEliProdAlm1y2.frx":99F8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo3"
      Tab(1).Control(1)=   "Combo2"
      Tab(1).Control(2)=   "Text6"
      Tab(1).Control(3)=   "Combo1"
      Tab(1).Control(4)=   "Text3"
      Tab(1).Control(5)=   "Text8"
      Tab(1).Control(6)=   "Text9"
      Tab(1).Control(7)=   "Text10"
      Tab(1).Control(8)=   "Text11"
      Tab(1).Control(9)=   "Text12"
      Tab(1).Control(10)=   "Text13"
      Tab(1).Control(11)=   "Text4"
      Tab(1).Control(12)=   "Label20"
      Tab(1).Control(13)=   "Label15"
      Tab(1).Control(14)=   "Label3"
      Tab(1).Control(15)=   "Label5"
      Tab(1).Control(16)=   "Label6"
      Tab(1).Control(17)=   "Label8"
      Tab(1).Control(18)=   "Label9"
      Tab(1).Control(19)=   "Label10"
      Tab(1).Control(20)=   "Label11"
      Tab(1).Control(21)=   "Label12"
      Tab(1).Control(22)=   "Label13"
      Tab(1).Control(23)=   "Precio_Venta"
      Tab(1).ControlCount=   24
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74400
         TabIndex        =   7
         Text            =   "SIMPLE"
         Top             =   1920
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -71400
         TabIndex        =   8
         Text            =   "S"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73320
         MaxLength       =   20
         TabIndex        =   5
         Top             =   960
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69840
         TabIndex        =   9
         Top             =   1920
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   5400
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   5400
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   4200
         Width           =   5535
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73800
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1440
         Width           =   5775
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74040
         MaxLength       =   20
         TabIndex        =   13
         Text            =   "<NINGUNO>"
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70560
         MaxLength       =   20
         TabIndex        =   14
         Text            =   "<NINGUNO>"
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70440
         MaxLength       =   7
         TabIndex        =   16
         Text            =   "0"
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox Text11 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74040
         MaxLength       =   6
         TabIndex        =   15
         Text            =   "0"
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox Text12 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73920
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72120
         MaxLength       =   300
         TabIndex        =   11
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70440
         TabIndex        =   12
         Top             =   2400
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   3
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
      Begin VB.Label Label20 
         Caption         =   "* Tipo"
         Height          =   255
         Left            =   -74880
         TabIndex        =   44
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "* Clave del Producto"
         Height          =   255
         Left            =   -74880
         TabIndex        =   43
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Marca"
         Height          =   255
         Left            =   -70440
         TabIndex        =   42
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   -74880
         TabIndex        =   41
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Clave del Producto"
         Height          =   255
         Left            =   -74880
         TabIndex        =   40
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccionado :"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "* Descripción"
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "* Venta WEB"
         Height          =   255
         Left            =   -72360
         TabIndex        =   36
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "* Marca"
         Height          =   255
         Left            =   -70440
         TabIndex        =   35
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "C. Maxima"
         Height          =   255
         Left            =   -71280
         TabIndex        =   34
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "C. Minima"
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "% Ganancia"
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "P. Compra"
         Height          =   255
         Left            =   -72960
         TabIndex        =   31
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Material"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Color"
         Height          =   255
         Left            =   -71160
         TabIndex        =   29
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Precio_Venta 
         Caption         =   "P. Venta"
         Height          =   255
         Left            =   -71160
         TabIndex        =   28
         Top             =   2400
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmEliProdAlm1y2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim AlmEli As String
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Dim Resp As Long
    Resp = SendMessageLong(Combo1.hWnd, &H14F, True, 0)
    If KeyAscii = 13 Then
        Text12.SetFocus
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
        .ColumnHeaders.Add , , "Clave del Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 4100
        .ColumnHeaders.Add , , "Minimo", 1000
        .ColumnHeaders.Add , , "Maximo", 1000
        .ColumnHeaders.Add , , "Tipo", 1000
        .ColumnHeaders.Add , , "Venta WEB", 1000
        .ColumnHeaders.Add , , "Marca", 1000
        .ColumnHeaders.Add , , "Material", 1000
        .ColumnHeaders.Add , , "Color", 1000
        .ColumnHeaders.Add , , "Ganancia", 1000
        .ColumnHeaders.Add , , "P. Costo", 1000
        .ColumnHeaders.Add , , "P. Venta", 1000
        .ColumnHeaders.Add , , "Almacen", 1000
    End With
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
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text6.Text = Item
    Text2.Text = Item
    Text3.Text = Item.SubItems(1)
    Combo3.Text = Item.SubItems(4)
    Combo2.Text = Item.SubItems(5)
    Combo1.Text = Item.SubItems(6)
    Text12.Text = CDbl(Item.SubItems(9)) * 100
    Text13.Text = Item.SubItems(10)
    Text4.Text = Item.SubItems(11)
    Text8.Text = Item.SubItems(7)
    Text9.Text = Item.SubItems(8)
    Text11.Text = Item.SubItems(2)
    Text10.Text = Item.SubItems(3)
    AlmEli = Item.SubItems(12)
    
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BusProd
    End If
End Sub
Private Sub BusProd()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If Option1.Value = True Then
        sBuscar = "SELECT * FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & Replace(Text1.Text, " ", "%") & "%'"
    Else
        sBuscar = "SELECT * FROM ALMACEN1 WHERE Descripcion LIKE '%" & Replace(Text1.Text, " ", "%") & "%'"
    End If
    ListView1.ListItems.Clear
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
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
            tLi.SubItems(12) = "1"
            tRs.MoveNext
        Loop
    End If
    If Option1.Value = True Then
        sBuscar = "SELECT * FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Replace(Text1.Text, " ", "%") & "%'"
    Else
        sBuscar = "SELECT * FROM ALMACEN2 WHERE Descripcion LIKE '%" & Replace(Text1.Text, " ", "%") & "%'"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
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
            tLi.SubItems(12) = "2"
            tRs.MoveNext
        Loop
    End If
End Sub
