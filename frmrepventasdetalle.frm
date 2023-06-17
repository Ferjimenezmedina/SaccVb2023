VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmrepventasdetalle 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Detalle de ventas por PRODUCTO y Sucursal"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8670
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7560
      TabIndex        =   15
      Top             =   3120
      Width           =   975
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmrepventasdetalle.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmrepventasdetalle.frx":030A
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EXCEL"
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
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7560
      TabIndex        =   9
      Top             =   4320
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmrepventasdetalle.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "frmrepventasdetalle.frx":2156
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
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   11520
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   5160
      Width           =   150
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   16711680
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmrepventasdetalle.frx":4238
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdBuscar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Combo1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Combo2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.OptionButton Option2 
         Caption         =   "Bus-Detallada"
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bus-Producto"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3120
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmdBuscar 
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
         Left            =   720
         Picture         =   "frmrepventasdetalle.frx":4254
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fechas"
         Height          =   1215
         Left            =   4800
         TabIndex        =   1
         Top             =   120
         Width           =   2295
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   720
            TabIndex        =   2
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50921473
            CurrentDate     =   39576
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   720
            TabIndex        =   3
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50921473
            CurrentDate     =   39576
         End
         Begin VB.Label Label1 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Al :"
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   840
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   6165
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
         Caption         =   "Sucursal::"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Lista De Productos :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   1935
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Index           =   1
      Left            =   7800
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmrepventasdetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim StrRep As String
Private Sub cmdBuscar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    ListView1.ListItems.Clear
    sBuscar = "SELECT ID_VENTA,FOLIO,ID_CLIENTE,NOMBRE,FECHA,SUM(TOTAL) AS TOTAL,ID_PRODUCTO,Descripcion,SUM(CANTIDAD) AS CANTIDAD FROM vsrepdatalle WHERE ID_PRODUCTO LIKE '%" & Combo1.Text & "%' AND SUCURSAL = '" & Combo2.Text & "' AND   FECHA BETWEEN '" & DTPicker1.value & "' AND '" & DTPicker2.value & " ' GROUP BY ID_VENTA,FOLIO,ID_CLIENTE,NOMBRE,FECHA,TOTAL,ID_PRODUCTO,Descripcion,CANTIDAD"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
            If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(1) = tRs.Fields("FOLIO")
            tLi.SubItems(2) = tRs.Fields("ID_CLIENTE")
            tLi.SubItems(3) = tRs.Fields("NOMBRE")
             tLi.SubItems(4) = tRs.Fields("FECHA")
            tLi.SubItems(5) = tRs.Fields("TOTAL")
            tLi.SubItems(6) = tRs.Fields("ID_PRODUCTO")
            tLi.SubItems(7) = tRs.Fields("Descripcion")
            tLi.SubItems(8) = tRs.Fields("CANTIDAD")
           
           tRs.MoveNext
        Loop
      StrRep = sBuscar
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.value = Format(Date, "dd/mm/yyyy")
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
        .ColumnHeaders.Add , , "Id_Venta", 10
        .ColumnHeaders.Add , , "Folio", 10
        .ColumnHeaders.Add , , "Id_Cliente", 0
        .ColumnHeaders.Add , , "Nombre", 2000
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Total", 0
        .ColumnHeaders.Add , , "Clave", 1000
        .ColumnHeaders.Add , , "Descripcion", 1000
         ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "TOTAL", 0
    End With
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT * FROM vsrepdatalle " 'GROUP BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    Combo1.AddItem "<TODAS>"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("ID_PRODUCTO")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE "
    Set tRs = cnn.Execute(sBuscar)
    Combo2.AddItem "<TODAS>"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo2.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdBuscar.value = True
    End If
End Sub
