VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmSustiInv 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Productos de Inventarios"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   8160
      TabIndex        =   24
      Text            =   "Text6"
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8160
      TabIndex        =   22
      Top             =   2640
      Width           =   975
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmSustiInv.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmSustiInv.frx":030A
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
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8160
      TabIndex        =   20
      Top             =   3960
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmSustiInv.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmSustiInv.frx":2156
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
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Sustituir"
      TabPicture(0)   =   "FrmSustiInv.frx":4238
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Sucursal"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Combo1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Ver Cambios"
      TabPicture(1)   =   "FrmSustiInv.frx":4254
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ListView2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Reporte"
      TabPicture(2)   =   "FrmSustiInv.frx":4270
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "DTPicker1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DTPicker2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "ListView4"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Command3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      Begin VB.CommandButton Command3 
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
         Left            =   -70920
         Picture         =   "FrmSustiInv.frx":428C
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   480
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   25
         Top             =   960
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   7011
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command2 
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
         Height          =   375
         Left            =   -68520
         Picture         =   "FrmSustiInv.frx":6C5E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   735
         Left            =   -74040
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   4200
         Width           =   5175
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   9
         Top             =   600
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   6165
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cambiar"
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
         Left            =   6120
         Picture         =   "FrmSustiInv.frx":9630
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6240
         TabIndex        =   7
         Top             =   3960
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cambiar por"
         Height          =   975
         Left            =   2760
         TabIndex        =   16
         Top             =   3840
         Width           =   2415
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tomar del producto"
         Height          =   975
         Left            =   240
         TabIndex        =   15
         Top             =   3840
         Width           =   2415
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   5640
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   5640
         TabIndex        =   2
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   1080
         Width           =   4455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   600
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   -72480
         TabIndex        =   26
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51380225
         CurrentDate     =   40016
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -74400
         TabIndex        =   27
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51380225
         CurrentDate     =   40016
      End
      Begin VB.Label Label5 
         Caption         =   "Del :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Al :"
         Height          =   255
         Left            =   -72840
         TabIndex        =   29
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Motivo :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "* Arrastre el producto a la casilla que corresponde."
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
         TabIndex        =   18
         Top             =   3480
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   5400
         TabIndex        =   17
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Sucursal 
         Caption         =   "Sucursal :"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmSustiInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim IdProd As String
Dim cant As String
Dim CanDE As String
Dim CanAL As String
Dim StrRep As String
Dim StrRep2 As String
Dim StrRep3 As String
Private Sub Combo1_DropDown()
    Me.Combo1.Clear
    Dim tRs As ADODB.Recordset
    Dim sBus As String
    sBus = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBus)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
    IdProd = ""
    cant = ""
    CanDE = ""
    CanAL = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
End Sub
Private Sub Command1_Click()
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim tLi As ListItem
    Dim exist As Integer
    If CDbl(CanDE) >= CDbl(Text4.Text) Then
        CanDE = CDbl(CanDE) - CDbl(Text4.Text)
        exist = CDbl(CanAL)
        CanAL = CDbl(CanAL) + CDbl(Text4.Text)
        sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CanDE & " WHERE ID_PRODUCTO = '" & Text2.Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
        cnn.Execute (sBuscar)
        If exist = 0 Then
            sBuscar = "INSERT INTO EXISTENCIAS(CANTIDAD, ID_PRODUCTO, SUCURSAL) VALUES (" & CanAL & ", '" & Text3.Text & "', '" & Combo1.Text & "');"
        Else
            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CanAL & " WHERE ID_PRODUCTO = '" & Text3.Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
        End If
        cnn.Execute (sBuscar)
        Set tLi = ListView2.ListItems.Add(, , Text2.Text & "")
        tLi.SubItems(1) = Text3.Text & ""
        tLi.SubItems(2) = Text4.Text & ""
        tLi.SubItems(3) = Combo1.Text & ""
        IdProd = ""
        cant = ""
        CanDE = ""
        CanAL = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        ListView1.ListItems.Clear
    Else
        MsgBox "EXISTENCIA INSUFICIENTE PARA HACER EL CAMBIO POR ESA CANTIDAD!", vbInformation, "SACC"
    End If
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim NumeroRegistros As Integer
    Dim Conta As Integer
    NumeroRegistros = ListView2.ListItems.COUNT
    For Conta = 1 To NumeroRegistros
        sBuscar = "INSERT INTO REEMP_INV (ID_PRODUCTO_DE, ID_PRODUCTO_PARA, SUCURSAL, CANTIDAD, MOTIVO, FECHA) VALUES ('" & ListView2.ListItems.Item(Conta) & "', '" & ListView2.ListItems.Item(Conta).SubItems(1) & "', '" & ListView2.ListItems.Item(Conta).SubItems(3) & "', " & ListView2.ListItems.Item(Conta).SubItems(2) & ", '" & Text5.Text & "', '" & Format(Date, "dd/mm/yyyy") & "');"
        cnn.Execute (sBuscar)
    Next
    Text5.Text = ""
    ListView2.ListItems.Clear
End Sub
Private Sub Command3_Click()
    ListView4.ListItems.Clear
    Dim sBuscar As String
    Dim NumeroRegistros As Integer
    Dim tLi As ListItem
    Dim Conta As Integer
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT * FROM REEMP_INV WHERE  FECHA BETWEEN '" & DTPicker1.value & "' AND '" & DTPicker2.value & " ' order BY  FECHA"
    Set tRs = cnn.Execute(sBuscar)
    StrRep = sBuscar
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView4.ListItems.Add(, , tRs.Fields("ID_REEMP"))
            tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO_DE")
            tLi.SubItems(2) = tRs.Fields("ID_PRODUCTO_PARA")
            tLi.SubItems(3) = tRs.Fields("CANTIDAD")
            tLi.SubItems(4) = tRs.Fields("MOTIVO")
            tLi.SubItems(5) = tRs.Fields("FECHA")
            tLi.SubItems(6) = tRs.Fields("SUCURSAL")
            tRs.MoveNext
           Loop
     End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Me.Command2.Enabled = False
    Me.Command1.Enabled = False
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    DTPicker1.value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.value = Format(Date, "dd/mm/yyyy")
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE", 1600
        .ColumnHeaders.Add , , "Descripcion", 4750
        .ColumnHeaders.Add , , "EXISTENCIA", 1000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "DEL PRODUCTO", 1600
        .ColumnHeaders.Add , , "AL PRODUCTO", 1600
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "SUCURSAL", 3150
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "# REMPLAZO", 1600
        .ColumnHeaders.Add , , "AL PRODD_PRODUCTO_DE", 1600
        .ColumnHeaders.Add , , "ID_PRODUCTO_PARA", 1000
        .ColumnHeaders.Add , , "CANTIDAD", 3150
        .ColumnHeaders.Add , , "MOTIVO", 1600
        .ColumnHeaders.Add , , "FECHA", 1600
        .ColumnHeaders.Add , , "SUCURSAL", 3150
    End With
End Sub
Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IdProd <> "" Then
        Text2.Text = IdProd
        IdProd = ""
        CanDE = cant
    End If
End Sub
Private Sub Frame2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IdProd <> "" Then
        Text3.Text = IdProd
        IdProd = ""
        CanAL = cant
    End If
End Sub
Private Sub Image10_Click()
    If ListView2.ListItems.COUNT > 0 Then
        Dim StrCopi As String
        Dim Con As Integer
        Dim Con2 As Integer
        Dim NumColum As Integer
        Dim Ruta As String
        Me.CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "Guardar como"
        CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
        Me.CommonDialog1.ShowSave
        Ruta = Me.CommonDialog1.FileName
        StrCopi = "ID_PRODUCTO_DEL" & Chr(9) & "ID_PRODUCTO_AL" & Chr(9) & "Cantidad" & Chr(9) & "Sucursal" & Chr(13)
        If Ruta <> "" Then
            NumColum = ListView2.ColumnHeaders.COUNT
            For Con = 1 To ListView2.ListItems.COUNT
                StrCopi = StrCopi & ListView2.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView2.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
    End If
    If ListView4.ListItems.COUNT > 0 Then
        Me.CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "Guardar como"
        CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
        Me.CommonDialog1.ShowSave
        Ruta = Me.CommonDialog1.FileName
        StrCopi = "# REMPLAZO" & Chr(9) & "ID_PRODUCTO_DE" & Chr(9) & "ID_PRODUCTO_PARA" & Chr(9) & "CANTIDAD" & Chr(9) & "MOTIVO" & Chr(9) & "FECHA" & Chr(9) & "SUCURSAL" & Chr(13)
        If Ruta <> "" Then
            NumColum = ListView4.ColumnHeaders.COUNT
            For Con = 1 To ListView4.ListItems.COUNT
                StrCopi = StrCopi & ListView4.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView4.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            'archivo TXT
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
    End If
End Sub
Private Sub Image9_Click()
    If ListView2.ListItems.COUNT = 0 Then
        Unload Me
    Else
        MsgBox "RECUERDA  GUARDA ANTES"
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdProd = Item
    cant = Item.SubItems(2)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Combo1.Text <> "" Then
            Dim tRs As ADODB.Recordset
            Dim sBuscar As String
            Dim tLi As ListItem
            ListView1.ListItems.Clear
            sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM VsExisALMACEN3Remp WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' AND NOMBRE = '" & Combo1.Text & "'" ' AND CANTIDAD >= 0"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.BOF And tRs.EOF) Then
                    tRs.MoveFirst
                    Do While Not tRs.EOF
                        Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO") & "")
                        If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion") & ""
                        If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD") & ""
                        tRs.MoveNext
                    Loop
            End If
        Else
            MsgBox "ES NECESARIO SELECCIONAR LA SUCURSAL!", vbInformation, "SACC"
        End If
    End If
End Sub
Private Sub Text2_Change()
    If Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
    End If
End Sub
Private Sub Text2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IdProd <> "" Then
        Text2.Text = IdProd
        IdProd = ""
        CanDE = cant
    End If
End Sub
Private Sub Text3_Change()
    If Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
    End If
End Sub
Private Sub Text3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IdProd <> "" Then
        Text3.Text = IdProd
        IdProd = ""
        CanAL = cant
    End If
End Sub
Private Sub Text4_Change()
    If Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
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
Private Sub Text5_Change()
    If ListView2.ListItems.COUNT <> 0 And Text5.Text <> "" Then
        Me.Command2.Enabled = True
    Else
        Me.Command2.Enabled = False
    End If
End Sub
