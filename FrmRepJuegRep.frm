VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRepJuegRep 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Juegos de Reparación"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6960
      TabIndex        =   11
      Top             =   3240
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
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepJuegRep.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepJuegRep.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6960
      TabIndex        =   9
      Top             =   2040
      Width           =   975
      Begin VB.Label Label3 
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
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FrmRepJuegRep.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepJuegRep.frx":26F6
         Top             =   240
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Condiciones"
      TabPicture(0)   =   "FrmRepJuegRep.frx":42C8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame Frame3 
         Caption         =   "Ordenar Por "
         Height          =   1935
         Left            =   2760
         TabIndex        =   15
         Top             =   2160
         Width           =   3615
         Begin VB.OptionButton Option6 
            Caption         =   "Descripción"
            Height          =   255
            Left            =   1080
            TabIndex        =   6
            Top             =   960
            Width           =   1935
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Clave del Producto"
            Height          =   255
            Left            =   1080
            TabIndex        =   5
            Top             =   600
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Marca"
            Height          =   255
            Left            =   1080
            TabIndex        =   7
            Top             =   1320
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1815
         Left            =   360
         TabIndex        =   4
         Top             =   2280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame2 
         Caption         =   "Clave del Producto"
         Height          =   1335
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   6015
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1440
            TabIndex        =   3
            Top             =   600
            Width           =   4215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Termine en :"
            Height          =   195
            Left            =   240
            TabIndex        =   2
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Inicie con :"
            Height          =   195
            Left            =   240
            TabIndex        =   1
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Contenga :"
            Height          =   195
            Left            =   240
            TabIndex        =   0
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Marcas :"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2040
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmRepJuegRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Attribute cnn.VB_VarHelpID = -1
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With Me.ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .CheckBoxes = True
        .ColumnHeaders.Add , , "MARCA", 1700
    End With
    BusMarca
End Sub
Private Sub Image1_Click()
    Dim Path As String
    Dim sBuscar As String
    Dim cont As Integer
    Dim veces As Integer
    veces = 0
    Path = App.Path
    sBuscar = "SELECT VsJuegosReparacionREP.ID_PRODUCTO , VsJuegosReparacionREP.DESCRIPCION, VsJuegosReparacionREP.Marca, VsJuegosReparacionREP.ID_INSUMO, VsJuegosReparacionREP.DESC_INSUMO, VsJuegosReparacionREP.Precio_Venta, VsJuegosReparacionREP.CANTIDAD, VsJuegosReparacionREP.PRECIO_COSTO From APTONER.dbo.VsJuegosReparacionREP VsJuegosReparacionREP"
    If Option1.Value = True Then
        sBuscar = sBuscar & " where id_producto like '%" & Replace(Text1.Text, " ", "%") & "%'"
    End If
    If Option2.Value = True Then
        sBuscar = sBuscar & " where id_producto like '" & Text1.Text & "%'"
    End If
    If Option3.Value = True Then
        sBuscar = sBuscar & " where id_producto like '%" & Text1.Text & "'"
    End If
    For cont = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(cont).Checked = True Then
            If veces = 0 Then
                sBuscar = sBuscar & " and marca = '" & ListView1.ListItems.Item(cont) & "'"
                veces = 1
            Else
                sBuscar = sBuscar & " or marca = '" & ListView1.ListItems.Item(cont) & "'"
            End If
        End If
    Next cont
    If Option4.Value = True Then
        sBuscar = sBuscar & " Order By VsJuegosReparacionREP.MARCA ASC"
    End If
    If Option5.Value = True Then
        sBuscar = sBuscar & " Order By VsJuegosReparacionREP.ID_PRODUCTO ASC"
    End If
    If Option6.Value = True Then
        sBuscar = sBuscar & " Order By VsJuegosReparacionREP.DESCRIPCION ASC"
    End If
    Set crReport = crApplication.OpenReport(Path & "\REPORTES\JuegosDeReparacion.rpt")
    crReport.Database.LogOnServer "p2ssql.dll", GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "1"), GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER"), GetSetting("APTONER", "ConfigSACC", "USUARIO", "1"), GetSetting("APTONER", "ConfigSACC", "PASSWORD", "1")
    crReport.SQLQueryString = sBuscar
    crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    frmRep.Show vbModal, Me
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub BusMarca()
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT MARCA FROM ALMACEN3 GROUP BY MARCA ORDER BY MARCA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("MARCA"))
            tRs.MoveNext
        Loop
    End If
End Sub
