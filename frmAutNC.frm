VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAutNC 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorizar Notas de Credito"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   8295
      Begin VB.TextBox txtTot 
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtNom 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtIDC 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtIDU 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtIDV 
         Height          =   285
         Left            =   840
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   600
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtImp 
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   4320
         TabIndex        =   8
         Top             =   2280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Autorizar"
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
         Left            =   6600
         Picture         =   "frmAutNC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Rechazar"
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
         Left            =   6600
         Picture         =   "frmAutNC.frx":29D2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   975
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Motivo de la nota:"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total: "
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   600
         Width           =   7695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nombre: "
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   240
         Width           =   7575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmAutNC.frx":53A4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.ListView ListView1 
         Height          =   2775
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   8295
         _ExtentX        =   14631
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
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8640
      TabIndex        =   4
      Top             =   4800
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmAutNC.frx":53C0
         MousePointer    =   99  'Custom
         Picture         =   "frmAutNC.frx":56CA
         Top             =   120
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmAutNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    Dim sBuscar As String
    sBuscar = "UPDATE SOLICITUD_NC SET AUTORIZADO = 'S' WHERE ID = " & txtID.Text
    cnn.Execute (sBuscar)
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    sBuscar = "UPDATE SOLICITUD_NC SET AUTORIZADO = 'S' WHERE ID = " & txtID.Text
    cnn.Execute (sBuscar)
    sBuscar = "INSERT INTO NOTA_CREDITO (IMPORTE, NOMBRE, TOTAL, FECHA, MOTIVOCAMBIO, ID_VENTA, ID_USUARIO, ID_CLIENTE) VALUES (" & txtImp.Text & ", '" & txtNom.Text & "', " & txtTot.Text & ", '" & txtFecha.Text & "', '" & Text1.Text & "', " & txtIDV.Text & ", '" & txtIDU.Text & "', " & txtIDC.Text & ");"
    cnn.Execute (sBuscar)
    'IMPRIMIR NOTA D CREDITO
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .CheckBoxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "NOMBRE", 3500
        .ColumnHeaders.Add , , "TOTAL", 1000
        .ColumnHeaders.Add , , "MOTIVO", 0
        .ColumnHeaders.Add , , "FECHA", 1440
        .ColumnHeaders.Add , , "IMPORTE", 0
        .ColumnHeaders.Add , , "ID_VENTA", 0
        .ColumnHeaders.Add , , "ID_USUARIO", 0
        .ColumnHeaders.Add , , "ID_CLIENTE", 0
    End With
End Sub
Private Sub Cargar()
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT * FROM SOLICITUD_NC WHERE AUTORIZADO = 'P'"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    Label1.Caption = "Nombre: "
    Label2.Caption = "Total: "
    Text1.Text = ""
    txtID.Text = ""
    txtImp.Text = ""
    txtFecha.Text = ""
    txtIDV.Text = ""
    txtIDU.Text = ""
    txtIDC.Text = ""
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID"))
                    tLi.SubItems(1) = .Fields("NOMBRE")
                    tLi.SubItems(2) = .Fields("TOTAL")
                    tLi.SubItems(3) = .Fields("MOTIVO")
                    tLi.SubItems(4) = .Fields("FECHA")
                    tLi.SubItems(5) = .Fields("IMPORTE")
                    tLi.SubItems(6) = .Fields("ID_VENTA")
                    tLi.SubItems(7) = .Fields("ID_USUARIO")
                    tLi.SubItems(8) = .Fields("ID_CLIENTE")
                .MoveNext
            Loop
        Else
            MsgBox "NO HAY NOTAS DE CREDITO PENDIENTES DE AUTORIZAR", vbInformation, "SACC"
        End If
    End With
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView1.ListItems.Count > 0 Then
        Label1.Caption = "Nombre: " & Item.SubItems(1)
        txtNom.Text = Item.SubItems(1)
        Label2.Caption = "Total: " & Item.SubItems(2)
        txtTot.Text = Replace(Item.SubItems(2), ",", ".")
        Text1.Text = Item.SubItems(3)
        txtID.Text = Item
        txtImp.Text = Replace(Item.SubItems(5), ",", ".")
        txtFecha.Text = Item.SubItems(4)
        txtIDV.Text = Item.SubItems(6)
        txtIDU.Text = Item.SubItems(7)
        txtIDC.Text = Item.SubItems(8)
    End If
End Sub
Private Sub txtID_Change()
    If txtID.Text = "" Then
        Command2.Enabled = False
        Command1.Enabled = False
    Else
        Command2.Enabled = True
        Command1.Enabled = True
    End If
End Sub
