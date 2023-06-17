VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRepAlm1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion del Reporte de Almacen 1"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6720
      TabIndex        =   3
      Top             =   2880
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepAlm1.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepAlm1.frx":030A
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6720
      TabIndex        =   1
      Top             =   1560
      Width           =   975
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FrmRepAlm1.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepAlm1.frx":26F6
         Top             =   240
         Width           =   720
      End
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   -2147483644
      TabCaption(0)   =   "Reporte"
      TabPicture(0)   =   "FrmRepAlm1.frx":42C8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Option1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame Frame2 
         Caption         =   "Sucursal"
         Height          =   1095
         Left            =   3120
         TabIndex        =   9
         Top             =   2280
         Width           =   3255
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Text            =   "BODEGA"
            Top             =   480
            Width           =   2775
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   5741
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
      Begin VB.OptionButton Option2 
         Caption         =   "Precio de Compra y de Venta"
         Height          =   195
         Left            =   3600
         TabIndex        =   6
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Precio de Venta"
         Height          =   195
         Left            =   3600
         TabIndex        =   5
         Top             =   1080
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Marca :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmRepAlm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Command1_Click()
    Set crReport = crApplication.OpenReport(Path & "\REPORTES\JuegosDeReparacion.rpt")
    crReport.Database.LogOnServer "p2ssql.dll", GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "1"), GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER"), GetSetting("APTONER", "ConfigSACC", "USUARIO", "1"), GetSetting("APTONER", "ConfigSACC", "PASSWORD", "1")
    crReport.SQLQueryString = sBuscar
    crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    frmRep.Show vbModal, Me
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
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
        .ColumnHeaders.Add , , "MARCA", 2550
    End With
    sBuscar = "SELECT MARCA FROM ALMACEN1 GROUP BY MARCA ORDER BY MARCA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("MARCA"))
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT NOMBRE FROM SUCURSALES ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image1_Click()
On Error GoTo ManejaError
    If Combo1.Text <> "" Then
        Dim Cont As Integer
        Dim NoReg As Integer
        Dim Path As String
        Dim sBuscar As String
        Dim Checo As Integer
        Path = App.Path
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = 64
        CommonDialog1.ShowPrinter
        sBuscar = "SELECT * FROM VSEXISALMA1 WHERE "
        NoReg = ListView1.ListItems.Count
        For Cont = 1 To NoReg
            If ListView1.ListItems(Cont).Checked = True Then
                If Checo <> 0 Then
                    sBuscar = sBuscar & "AND "
                End If
                sBuscar = sBuscar & " MARCA = '" & ListView1.ListItems(Cont) & "' "
                Checo = 1
            End If
        Next Cont
        If Checo = 1 Then
            sBuscar = sBuscar & "AND SUCURSAL = '" & Combo1.Text & "' ORDER BY MARCA, ID_PRODUCTO"
        Else
            sBuscar = sBuscar & "SUCURSAL = '" & Combo1.Text & "' ORDER BY MARCA, ID_PRODUCTO"
        End If
    Else
        MsgBox "ES NECESATIO TOMAR UNA SUCURSAL PARA LA EXISTENCIA!", vbInformation, "SACC"
    End If
    If Option1.Value = True Then
        Set crReport = crApplication.OpenReport(Path & "\REPORTES\CARTVACVENTA.rpt")
        crReport.Database.LogOnServer "p2ssql.dll", GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "1"), GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER"), GetSetting("APTONER", "ConfigSACC", "USUARIO", "1"), GetSetting("APTONER", "ConfigSACC", "PASSWORD", "1")
        crReport.SQLQueryString = sBuscar
        crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        frmRep.Show vbModal, Me
    Else
        Set crReport = crApplication.OpenReport(Path & "\REPORTES\CARTVACCOMPRA.rpt")
        crReport.Database.LogOnServer "p2ssql.dll", GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "1"), GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER"), GetSetting("APTONER", "ConfigSACC", "USUARIO", "1"), GetSetting("APTONER", "ConfigSACC", "PASSWORD", "1")
        crReport.SQLQueryString = sBuscar
        crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        frmRep.Show vbModal, Me
    End If
    CommonDialog1.Copies = 1
    Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
