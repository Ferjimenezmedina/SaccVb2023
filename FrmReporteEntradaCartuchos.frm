VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmReporteEntradaCartuchos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte entradas cartuchos "
   ClientHeight    =   2085
   ClientLeft      =   7245
   ClientTop       =   4980
   ClientWidth     =   4095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmReporteEntradaCartuchos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmReporteEntradaCartuchos.frx":030A
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3201
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmReporteEntradaCartuchos.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dtpA"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtpDe"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdVer"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton cmdVer 
         Caption         =   "Ver"
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
         Left            =   840
         Picture         =   "FrmReporteEntradaCartuchos.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpDe 
         Height          =   375
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   56426497
         CurrentDate     =   39094
      End
      Begin MSComCtl2.DTPicker dtpA 
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   56426497
         CurrentDate     =   39094
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "a:"
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
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "de:"
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
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
End
Attribute VB_Name = "FrmReporteEntradaCartuchos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Attribute cnn.VB_VarHelpID = -1
Dim sqlQuery As String
Dim tRs As Recordset
Private Sub cmdVer_Click()
    Dim sBuscar As String
    Dim sBuscar2 As String
    Dim sFecha1 As String
    Dim sFecha2 As String
    sFecha1 = dtpDe.Value
    sFecha2 = dtpA.Value
    Dim Path As String
    Path = App.Path
    sBuscar = "SELECT * FROM vsRepCartVac WHERE FECHA >= '" & sFecha1 & "'AND FECHA <= '" & sFecha2 & "'"
    Set crReport = crApplication.OpenReport(Path & "\REPORTES\RepCartVac.rpt")
    crReport.Database.LogOnServer "p2ssql.dll", GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "1"), GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER"), GetSetting("APTONER", "ConfigSACC", "USUARIO", "1"), GetSetting("APTONER", "ConfigSACC", "PASSWORD", "1")
    crReport.SQLQueryString = sBuscar
    crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    frmRep.Show vbModal, Me
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Me.dtpA.Value = Format(Date, "dd/mm/yyyy")
    Me.dtpDe.Value = Format(Date, "dd/mm/yyyy")
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
