VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D9812703-90EB-11D2-887A-BD32CB08A467}#1.0#0"; "FldrView.ocx"
Begin VB.Form FrmRespalda 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Respaldo programado SACC"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9000
      MultiLine       =   -1  'True
      TabIndex        =   40
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   9000
      TabIndex        =   39
      Top             =   960
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   34
      Top             =   4800
      Width           =   975
      Begin VB.Label Label9 
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
         TabIndex        =   35
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRespalda.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRespalda.frx":030A
         ToolTipText     =   "Guardar cambios en la configuración del respaldo"
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   32
      ToolTipText     =   "Bajar a EXCEL el historial de respaldos"
      Top             =   1200
      Width           =   975
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
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRespalda.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "FrmRespalda.frx":1FD6
         ToolTipText     =   "Enviar a EXCEL reporte de respaldos"
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   24
      Top             =   2400
      Width           =   975
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   29
         Top             =   0
         Width           =   975
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmRespalda.frx":3B18
            MousePointer    =   99  'Custom
            Picture         =   "FrmRespalda.frx":3E22
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label12 
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
            TabIndex        =   30
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   27
         Top             =   1320
         Width           =   975
         Begin VB.Image Image4 
            Height          =   675
            Left            =   120
            MouseIcon       =   "FrmRespalda.frx":58D4
            MousePointer    =   99  'Custom
            Picture         =   "FrmRespalda.frx":5BDE
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label6 
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
            TabIndex        =   28
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   25
         Top             =   1320
         Width           =   975
         Begin VB.Label Label7 
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
            TabIndex        =   26
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image5 
            Height          =   720
            Left            =   120
            MouseIcon       =   "FrmRespalda.frx":7408
            MousePointer    =   99  'Custom
            Picture         =   "FrmRespalda.frx":7712
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Image Image2 
         Height          =   780
         Left            =   120
         MouseIcon       =   "FrmRespalda.frx":90D4
         MousePointer    =   99  'Custom
         Picture         =   "FrmRespalda.frx":93DE
         ToolTipText     =   "Limpiar historial de respaldos"
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame23 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   22
      Top             =   3600
      Width           =   975
      Begin VB.Image Image21 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FrmRespalda.frx":B3D0
         MousePointer    =   99  'Custom
         Picture         =   "FrmRespalda.frx":B6DA
         ToolTipText     =   "Crear el respaldo en este momento"
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Respaldar"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   6
      Top             =   6000
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRespalda.frx":D1E8
         MousePointer    =   99  'Custom
         Picture         =   "FrmRespalda.frx":D4F2
         ToolTipText     =   "Salir del sistema de respaldo"
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Archvo"
      TabPicture(0)   =   "FrmRespalda.frx":F5D4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DTPicker1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Calendar1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Opciones de Respaldo"
      TabPicture(1)   =   "FrmRespalda.frx":F5F0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame14"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame14 
         Caption         =   "Periodicidad"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   36
         Top             =   2880
         Width           =   2655
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "FrmRespalda.frx":F60C
            Left            =   240
            List            =   "FrmRespalda.frx":F62E
            TabIndex        =   37
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label11 
            Caption         =   "Dias"
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
            TabIndex        =   38
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Unidad 2 (Red/Local)"
         Height          =   2055
         Left            =   -72120
         TabIndex        =   20
         Top             =   4800
         Visible         =   0   'False
         Width           =   5535
         Begin FolderViewControl.FolderView FolderView2 
            Height          =   1695
            Left            =   840
            TabIndex        =   42
            Top             =   240
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   2990
            _StockProps     =   224
            Rootfolder      =   "FrmRespalda.frx":F653
         End
         Begin VB.Label Label5 
            Caption         =   "Folder"
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
            TabIndex        =   21
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Unidad 1 (Red/Local)"
         Height          =   2055
         Left            =   -72120
         TabIndex        =   18
         Top             =   2760
         Visible         =   0   'False
         Width           =   5535
         Begin FolderViewControl.FolderView FolderView1 
            Height          =   1695
            Left            =   840
            TabIndex        =   41
            Top             =   240
            Width           =   4575
            _Version        =   65536
            _ExtentX        =   8070
            _ExtentY        =   2990
            _StockProps     =   224
            Rootfolder      =   "FrmRespalda.frx":F66B
         End
         Begin VB.Label Label4 
            Caption         =   "Folder"
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
            TabIndex        =   19
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Correo electrónico"
         Height          =   2055
         Left            =   -72120
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1200
            TabIndex        =   16
            Top             =   720
            Width           =   4095
         End
         Begin VB.Label Label3 
            Caption         =   "Dirección"
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
            TabIndex        =   17
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Método de extracción"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   11
         Top             =   720
         Width           =   2655
         Begin VB.CheckBox Check3 
            Caption         =   "Unidad 2 (Red/Local)"
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
            TabIndex        =   14
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Unidad 1 (Red/Local)"
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
            TabIndex        =   13
            Top             =   960
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Correo electrónico"
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
            TabIndex        =   12
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Archivo"
         Height          =   3495
         Left            =   3960
         TabIndex        =   8
         Top             =   480
         Width           =   4695
         Begin VB.DirListBox Dir1 
            Height          =   2790
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   4455
         End
         Begin VB.Label Label2 
            Caption         =   "Ruta del archivo"
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
            TabIndex        =   10
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ultimos respaldos hchos"
         Height          =   3015
         Left            =   120
         TabIndex        =   5
         Top             =   3960
         Width           =   8535
         Begin MSComctlLib.ListView ListView1 
            Height          =   2535
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   4471
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
      Begin MSACAL.Calendar Calendar1 
         Height          =   2775
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   3615
         _Version        =   524288
         _ExtentX        =   6376
         _ExtentY        =   4895
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2010
         Month           =   7
         Day             =   9
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   3360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17170434
         CurrentDate     =   40368
      End
      Begin VB.Label Label1 
         Caption         =   "Hora del respaldo"
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
         TabIndex        =   4
         Top             =   3480
         Width           =   1815
      End
   End
   Begin VB.Timer TmrExisteArchivo 
      Interval        =   10000
      Left            =   9000
      Top             =   120
   End
End
Attribute VB_Name = "FrmRespalda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
' UDT para FindFirstFile
Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * 14
End Type
' Apis para buscar ficheros
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFilename As String, lpFindFileData As WIN32_FIND_DATA) As Long
' retorna Verdadero si el archivo existe
Public Function Existe(ByVal strFile As String) As Boolean
    Dim lHandle As Long             ' Handle del archivo
    Dim wFD As WIN32_FIND_DATA      ' udt con los datos
    ' Comprobar la barra separadora de path y la longitud
    If ((Len(strFile) > 3) And (Right$(strFile, 1) = "\")) Then
        strFile = Left$(strFile, Len(strFile) - 1)
    End If
    lHandle = FindFirstFile(strFile, wFD) ' buscar
    ' si el código del handle es válido ...
    Existe = lHandle <> INVALID_HANDLE_VALUE
    ' Liberar y cerrar el archivo con la función FindClose
    Call FindClose(lHandle)
End Function
Private Sub Check1_Click()
    If Check1.Value = 0 Then
        Frame5.Visible = False
    Else
        Frame5.Visible = True
    End If
End Sub
Private Sub Check2_Click()
    If Check2.Value = 0 Then
        Frame6.Visible = False
    Else
        Frame6.Visible = True
    End If
End Sub
Private Sub Check3_Click()
    If Check3.Value = 0 Then
        Frame7.Visible = False
    Else
        Frame7.Visible = True
    End If
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Calendar1.Value = Date
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Archivo", 1600
        .ColumnHeaders.Add , , "Tipo de respaldo", 3600
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Hora", 1500
    End With
    DTPicker1.Value = GetSetting("SACC", "RespaldoSACC", "HoraRespaldo", "12:00:00 a.m.")
    Dir1.Path = GetSetting("SACC", "RespaldoSACC", "RutaArchivo", "C:\")
    If GetSetting("SACC", "RespaldoSACC", "CorreoElectnonico", "NO") = "SI" Then
        Check1.Value = 1
        Text1.Text = GetSetting("SACC", "RespaldoSACC", "Direccion", "")
    Else
        Check1.Value = 0
    End If
    If GetSetting("SACC", "RespaldoSACC", "Unidad1", "NO") = "SI" Then
        Check2.Value = 1
        FolderView1.Tag = GetSetting("SACC", "RespaldoSACC", "Ruta1", "C:\")
    Else
        Check2.Value = 0
    End If
    If GetSetting("SACC", "RespaldoSACC", "Unidad2", "NO") = "SI" Then
        Check3.Value = 1
        FolderView2.Tag = GetSetting("SACC", "RespaldoSACC", "Ruta2", "C: \")
    Else
        Check3.Value = 0
    End If
    Combo1.Text = GetSetting("SACC", "RespaldoSACC", "Perioricidad", "0")
End Sub
Private Sub Image10_Click()
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Me.CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
    Me.CommonDialog1.ShowSave
    Text2.Text = ""
    Ruta = Me.CommonDialog1.FileName
    If ListView1.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.Count
            For Con = 1 To ListView1.ColumnHeaders.Count
                StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
            Next
            ProgressBar1.Value = 0
            ProgressBar1.Visible = True
            ProgressBar1.Min = 0
            ProgressBar1.Max = ListView1.ListItems.Count
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView1.ListItems.Count
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
                ProgressBar1.Value = Con
            Next
            Text2.Text = StrCopi
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, Text2.Text
            Close #foo
        End If
        ProgressBar1.Visible = False
        ProgressBar1.Value = 0
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
Private Sub Image2_Click()
    ListView1.ListItems.Clear
End Sub
Private Sub Image8_Click()
    SaveSetting "SACC", "RespaldoSACC", "HoraRespaldo", DTPicker1.Value
    SaveSetting "SACC", "RespaldoSACC", "RutaArchivo", Dir1.Path
    If Check1.Value = 1 Then
        SaveSetting "SACC", "RespaldoSACC", "CorreoElectnonico", "SI"
        SaveSetting "SACC", "RespaldoSACC", "Direccion", Text1.Text
    Else
        SaveSetting "SACC", "RespaldoSACC", "CorreoElectnonico", "NO"
    End If
    If Check2.Value = 1 Then
        SaveSetting "SACC", "RespaldoSACC", "Unidad1", "SI"
        SaveSetting "SACC", "RespaldoSACC", "Ruta1", FolderView1.Tag
    Else
        SaveSetting "SACC", "RespaldoSACC", "Unidad1", "NO"
    End If
    If Check3.Value = 1 Then
        SaveSetting "SACC", "RespaldoSACC", "Unidad2", "SI"
        SaveSetting "SACC", "RespaldoSACC", "Ruta2", FolderView2.Tag
    Else
        SaveSetting "SACC", "RespaldoSACC", "Unidad2", "NO"
    End If
    SaveSetting "SACC", "RespaldoSACC", "Perioricidad", Combo1.Text
    MsgBox "Configuración guardada!", vbExclamation, "Respaldo SACC"
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        Frame12.Visible = False
        Frame8.Visible = False
        Frame23.Visible = False
    Else
        Frame12.Visible = True
        Frame8.Visible = True
        Frame23.Visible = True
    End If
End Sub
Private Sub TmrExisteArchivo_Timer()
    Dim fso As Object
    'Instanciar el objeto FSO para poder _
    usar las funciones FileExists y FolderExists
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Comprobar archivo
    MsgBox fso.FileExists("c:\windows\notepad.exe")
    Set fso = Nothing
End Sub
Private Sub RespaldoCorreo()
    oMail.Sender = "sacc@jlbproducts.com"
    oMail.From = "Sistema SACC (Respaldo)"
End Sub
