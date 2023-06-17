VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfig 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración del Acceso del Servidor"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Crear Respldos"
      Height          =   255
      Left            =   840
      TabIndex        =   20
      Top             =   5400
      Width           =   3615
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   18
      Top             =   5040
      Width           =   4575
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   16
      Top             =   2040
      Width           =   4575
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmConfig.frx":0000
      Left            =   840
      List            =   "frmConfig.frx":000D
      TabIndex        =   15
      Text            =   "SQLOLEDB.1"
      Top             =   4440
      Width           =   4575
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   2640
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   3840
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   3240
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   4575
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   6240
      TabIndex        =   6
      Top             =   4560
      Width           =   975
      Begin VB.Image Image6 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmConfig.frx":0032
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":033C
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label2 
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
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   6240
      TabIndex        =   4
      Top             =   3360
      Width           =   975
      Begin VB.Label Label8 
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   705
         Left            =   120
         MouseIcon       =   "frmConfig.frx":241E
         MousePointer    =   99  'Custom
         Picture         =   "frmConfig.frx":2728
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Base de Datos de Facturación"
      Height          =   255
      Left            =   840
      TabIndex        =   19
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre de la Empresa"
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Driver de la Base de Datos"
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cargando :"
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre de la Base de Datos"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contraseña de la Base de Datos"
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Usuario de la Base de Datos"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Line Line2 
      X1              =   6120
      X2              =   6120
      Y1              =   0
      Y2              =   5880
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6000
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre del Servidor"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmConfig.frx":41DA
      Top             =   120
      Width           =   6000
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCD As cgsFileOpR
Private cnn As ADODB.Connection
Const MAX_COMPUTERNAME_LENGTH = 255
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
     ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Function ComputerName() As String
    Dim sComputerName As String
    Dim ComputerNameLength As Long
    sComputerName = String(MAX_COMPUTERNAME_LENGTH + 1, 0)
    ComputerNameLength = MAX_COMPUTERNAME_LENGTH
    Call GetComputerName(sComputerName, ComputerNameLength)
    ComputerName = Mid(sComputerName, 1, ComputerNameLength)
End Function
Private Sub IniWrite(ByVal sFileName As String, ByVal sSection As String, ByVal sKeyName As String, ByVal sValue As String)
    Call WritePrivateProfileString(sSection, sKeyName, sValue, sFileName)
End Sub
Private Function AppPath(Optional ByVal ConBackSlash As Boolean = True) As String
    Dim s As String
    s = App.Path
    If ConBackSlash Then
        If Right$(s, 1) <> "\" Then
            s = s & "\"
        End If
    Else
        If Right$(s, 1) = "\" Then
            s = Left$(s, Len(s) - 1)
        End If
    End If
    AppPath = s
End Function
Private Sub Form_Load()
On Error GoTo Borralo
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Text1.Text = ComputerName
    Text1.Text = GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "")
    Text5.Text = GetSetting("APTONER", "ConfigSACC", "SERVIDORPPAL", "")
    Text3.Text = GetSetting("APTONER", "ConfigSACC", "PASSWORD", "")
    Text2.Text = GetSetting("APTONER", "ConfigSACC", "USUARIO", "")
    Text4.Text = GetSetting("APTONER", "ConfigSACC", "DATABASE", "")
    Combo1.Text = GetSetting("APTONER", "ConfigSACC", "PROVIDER", "SQLOLEDB.1")
    Text6.Text = GetSetting("APTONER", "ConfigSACC", "BDFacturacion", "FacturaGlobal")
    If GetSetting("APTONER", "ConfigSACC", "CreaRespaldo", "N") = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
Exit Sub
Borralo:
    Err.Clear
End Sub
Private Sub Image6_Click()
On Error GoTo Borralo
    If GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "0") <> "0" Then
        Unload Me
        NvoMen.Show
    Else
        If MsgBox("        ¿DESEA SALIR SIN ESPECIFICAR SERVIDOR?" & Chr(13) & "LA PROXIMA VES QUE INICIE DEBERA CONFIGURARLO", vbCritical + vbYesNo, "SACC") = vbYes Then
            Unload Me
        End If
    End If
Exit Sub
Borralo:
    Err.Clear
End Sub
Private Sub Image8_Click()
On Error GoTo ManejaError
    If Text1.Text <> "" Then
        SaveSetting "APTONER", "ConfigSACC", "SERVIDOR", Text1.Text
        SaveSetting "APTONER", "ConfigSACC", "SERVIDORPPAL", Text5.Text
        SaveSetting "APTONER", "ConfigSACC", "PASSWORD", Text3.Text
        SaveSetting "APTONER", "ConfigSACC", "USUARIO", Text2.Text
        SaveSetting "APTONER", "ConfigSACC", "DATABASE", Text4.Text
        SaveSetting "APTONER", "ConfigSACC", "PROVIDER", Combo1.Text
        SaveSetting "APTONER", "ConfigSACC", "BDFacturacion", Text6.Text
        If Check1.Value = 1 Then
            SaveSetting "APTONER", "ConfigSACC", "CreaRespaldo", "S"
        Else
            SaveSetting "APTONER", "ConfigSACC", "CreaRespaldo", "N"
        End If
        Set cnn = New ADODB.Connection
        With cnn
            .ConnectionString = _
                "Provider=" & Combo1.Text & ";Password=" & Text3.Text & ";Persist Security Info=True;User ID=" & Text2.Text & ";Initial Catalog=" & Text4.Text & ";Data Source=" & Text1.Text & ";"
                '"Provider=SQLNCLI10;SERVER=" & Text1.Text & ";User ID=" & Text2.Text & ";PWD=" & Text3.Text & ";Initial Catalog=" & Text4.Text & ""
            .Open
        End With
        Dim sFicINI As String
        Set mCD = New cgsFileOpR
        sFicINI = mCD.AddBackSlash(App.Path) & "Server.ini"
        IniWrite sFicINI, "Servidor", "Nombre", Text1.Text
        MsgBox "SE VA REINICIAR LA APLICACION", vbInformation, "SACC"
        Unload Me
    Else
        MsgBox "NO SE PUEDE GUARDAR SIN ESPECIFICAR UN SERVIDOR VALIDO", vbCritical, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
