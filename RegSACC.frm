VERSION 5.00
Begin VB.Form RegSACC 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Frame Frame21 
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   3960
      TabIndex        =   2
      Top             =   1920
      Width           =   975
      Begin VB.Label Label21 
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Image imgLeer 
         Height          =   705
         Left            =   120
         MouseIcon       =   "RegSACC.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "RegSACC.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   5040
      TabIndex        =   0
      Top             =   1920
      Width           =   975
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image6 
         Height          =   870
         Left            =   120
         MouseIcon       =   "RegSACC.frx":1DBC
         MousePointer    =   99  'Custom
         Picture         =   "RegSACC.frx":20C6
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "www.jlbsystems.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      LinkItem        =   "www.jlbsystems.com"
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Para obtener el numero de registro acuda al sitio Web. Entre al apartado de registro llene sus datos y obtenga el numero."
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Registro :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Equipo :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "RegSACC.frx":41A8
      Top             =   120
      Width           =   6000
   End
End
Attribute VB_Name = "RegSACC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_NORMAL = 1
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public lngVolumeID As Long
Private Sub Form_Load()
On Error GoTo Borralo
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Dim nRet As Long
    Dim VolName As String
    Dim MaxCompLen As Long
    Dim VolFlags As Long
    Dim VolFileSys As String
    VolName = Space$(256)
    VolFileSys = Space$(256)
    nRet = GetVolumeInformation("C:\", VolName, Len(VolName), lngVolumeID, MaxCompLen, VolFlags, VolFileSys, Len(VolFileSys))
    Label3.Caption = lngVolumeID
    If GetSetting("APTONER", "ConfigSACC", "INTENTOS", "NO VALUE") = "0" Then
        Unload Me
    End If
Borralo:
    Err.Clear
End Sub
Private Sub Image6_Click()
    Unload Me
End Sub
Private Sub imgLeer_Click()
On Error GoTo Borralo
    Dim LonNum As Integer
    Dim NSRegistro As String
    Dim NSRegistro2 As String
    Dim StrFinal
    Dim ConInt As String
    Dim Con As Integer
    NSRegistro = Int(Label3.Caption / (Len(Label3.Caption) * 3.5))
    LonNum = Len(NSRegistro)
    NSRegistro = NSRegistro * LonNum
    LonNum = LonNum + 1
    For Con = 1 To Len(NSRegistro)
        StrFinal = StrFinal & Mid(NSRegistro, LonNum, 1)
        LonNum = LonNum - 1
    Next Con
    LonNum = Len(NSRegistro)
    LonNum = CDbl(LonNum) + 3
    StrFinal = StrFinal & "-" & CDbl(NSRegistro) * CDbl(LonNum)
    StrFinal = Replace(StrFinal, "0", "9")
    StrFinal = Replace(StrFinal, "1", "5")
    StrFinal = Replace(StrFinal, "2", "3")
    StrFinal = Replace(StrFinal, "3", "7")
    StrFinal = Replace(StrFinal, "4", "1")
    StrFinal = Replace(StrFinal, "5", "8")
    StrFinal = Replace(StrFinal, "6", "2")
    StrFinal = Replace(StrFinal, "7", "0")
    StrFinal = Replace(StrFinal, "8", "6")
    StrFinal = Replace(StrFinal, "9", "3")
    If Text1.Text = StrFinal Then
        Unload Me
        SaveSetting "APTONER", "ConfigSACC", "INTENTOS", "100"
        SaveSetting "APTONER", "ConfigSACC", "RegAprovSACC", "ValAprovReg"
        frmConfig.Show
    Else
        If GetSetting("APTONER", "ConfigSACC", "INTENTOS", "NO VALUE") = "NO VALUE" Then
            SaveSetting "APTONER", "ConfigSACC", "INTENTOS", "100"
        Else
            ConInt = CDbl(GetSetting("APTONER", "ConfigSACC", "INTENTOS", "NO VALUE")) - 1
            SaveSetting "APTONER", "ConfigSACC", "INTENTOS", ConInt
        End If
        MsgBox "EL NUMERO DE REGISTRO ES INCORRECTO!, TIENE DISPONIBLES " & GetSetting("APTONER", "ConfigSACC", "INTENTOS", "NO VALUE") & " INTENTOS DE REGISTRO", vbCritical, "SACC"
        If GetSetting("APTONER", "ConfigSACC", "INTENTOS", "NO VALUE") = "0" Then
            Unload Me
        End If
    End If
Borralo:
    Err.Clear
End Sub
Private Sub Label6_Click()
On Error GoTo Borralo
    Dim X
    X = ShellExecute(Me.hWnd, "Open", "http://www.jlbsystems.com", &O0, &O0, SW_NORMAL)
Borralo:
    Err.Clear
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo Borralo
    Dim Valido As String
    Valido = "1234567890-"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Borralo:
    Err.Clear
End Sub
