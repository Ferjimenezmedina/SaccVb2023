VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form PermisoVenta 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verificar Permiso"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5640
      TabIndex        =   5
      Top             =   1200
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
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "PermisoVenta.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "PermisoVenta.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4048
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "PermisoVenta.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Verificar"
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
         Left            =   3960
         Picture         =   "PermisoVenta.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Calve de la venta"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "1).- La clave funciona solo el día de hoy o al ser usada una ves, si la clave no funciona debe crear una clave nueva."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   4935
      End
   End
End
Attribute VB_Name = "PermisoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    Verifica
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Me.Command1.Enabled = False
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Verifica()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT CLAVE FROM VENTA_ESPECIAL WHERE CLAVE = '" & Text1.Text & "' AND FECHA = '" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    If tRs.EOF And tRs.BOF Then
        MsgBox "CLAVE INCORRECTA!", vbInformation, "SACC"
    Else
        sBuscar = "DELETE FROM VENTA_ESPECIAL WHERE CLAVE = '" & tRs.Fields("CLAVE") & "'"
        cnn.Execute (sBuscar)
        VentasEsp.Show vbModal
        Unload Me
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Text1_Change()
    If Text1.Text <> "" Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
    End If
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Verifica
    End If
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
