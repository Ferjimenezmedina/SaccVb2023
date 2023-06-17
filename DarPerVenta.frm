VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form DarPerVenta 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Dar Permiso de Venta Especial"
   ClientHeight    =   2535
   ClientLeft      =   6090
   ClientTop       =   4680
   ClientWidth     =   6735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   6735
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5640
      TabIndex        =   7
      Top             =   1200
      Width           =   975
      Begin VB.Label Label26 
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
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "DarPerVenta.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "DarPerVenta.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   120
      TabIndex        =   3
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
      TabPicture(0)   =   "DarPerVenta.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DTPicker1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
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
         Height          =   405
         Left            =   3960
         Picture         =   "DarPerVenta.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16252929
         CurrentDate     =   39258
      End
      Begin VB.Label Label1 
         Caption         =   "Clave de la venta"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de la Venta"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "1).- La clave funciona solo el día de hoy o al ser usada una vez, si esto pasara debe crear una clave nueva."
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
         TabIndex        =   4
         Top             =   1440
         Width           =   4935
      End
   End
End
Attribute VB_Name = "DarPerVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    sBuscar = "INSERT INTO VENTA_ESPECIAL (FECHA,CLAVE,CUMPLIDO) VALUES ('" & DTPicker1.Value & "', '" & Text1.Text & "', '0');"
    cnn.Execute (sBuscar)
    Text1.Text = ""
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    DTPicker1 = Format(Date, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Text1_Change()
    On Error GoTo ManejaError
    If Text1.Text = "" Then
        Me.Command1.Enabled = False
    Else
        Me.Command1.Enabled = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Command1.Value = True
    End If
    If Index = 0 Then
        Dim Valido As String
        Valido = "1234567890."
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
