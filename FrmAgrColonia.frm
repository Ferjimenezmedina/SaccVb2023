VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmAgrColonia 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Colonia"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5160
      TabIndex        =   6
      Top             =   1320
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmAgrColonia.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmAgrColonia.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Colonia"
      TabPicture(0)   =   "FrmAgrColonia.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Combo1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   0
         Top             =   840
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Registrar"
         Enabled         =   0   'False
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
         Left            =   3480
         Picture         =   "FrmAgrColonia.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "* Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "* Zona"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmAgrColonia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Combo1_Change()
On Error GoTo ManejaError
    If Combo1.Text <> "" And Text1.Text <> "" Then
        Me.Command1.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Combo1_Click()
    If Combo1.Text <> "" And Text1.Text <> "" Then
        Me.Command1.Enabled = True
    End If
End Sub
Private Sub Combo1_DropDown()
On Error GoTo ManejaError
    Combo1.Clear
    Combo1.AddItem "NORTE"
    Combo1.AddItem "SUR"
    Combo1.AddItem "ESTE"
    Combo1.AddItem "OESTE"
    Combo1.AddItem "CENTRO"
    Combo1.AddItem "NORESTE"
    Combo1.AddItem "NOROESTE"
    Combo1.AddItem "SURESTE"
    Combo1.AddItem "SUROESTE"
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Combo1_GotFocus()
On Error GoTo ManejaError
    Combo1.BackColor = &HFFE1E1
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.Command1.SetFocus
    End If
    KeyAscii = 0
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Combo1_LostFocus()
On Error GoTo ManejaError
    Combo1.BackColor = &H80000005
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
    If Combo1.Text <> "" And Text1.Text <> "" Then
        Me.Command1.Enabled = True
    End If
End Sub
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim sqlComanda As String
    Dim tRs As ADODB.Recordset
    sqlComanda = "SELECT ZONA FROM COLONIAS WHERE NOMBRE LIKE '" & Text1.Text & "'"
    Set tRs = cnn.Execute(sqlComanda)
    If (tRs.EOF And tRs.BOF) Then
        sqlComanda = "INSERT INTO COLONIAS (NOMBRE, ZONA) VALUES ('" & Text1.Text & "', '" & Combo1.Text & "');"
        cnn.Execute (sqlComanda)
    Else
        MsgBox "LA COLONIA " & Text1.Text & " YA ESTA REGISTRADA COMO DE LA ZONA " & tRs.Fields("ZONA"), vbInformation, "SACC"
    End If
    Text1.Text = ""
    Combo1.Text = ""
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
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
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Text1_Change()
On Error GoTo ManejaError
    If Combo1.Text <> "" And Text1.Text <> "" Then
        Me.Command1.Enabled = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_GotFocus()
On Error GoTo ManejaError
    Text1.BackColor = &HFFE1E1
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Combo1.SetFocus
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_LostFocus()
On Error GoTo ManejaError
    If Combo1.Text <> "" And Text1.Text <> "" Then
        Me.Command1.Enabled = True
    End If
    Text1.BackColor = &H80000005
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
