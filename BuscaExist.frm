VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form BuscaExist 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Existencias en Sucursales"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6480
      TabIndex        =   7
      Top             =   3360
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
         MouseIcon       =   "BuscaExist.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "BuscaExist.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "BuscaExist.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
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
         Left            =   4920
         Picture         =   "BuscaExist.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3960
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripcion"
         Height          =   195
         Left            =   4680
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   4680
         TabIndex        =   1
         Top             =   120
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   6015
         _ExtentX        =   10610
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
      Begin VB.Label Label1 
         Caption         =   "Buscar Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "BuscaExist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Command1_Click()
On Error GoTo ManejaError
    Buscar
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
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "SUCURSAL", 2300
        .ColumnHeaders.Add , , "EXISTENCIA", 1200
    End With
    Me.Command1.Enabled = False
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Buscar()
On Error GoTo ManejaError
    If Text1.Text = "" Then
        ListView1.ListItems.Clear
    Else
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        sBuscar = Text1.Text
        If Option1.Value Then
            sBuscar = "SELECT ID_PRODUCTO, SUCURSAL, CANTIDAD FROM VSEXISALMA3 WHERE ID_PRODUCTO LIKE '%" & Replace(sBuscar, "*", "%") & "%' ORDER BY Descripcion"
        Else
            sBuscar = "SELECT ID_PRODUCTO, SUCURSAL, CANTIDAD FROM VSEXISALMA3 WHERE Descripcion LIKE '%" & Replace(sBuscar, "*", "%") & "%' ORDER BY Descripcion"
        End If
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If Not (.BOF And .EOF) Then
                ListView1.ListItems.Clear
                .MoveFirst
                Do While Not .EOF
                    Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    tLi.SubItems(1) = .Fields("SUCURSAL") & ""
                    tLi.SubItems(2) = .Fields("CANTIDAD")
                    .MoveNext
                Loop
            Else
                ListView1.ListItems.Clear
                Set tLi = ListView1.ListItems.Add(, , "SIN RESULTADOS")
                    tLi.SubItems(1) = "........."
                    tLi.SubItems(2) = "........."
            End If
        End With
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
    If Text1.Text = "" Then
        Me.Command1.Enabled = False
    Else
        Me.Command1.Enabled = True
    End If
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Buscar
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890.* "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
