VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmInv 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INVENTARIOS"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5040
      TabIndex        =   6
      Top             =   4920
      Width           =   975
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Image cmdCancelar 
         Height          =   705
         Left            =   120
         MouseIcon       =   "frmInv.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmInv.frx":030A
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5040
      TabIndex        =   4
      Top             =   6120
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmInv.frx":1DBC
         MousePointer    =   99  'Custom
         Picture         =   "frmInv.frx":20C6
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmInv.frx":41A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvwInv"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdInv"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdInv 
         Caption         =   "Inventario"
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
         Left            =   1800
         Picture         =   "frmInv.frx":41C4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6720
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwInv 
         Height          =   5775
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   10186
         View            =   3
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
         Alignment       =   2  'Center
         Caption         =   "LISTA DE PRODUCTOS"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim ItMx As ListItem
Public pro As String
Sub Llenar_Lista_Existencias_Sucursal(Clave As String)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT  ID_PRODUCTO, CANTIDAD From EXISTENCIAS WHERE SUCURSAL = '" & Clave & "'"
    Set tRs = cnn.Execute(sBuscar)
    Do While Not tRs.EOF
        Set ItMx = Me.lvwInv.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
        If Not IsNull(tRs.Fields("CANTIDAD")) Then ItMx.SubItems(1) = tRs.Fields("CANTIDAD")
        tRs.MoveNext
    Loop
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdInv_Click()
On Error GoTo ManejaError
    frmProInv.Show vbModal
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
    Me.Caption = "INVENTARIO DE " & frmSucInv.NSUC
    Llenar_Lista_Existencias_Sucursal frmSucInv.NSUC
    With lvwInv
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Producto", 2900
        .ColumnHeaders.Add , , "Cantidad", 1500
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    On Error GoTo ManejaError
    If bInvCre = True Then
        If MsgBox("CREO UN INVENTARIO NUEVO. SI CIERRA ESTA VENTANA TAMBIEN SE DARA POR TERMINADO EL INVENTARIO", vbYesNo, "SACC") = vbYes Then
            bInvCre = False
            Unload Me
        End If
    Else
        Unload Me
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub

Private Sub lvwInv_Click()
On Error GoTo ManejaError
    pro = Me.lvwInv.SelectedItem
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwInv_DblClick()
On Error GoTo ManejaError
    Me.cmdInv.Value = True
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
