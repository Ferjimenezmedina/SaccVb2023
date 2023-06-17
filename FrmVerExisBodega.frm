VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmVerExisBodega 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Existencia en Bodega"
   ClientHeight    =   5190
   ClientLeft      =   660
   ClientTop       =   2010
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Traspasos"
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
      Left            =   10800
      Picture         =   "FrmVerExisBodega.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10800
      TabIndex        =   5
      Top             =   3840
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmVerExisBodega.frx":29D2
         MousePointer    =   99  'Custom
         Picture         =   "FrmVerExisBodega.frx":2CDC
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
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmVerExisBodega.frx":4DBE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   480
         Width           =   6975
      End
      Begin VB.CommandButton Command2 
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
         Left            =   9480
         Picture         =   "FrmVerExisBodega.frx":4DDA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   7011
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
         Caption         =   "Clave del Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmVerExisBodega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim SelMod As String
Dim NoPed As Integer
Dim CantPed As Double
Dim CantSurt As Double
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Command1_Click()
    frmtrassucursal.Show vbModal
End Sub
Private Sub Command2_Click()
    Buscar
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Me.Command2.Enabled = False
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
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 7000
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Subtotal", 1500
        .ColumnHeaders.Add , , "IVA", 1500
        .ColumnHeaders.Add , , "Total", 1500
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_PRODUCTO, CANTIDAD, DESCRIPCION, GANANCIA, PRECIO_COSTO FROM VSINVALM3 WHERE CANTIDAD > 0 AND SUCURSAL = 'BODEGA' AND ID_PRODUCTO LIKE '%" & Text1.Text & "%'"
    Set tRs = cnn.Execute(sBuscar)
    If (tRs.BOF And tRs.EOF) Then
        MsgBox "No se ha encontrado el producto"
    Else
        ListView1.ListItems.Clear
        tRs.MoveFirst
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO") & "")
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion") & ""
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD") & ""
            If Not IsNull(tRs.Fields("GANANCIA")) And Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(3) = Format((1 + CDbl(tRs.Fields("GANANCIA"))) * CDbl(tRs.Fields("PRECIO_COSTO")), "###,###,##0.00")
            If Not IsNull(tRs.Fields("GANANCIA")) And Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(4) = Format(((1 + CDbl(tRs.Fields("GANANCIA"))) * CDbl(tRs.Fields("PRECIO_COSTO"))) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
            If Not IsNull(tRs.Fields("GANANCIA")) And Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(5) = Format(((1 + CDbl(tRs.Fields("GANANCIA"))) * CDbl(tRs.Fields("PRECIO_COSTO"))) * (CDbl(CDbl(VarMen.Text4(7).Text) / 100) + 1), "###,###,##0.00")
            tRs.MoveNext
        Loop
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
        Me.Command2.Enabled = False
    Else
        Me.Command2.Enabled = True
    End If
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Buscar
    End If
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
