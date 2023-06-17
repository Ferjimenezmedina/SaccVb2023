VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmProInv 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INVENTARIOS"
   ClientHeight    =   5670
   ClientLeft      =   7260
   ClientTop       =   5025
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5040
      TabIndex        =   8
      Top             =   4320
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
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmProInv.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmProInv.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmProInv.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvwInvSucPro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   120
         TabIndex        =   3
         Top             =   3120
         Width           =   4575
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
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
            Left            =   3240
            Picture         =   "frmProInv.frx":2408
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Text            =   "ESCRIBA AQUÍ LA CANTIDAD REAL"
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label lblInv 
            Alignment       =   2  'Center
            Caption         =   "INVENTARIO"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label lblInv2 
            Alignment       =   2  'Center
            Caption         =   "..."
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   1560
            Width           =   4335
         End
      End
      Begin MSComctlLib.ListView lvwInvSucPro 
         Height          =   2055
         Left            =   120
         TabIndex        =   0
         Top             =   960
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "LISTA DE SUCURSALES QUE CUENTAN  CON EL PRODUCTO"
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   480
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmProInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim ItMx As ListItem
Private Sub cmdOk_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If bInvCre = True Then
        sBuscar = "INSERT INTO INVENTARIO_DETALLE (ID_INVENTARIO, ID_PRODUCTO, CANTIDAD) Values (" & INV & ", '" & frmInv.pro & "', " & Val(Me.txtCantidad.Text) & ")"
        cnn.Execute (sBuscar)
    Else
        bInvCre = True
        sBuscar = "INSERT INTO INVENTARIOS (ID_SUCURSAL, FECHA) Values (" & frmSucInv.NSUC & ", '" & Format(Date, "dd/mm/yyyy") & "')"
        cnn.Execute (sBuscar)
        sBuscar = "SELECT TOP 1 ID_INVENTARIO From INVENTARIOS ORDER BY ID_INVENTARIO DESC"
        Set tRs = cnn.Execute(sBuscar)
        INV = tRs.Fields("ID_INVENTARIO")
        sBuscar = "INSERT INTO INVENTARIO_DETALLE (ID_INVENTARIO, ID_PRODUCTO, CANTIDAD) Values (" & INV & ", '" & frmInv.pro & "', " & Val(Me.txtCantidad.Text) & ")"
        cnn.Execute (sBuscar)
    End If
    Me.cmdOk.Enabled = False
    Me.lblInv2.Caption = "PRODUCTO: " & frmInv.pro & " CANTIDAD: " & Val(Me.txtCantidad.Text)
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
    With lvwProductos
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "PRODUCTO", 2880, 2
        .ColumnHeaders.Add , , "CANTIDAD", 1440
    End With
    Me.Caption = "INVENTARIO DE " & frmInv.pro
    Llenar_Lista_Inventario_Producto frmInv.pro
    Me.lblInv.Caption = "HACER INVENTARIO DE " & frmInv.pro & " EN LA SUCURSAL " & frmSucInv.NSUC
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Inventario_Producto(Clave As String)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT  SUCURSAL, CANTIDAD From EXISTENCIAS WHERE ID_PRODUCTO= '" & Clave & "' ORDER BY SUCURSAL"
    While Not tRs.EOF
        Set ItMx = Me.lvwInvSucPro.ListItems.Add(, , tRs.Fields("Sucursal"))
        If Not IsNull(tRs.Fields("CANTIDAD")) Then ItMx.SubItems(1) = tRs.Fields("CANTIDAD")
        tRs.MoveNext
    Wend
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub txtCantidad_GotFocus()
On Error GoTo ManejaError
    Me.txtCantidad.BackColor = vbWhite
    Me.txtCantidad.Text = ""
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.cmdOk.Value = True
    Else
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
