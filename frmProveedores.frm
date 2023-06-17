VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmProveedores 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8880
      TabIndex        =   5
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
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmProveedores.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmProveedores.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmProveedores.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvwProveedores"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   480
         Width           =   5295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Busca"
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
         Left            =   7320
         Picture         =   "frmProveedores.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwProveedores 
         Height          =   4335
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7646
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
         Caption         =   "Buscar por nombre :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Dim Cont As Integer
Dim NoRe As Integer
Dim NomProv As String
Dim DirProv As String
Dim ColProv As String
Dim CiuProv As String
Dim CPProv As String
Dim RFCProv As String
Dim Te1Prov As String
Dim Te2Prov As String
Dim Te3Prov As String
Dim NotProv As String
Dim EstProv As String
Dim PaiProv As String
Private Sub Command1_Click()
On Error GoTo ManejaError
    Llenar_Lista_Proveedores
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
    With Me.lvwProveedores
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 4500
        .ColumnHeaders.Add , , "DIRECCION", 5500
        .ColumnHeaders.Add , , "COLONIA", 1440
        .ColumnHeaders.Add , , "CIUDAD", 1440
        .ColumnHeaders.Add , , "CP", 1440
        .ColumnHeaders.Add , , "RFC", 1440
        .ColumnHeaders.Add , , "TELEFONO 1", 1440
        .ColumnHeaders.Add , , "TELEFONO 2", 1440
        .ColumnHeaders.Add , , "FAX", 1440
        .ColumnHeaders.Add , , "NOTAS", 1440
        .ColumnHeaders.Add , , "ESTADO", 1440
        .ColumnHeaders.Add , , "PAIS", 1440
        .ColumnHeaders.Add , , "BANCO", 1440
        .ColumnHeaders.Add , , "DIRECCION BANCO", 1440
        .ColumnHeaders.Add , , "CIUDAD BANCO", 1440
        .ColumnHeaders.Add , , "ROUTING", 1440
        .ColumnHeaders.Add , , "CUENTA BANCO", 1440
        .ColumnHeaders.Add , , "CLAVE SWIFT", 1440
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Proveedores()
On Error GoTo ManejaError
    sqlQuery = "SELECT * FROM PROVEEDOR WHERE NOMBRE LIKE '%" & Text1.Text & "%' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        Me.lvwProveedores.ListItems.Clear
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvwProveedores.ListItems.Add(, , .Fields("ID_PROVEEDOR"))
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = Trim(.Fields("NOMBRE"))
                If Not IsNull(.Fields("DIRECCION")) Then tLi.SubItems(2) = Trim(.Fields("DIRECCION"))
                If Not IsNull(.Fields("COLONIA")) Then tLi.SubItems(3) = Trim(.Fields("COLONIA"))
                If Not IsNull(.Fields("CIUDAD")) Then tLi.SubItems(4) = Trim(.Fields("CIUDAD"))
                If Not IsNull(.Fields("CP")) Then tLi.SubItems(5) = Trim(.Fields("CP"))
                If Not IsNull(.Fields("RFC")) Then tLi.SubItems(6) = Trim(.Fields("RFC"))
                If Not IsNull(.Fields("TELEFONO1")) Then tLi.SubItems(7) = Trim(.Fields("TELEFONO1"))
                If Not IsNull(.Fields("TELEFONO2")) Then tLi.SubItems(8) = Trim(.Fields("TELEFONO2"))
                If Not IsNull(.Fields("TELEFONO3")) Then tLi.SubItems(9) = Trim(.Fields("TELEFONO3"))
                If Not IsNull(.Fields("NOTAS")) Then tLi.SubItems(10) = Trim(.Fields("NOTAS"))
                If Not IsNull(.Fields("ESTADO")) Then tLi.SubItems(11) = Trim(.Fields("ESTADO"))
                If Not IsNull(.Fields("PAIS")) Then tLi.SubItems(12) = Trim(.Fields("PAIS"))
                If Not IsNull(.Fields("TRANS_BANCO")) Then tLi.SubItems(13) = Trim(.Fields("TRANS_BANCO"))
                If Not IsNull(.Fields("TRANS_DIRECCION")) Then tLi.SubItems(14) = Trim(.Fields("TRANS_DIRECCION"))
                If Not IsNull(.Fields("TRANS_CIUDAD")) Then tLi.SubItems(15) = Trim(.Fields("TRANS_CIUDAD"))
                If Not IsNull(.Fields("TRANS_ROUTING")) Then tLi.SubItems(16) = Trim(.Fields("TRANS_ROUTING"))
                If Not IsNull(.Fields("TRANS_CUENTA")) Then tLi.SubItems(17) = Trim(.Fields("TRANS_CUENTA"))
                If Not IsNull(.Fields("TRANS_CLAVE_SWIFT")) Then tLi.SubItems(18) = Trim(.Fields("TRANS_CLAVE_SWIFT"))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwProveedores_DblClick()
On Error GoTo ManejaError
    If MsgBox("          ¿DESEA IMPRIMIR LOS DATOS DEL PROVEEDOR " & NomProv & "?          ", vbInformation + vbYesNo + vbDefaultButton1, "MESAJE DEL SISTEMA") = vbYes Then
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "     FECHA : " & Format(Date, "dd/mm/yyyy")
        Printer.CurrentY = 2600
        Printer.CurrentX = 600
        Printer.Print "Nombre del Proveedor    : " & NomProv
        Printer.Print "Direccion               : " & DirProv
        Printer.Print "Colonia                 : " & ColProv
        Printer.Print "Codigo Postal           : " & CPProv
        Printer.Print "Ciudad                  : " & CiuProv
        Printer.Print "Estado                  : " & EstProv
        Printer.Print "Pais                    : " & PaiProv
        Printer.Print "RFC                     : " & RFCProv
        Printer.Print "Telefono Principal      : " & Te1Prov
        Printer.Print "Telefono Secundario     : " & Te2Prov
        Printer.Print "Fax                     : " & Te3Prov
        Printer.Print "Notas                   : " & NotProv
        Printer.Print "_______________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________"
        Printer.EndDoc
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwProveedores_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    NomProv = Item.SubItems(1)
    DirProv = Item.SubItems(2)
    ColProv = Item.SubItems(3)
    CiuProv = Item.SubItems(4)
    CPProv = Item.SubItems(5)
    RFCProv = Item.SubItems(6)
    Te1Prov = Item.SubItems(7)
    Te2Prov = Item.SubItems(8)
    Te3Prov = Item.SubItems(9)
    NotProv = Item.SubItems(10)
    EstProv = Item.SubItems(11)
    PaiProv = Item.SubItems(12)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
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
    If KeyAscii = 13 And Text1.Text <> "" Then
        Llenar_Lista_Proveedores
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
