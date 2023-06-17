VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmProvAlmace1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Proveedor Almacen1"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8040
      TabIndex        =   14
      Top             =   2040
      Width           =   975
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   19
         Top             =   1320
         Width           =   975
         Begin VB.Image Image15 
            Height          =   720
            Left            =   120
            MouseIcon       =   "FrmProvAlmace1.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "FrmProvAlmace1.frx":030A
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   17
         Top             =   1320
         Width           =   975
         Begin VB.Label Label17 
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
            TabIndex        =   18
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "FrmProvAlmace1.frx":1CCC
            MousePointer    =   99  'Custom
            Picture         =   "FrmProvAlmace1.frx":1FD6
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   15
         Top             =   0
         Width           =   975
         Begin VB.Label Label18 
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
            TabIndex        =   16
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image17 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmProvAlmace1.frx":3800
            MousePointer    =   99  'Custom
            Picture         =   "FrmProvAlmace1.frx":3B0A
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Eliminar"
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
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   750
         Left            =   120
         MouseIcon       =   "FrmProvAlmace1.frx":55BC
         MousePointer    =   99  'Custom
         Picture         =   "FrmProvAlmace1.frx":58C6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8040
      TabIndex        =   12
      Top             =   3240
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmProvAlmace1.frx":75F0
         MousePointer    =   99  'Custom
         Picture         =   "FrmProvAlmace1.frx":78FA
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
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8040
      TabIndex        =   6
      Top             =   840
      Width           =   975
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmProvAlmace1.frx":99DC
         MousePointer    =   99  'Custom
         Picture         =   "FrmProvAlmace1.frx":9CE6
         Top             =   240
         Width           =   675
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Proveedor"
      TabPicture(0)   =   "FrmProvAlmace1.frx":B6A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   4
         Top             =   3840
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   3
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   2
         Top             =   3120
         Width           =   6135
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3625
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
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label9 
         Caption         =   "E - Mail"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "* Telefono"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "* Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmProvAlmace1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim IdProv As String
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Proveedor", 0
        .ColumnHeaders.Add , , "Nombre", 5500
        .ColumnHeaders.Add , , "Telefono", 1500
        .ColumnHeaders.Add , , "Correo Electronico", 1500
    End With
End Sub
Private Sub Image2_Click()
    If IdProv <> "" Then
        If MsgBox("          Esta seguro que desea eliminar el registro?          ", vbYesNo) = vbYes Then
            Dim sBuscar As String
            Dim tRs As ADODB.Recordset
            sBuscar = "DELETE FROM PROVEEDOR_ALMACEN1 WHERE ID_PROVEEDOR = " & IdProv
            Set tRs = cnn.Execute(sBuscar)
            MsgBox "REGISTRO ELIMINADO", vbInformation, "SACC"
        End If
    Else
        MsgBox "ES NECESARIO QUE SELECCIONE UN PROVEEDOR!", vbInformation, "SACC"
    End If
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    ListView1.ListItems.Clear
End Sub
Private Sub Image8_Click()
    Dim sBuscar As String
    If IdProv <> "" Then
        Dim tRs As ADODB.Recordset
        sBuscar = "UPDATE PROVEEDOR_ALMACEN1 SET NOMBRE = '" & Text2.Text & "', TELEFONO = '" & Text3.Text & "', MAIL = '" & Text4.Text & "' WHERE ID_PROVEEDOR = " & IdProv
        Set tRs = cnn.Execute(sBuscar)
        IdProv = ""
        MsgBox "LA INFORMACIÓN DEL PROVEEDOR HA SIDO MODIFICADA", vbInformation, "SACC"
    Else
        sBuscar = "INSERT INTO PROVEEDOR_ALMACEN1 (NOMBRE, TELEFONO, MAIL) VALUES ('" & Text2.Text & "', '" & Text3.Text & "', '" & Text4.Text & "' );"
        cnn.Execute (sBuscar)
        MsgBox "LA INFORMACIÓN DEL PROVEEDOR HA SIGO GUARDADA!", vbInformation, "SACC"
    End If
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    ListView1.ListItems.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdProv = Item
    Text2.Text = Item.SubItems(1)
    Text3.Text = Item.SubItems(2)
    Text4.Text = Item.SubItems(3)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text1.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        ListView1.ListItems.Clear
        sBuscar = "SELECT * FROM PROVEEDOR_ALMACEN1 WHERE NOMBRE LIKE '%" & Text1.Text & "%' ORDER BY NOMBRE"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                    Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR"))
                    tLi.SubItems(1) = tRs.Fields("NOMBRE")
                    tLi.SubItems(2) = tRs.Fields("TELEFONO")
                    tLi.SubItems(3) = tRs.Fields("MAIL")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890()- "
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ@_-"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
