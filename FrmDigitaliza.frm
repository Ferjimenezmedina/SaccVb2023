VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmDigitaliza 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guardar Documentos en la Base de Datos"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7680
      TabIndex        =   14
      Top             =   120
      Width           =   975
      Begin VB.Image Image2 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmDigitaliza.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmDigitaliza.frx":030A
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label4 
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
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7680
      TabIndex        =   6
      Top             =   2880
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
         MouseIcon       =   "FrmDigitaliza.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "FrmDigitaliza.frx":1FD6
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7680
      TabIndex        =   1
      Top             =   4080
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmDigitaliza.frx":3998
         MousePointer    =   99  'Custom
         Picture         =   "FrmDigitaliza.frx":3CA2
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Datos del Cliente"
      TabPicture(0)   =   "FrmDigitaliza.frx":5D84
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CommonDialog1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Imagen"
      TabPicture(1)   =   "FrmDigitaliza.frx":5DA0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture1 
         Height          =   4455
         Left            =   -74880
         ScaleHeight     =   4395
         ScaleWidth      =   7155
         TabIndex        =   16
         Top             =   600
         Width           =   7215
      End
      Begin VB.CommandButton Command3 
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
         Left            =   6120
         Picture         =   "FrmDigitaliza.frx":5DBC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   4260
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Archivo"
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
         Left            =   6120
         Picture         =   "FrmDigitaliza.frx":878E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4260
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3900
         Width           =   5055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Digitalizar"
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
         Left            =   4800
         Picture         =   "FrmDigitaliza.frx":B160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4320
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2535
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4471
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
         TabIndex        =   4
         Top             =   420
         Width           =   4935
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3420
         Width           =   7215
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion del Documento :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3900
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   420
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmDigitaliza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim ClvCliente As String
Dim ImgPath As String
Private Sub Command2_Click()
    Dim archivo As String
    On Error GoTo errorhandler ' introducimos un control de error por código
    With CommonDialog1
        'determina el tipo de archivo a abrir
        'common dialog control
        .Filter = "Archivos de Imagen(*.jpg,*.gif,*.bmp)|*.jpg|*.gif|*.bmp|"
        .ShowOpen ' muestra la ventana típica de abrir archivos
        If Len(.FileName) = 0 Then ' si no seleccionamos ningún archivo sale del 'procedimiento.
            Exit Sub
        End If
        archivo = .FileName ' del commondialog
    End With
    ImgPath = CommonDialog1.FileName
    Picture1.Picture = LoadPicture(CommonDialog1.FileName) 'carga la imagen
errorhandler:
    Exit Sub
End Sub
Private Sub Command3_Click()
    Buscar
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .CheckBoxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave", 0
        .ColumnHeaders.Add , , "Nombre", 7150
    End With
End Sub
Private Sub Image2_Click()
    'On Error GoTo ControlError
    Dim chunk() As Byte
    Dim fd As Integer
    Dim flen As Long
    Dim pat As String
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim OUTPUT_FILE_PATH As String
    pat = App.Path
    OUTPUT_FILE_PATH = pat & "\imagen.jpg"
    sBuscar = "SELECT ID_CLIENTE, DOCUMENTO, IMAGEN FROM DOCUMENTOS_CLIENTE WHERE ID_CLIENTE = " & ClvCliente
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        flen = tRs.Fields("IMAGEN").ActualSize
        fd = FreeFile
        Open OUTPUT_FILE_PATH For Binary Access Write As fd
        ReDim chunk(1 To flen)
        Put fd, , chunk()
        chunk() = tRs.Fields("IMAGEN") '.GetChunk(flen)
    End If
    Exit Sub
ControlError:
End Sub
Private Sub Image8_Click()
    Dim chunk() As Byte
    Dim fd As Integer
    Dim flen As Long
    Dim sBuscar As String
    Dim INPUT_FILE_PATH As String
    'Dim cmd As ADODB.Command
    'Dim co_empresa As ADODB.Parameter
    'Dim co_filial As ADODB.Parameter
    'Dim co_activo As ADODB.Parameter
    'Dim nb_imagen As ADODB.Parameter
    'Dim id_usuario As ADODB.Parameter
    'Dim imagen As ADODB.Parameter
    INPUT_FILE_PATH = ImgPath
    'Set C = New oculta
    'C.InhabilitaSmart Me
    'btnAdd.Enabled = False
    'frmFichaActivos.MousePointer = vbHourglass
    'btnCancelar.Enabled = False
    'Set cmd = New ADODB.Command
    'cmd.ActiveConnection = BaseRemota
    'Select Case paso
    '    Case "Fotografias"
    '        StatusBar1.Panels(1).Text = "Guardando Fotografía... espere un momento por favor"
            'cmd.CommandText = "insert into samafact (co_empresa,co_filial,co_activo,  nb_imagen,id_usuario,imagen)values ( ?,?,?,?,?,?)"
    '    Case "Factura"
    '        StatusBar1.Panels(1).Text = "Guardando Factura... espere un momento por favor"
            'cmd.CommandText = "insert into samafact (co_empresa,co_filial,co_activo,  nb_imagen,id_usuario,imagen) values ( ?,?,?,?,?,?)"
    '    Case "Documentos"
    '        StatusBar1.Panels(1).Text = "Guardando Documentos... espere un momento por favor"
            'cmd.CommandText = "insert into samaimdo (co_empresa,co_filial,co_activo,  nb_imagen,id_usuario,imagen) values ( ?,?,?,?,?,?)"
    '    Case "Planos"
    '        StatusBar1.Panels(1).Text = "Guardando Planos... espere un momento por favor"
            'cmd.CommandText = "insert into samaplac (co_empresa,co_filial,co_activo,  nb_imagen,id_usuario,imagen) values ( ?,?,?,?,?,?)"
    'End Select
    'cmd.CommandType = adCmdText
    'Set co_empresa = cmd.CreateParameter("co_empresa", adInteger, adParamInput)
    'co_empresa.Value = UserCia
    'cmd.Parameters.Append co_empresa
    'Set co_filial = cmd.CreateParameter("co_filial", adInteger, adParamInput)
    'co_filial.Value = UserSede
    'cmd.Parameters.Append co_filial
    'Set co_activo = cmd.CreateParameter("co_activo", adLongVarChar, adParamInput, 60)
    'co_activo.Value = Trim(txt_nu_activo_fijo.Text)
    'cmd.Parameters.Append co_activo
    'Set nb_imagen = cmd.CreateParameter("nb_imagen", adLongVarChar, adParamInput, 60)
    'nb_imagen.Value = UCase(Trim(txt_nb_imagen.Text))
    'cmd.Parameters.Append nb_imagen
    'Set id_usuario = cmd.CreateParameter("id_usuario", adLongVarChar, adParamInput, 25)
    'id_usuario.Value = UserID
    'cmd.Parameters.Append id_usuario
    
    fd = FreeFile
    Open INPUT_FILE_PATH For Binary Access Read As fd
    flen = LOF(fd)
    If flen = 0 Then
        Close
        MsgBox "error"
        End
    Else
    End If
    'Set imagen = cmd.CreateParameter("imagen", adLongVarBinary, adParamInput, flen)
    ReDim chunk(1 To flen)
    Get fd, , chunk()
    'imagen.AppendChunk chunk()
    sBuscar = "INSERT INTO DOCUMENTOS_CLIENTE (ID_CLIENTE, DOCUMENTO, IMAGEN) VALUES (" & ClvCliente & ", '" & Text2.Text & "', " & flen & ")"
    MsgBox sBuscar
    cnn.Execute (sBuscar)
    'Call inserta_imagen
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.Label3.Caption = Item.SubItems(1)
    ClvCliente = Item
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Buscar
    End If
End Sub
Private Sub Buscar()
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub

