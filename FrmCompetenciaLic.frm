VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCompetenciaLic 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Competidores en Licitaciones"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6360
      TabIndex        =   19
      Top             =   2400
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
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmCompetenciaLic.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCompetenciaLic.frx":030A
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6360
      TabIndex        =   11
      Top             =   1200
      Width           =   975
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   16
         Top             =   1320
         Width           =   975
         Begin VB.Image Image15 
            Height          =   720
            Left            =   120
            MouseIcon       =   "FrmCompetenciaLic.frx":1CCC
            MousePointer    =   99  'Custom
            Picture         =   "FrmCompetenciaLic.frx":1FD6
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
            TabIndex        =   17
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   14
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
            TabIndex        =   15
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "FrmCompetenciaLic.frx":3998
            MousePointer    =   99  'Custom
            Picture         =   "FrmCompetenciaLic.frx":3CA2
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   12
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
            TabIndex        =   13
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image17 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmCompetenciaLic.frx":54CC
            MousePointer    =   99  'Custom
            Picture         =   "FrmCompetenciaLic.frx":57D6
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
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image18 
         Height          =   750
         Left            =   120
         MouseIcon       =   "FrmCompetenciaLic.frx":7288
         MousePointer    =   99  'Custom
         Picture         =   "FrmCompetenciaLic.frx":7592
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6360
      TabIndex        =   9
      Top             =   3600
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCompetenciaLic.frx":92BC
         MousePointer    =   99  'Custom
         Picture         =   "FrmCompetenciaLic.frx":95C6
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
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Buscar"
      TabPicture(0)   =   "FrmCompetenciaLic.frx":B6A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Información"
      TabPicture(1)   =   "FrmCompetenciaLic.frx":B6C4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Text5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label3"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Height          =   1695
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2640
         Width           =   5775
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   -73920
         MaxLength       =   25
         TabIndex        =   5
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -73920
         MaxLength       =   80
         TabIndex        =   4
         Top             =   1440
         Width           =   4935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -73920
         MaxLength       =   50
         TabIndex        =   3
         Top             =   960
         Width           =   4935
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5953
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
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
         Left            =   4680
         Picture         =   "FrmCompetenciaLic.frx":B6E0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "e-mail :"
         Height          =   255
         Left            =   -72120
         TabIndex        =   26
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Nota :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   25
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Telefono :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Direccion :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "* Nombre :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmCompetenciaLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT * FROM COMPETIDOR_LICITACION WHERE NOMBRE = '" & Text1.Text & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Set tLi = ListView1.ListItems.Add(, , tRs.Fields("NOMBRE"))
        If Not IsNull(tRs.Fields("DIRECCION")) Then tLi.SubItems(1) = tRs.Fields("DIRECCION")
        If Not IsNull(tRs.Fields("TELEFONO")) Then tLi.SubItems(2) = tRs.Fields("TELEFONO")
        If Not IsNull(tRs.Fields("MAIL")) Then tLi.SubItems(3) = tRs.Fields("MAIL")
        If Not IsNull(tRs.Fields("NOTAS")) Then tLi.SubItems(4) = tRs.Fields("NOTAS")
        tRs.MoveNext
    End If
End Sub
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
        .CheckBoxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "NOMBRE", 250
        .ColumnHeaders.Add , , "DIRECCION", 2200
        .ColumnHeaders.Add , , "TELEFONO", 1000
        .ColumnHeaders.Add , , "E-MAIL", 1000
        .ColumnHeaders.Add , , "NOTAS", 2500
    End With
End Sub
Private Sub Image18_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If Text2.Text <> "" Then
        sBuscar = "SELECT NOMBRE FROM COMPETIDOR_LICITACION WHERE NOMBRE = '" & Text2.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            sBuscar = "DELETE FROM COMPETIDOR_LICITACION NOMBRE = '" & Text2.Text & "'"
            cnn.Execute (sBuscar)
        Else
            MsgBox "EL REGISTRO NO SE ENCONTRO!", vbInformation, "SACC"
        End If
    Else
        MsgBox "ES NECESARIO SELECCIONAR UN REGISTRO A ELIMINAR", vbInformation, "SACC"
    End If
End Sub
Private Sub Image8_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If Text2.Text <> "" Then
        sBuscar = "SELECT NOMBRE FROM COMPETIDOR_LICITACION WHERE NOMBRE = '" & Text2.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If (tRs.EOF And tRs.BOF) Then
            sBuscar = "INSERT INTO COMPETIDOR_LICITACION (NOMBRE, DIRECCION, TELEFONO, MAIL, NOTAS) VALUES('" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "')"
            cnn.Execute (sBuscar)
        Else
            If MsgBox("YA EXISTE UN REGISTRO CON ESE NOMBRE, ¿DESEA MODIFICARLO?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
                sBuscar = "UPDATE COMPETIDOR_LICITACION SET NOMBRE = '" & Text2.Text & "', DIRECCION = '" & Text3.Text & "', TELEFONO = '" & Text4.Text & "', MAIL = '" & Text5.Text & "', NOTAS = '" & Text6.Text & "' WHERE NOMBRE = '" & Text2.Text & "'"
                cnn.Execute (sBuscar)
            End If
        End If
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
    Else
        MsgBox "ES NECESARIO UN NOMBRE PARA EL REGISTRO", vbInformation, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text2.Text = Item
    Text3.Text = Item.SubItems(1)
    Text4.Text = Item.SubItems(2)
    Text5.Text = Item.SubItems(3)
    Text6.Text = Item.SubItems(4)
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890-()"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
