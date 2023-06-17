VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRecordatorios 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recordatorios"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   1
      Top             =   5520
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRecordatorios.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRecordatorios.frx":030A
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11668
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Recordatorios"
      TabPicture(0)   =   "FrmRecordatorios.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Nuevo"
      TabPicture(1)   =   "FrmRecordatorios.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Combo1"
      Tab(1).Control(3)=   "Option1"
      Tab(1).Control(4)=   "Option2"
      Tab(1).Control(5)=   "Option3"
      Tab(1).Control(6)=   "Text2"
      Tab(1).Control(7)=   "Command1"
      Tab(1).Control(8)=   "DTPicker1"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Este més"
      TabPicture(2)   =   "FrmRecordatorios.frx":2424
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView2"
      Tab(2).Control(1)=   "Text3"
      Tab(2).ControlCount=   2
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   2775
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   3600
         Width           =   8535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -67680
         TabIndex        =   12
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50266113
         CurrentDate     =   45072
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
         Height          =   375
         Left            =   -67560
         Picture         =   "FrmRecordatorios.frx":2440
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   2775
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1440
         Width           =   8535
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Anual"
         Height          =   255
         Left            =   -70440
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Mensual"
         Height          =   255
         Left            =   -70440
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Único Aviso"
         Height          =   255
         Left            =   -70440
         TabIndex        =   7
         Top             =   600
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -73560
         TabIndex        =   5
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   3600
         Width           =   8535
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5106
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   15
         Top             =   600
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5106
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Iniciar el día"
         Height          =   255
         Left            =   -68760
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento"
         Height          =   255
         Left            =   -74760
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmRecordatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    If Text2.Text <> "" And Combo1.Text <> "" Then
        Dim sBuscar As String
        Dim sTipo As String
        If Option1.Value Then
            sTipo = "U"
        Else
            If Option2.Value Then
                sTipo = "M"
            Else
                sTipo = "A"
            End If
        End If
        sBuscar = "INSERT INTO RECORDATORIOS (ID_USUARIO, DEPARTAMENTO, MENSAJE, TIPO, FECHA_RECORDAR, FECHA_ALTA) VALUES ('" & VarMen.Text1(0).Text & "', '" & VarMen.Text1(75).Text & "', '" & Text1.Text & "', '" & sTipo & "', '" & DTPicker1.Value & "', GETDATE())"
        cnn.Execute (sBuscar)
        Text2.Text = ""
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim tLi As ListItem
    DTPicker1.Value = Date
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
        .FullRowSelect = True
        .ColumnHeaders.Add , , "RECORDATORIO", 9200
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "RECORDATORIO", 9200
    End With
    sBuscar = "SELECT DEPARTAMENTO FROM DEPARTAMENTOS WHERE ESTATUS = 'A' AND TIPO = 'T' ORDER BY DEPARTAMENTO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("DEPARTAMENTO")) Then Combo1.AddItem tRs.Fields("DEPARTAMENTO")
            tRs.MoveNext
        Loop
    End If
    BuscaPendientes
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub BuscaPendientes()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT MENSAJE FROM RECORDATORIOS WHERE (TIPO = 'A') AND (DAY(FECHA_RECORDAR) = DAY(GETDATE())) AND (MONTH(FECHA_RECORDAR) = MONTH(GETDATE())) AND (DEPARTAMENTO = '" & VarMen.Text1(75).Text & "')"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("MENSAJE"))
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT MENSAJE FROM RECORDATORIOS WHERE (TIPO = 'M') AND (DAY(FECHA_RECORDAR) = DAY(GETDATE())) AND (DEPARTAMENTO = '" & VarMen.Text1(75).Text & "')"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("MENSAJE"))
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT MENSAJE FROM RECORDATORIOS WHERE (TIPO = 'U') AND (FECHA_RECORDAR = '" & Date & "') AND (DEPARTAMENTO = '" & VarMen.Text1(75).Text & "')"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("MENSAJE"))
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub BuscaPendientesMes()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView2.ListItems.Clear
    sBuscar = "SELECT MENSAJE FROM RECORDATORIOS WHERE (MONTH(FECHA_RECORDAR) = MONTH(GETDATE())) AND (DEPARTAMENTO = '" & VarMen.Text1(75).Text & "')"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("MENSAJE"))
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item
End Sub
