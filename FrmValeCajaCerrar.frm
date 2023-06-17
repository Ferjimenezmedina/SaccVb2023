VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmValeCajaCerrar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cerrar Valde de Caja"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccionado"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   9135
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No. Vale de Caja :"
         Height          =   255
         Left            =   6360
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame21 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9360
      TabIndex        =   6
      Top             =   3840
      Width           =   975
      Begin VB.Label Label21 
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Image imgLeer 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmValeCajaCerrar.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmValeCajaCerrar.frx":030A
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9360
      TabIndex        =   4
      Top             =   5040
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmValeCajaCerrar.frx":1DBC
         MousePointer    =   99  'Custom
         Picture         =   "FrmValeCajaCerrar.frx":20C6
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
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Aplicadas"
      TabPicture(0)   =   "FrmValeCajaCerrar.frx":41A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Pendientes"
      TabPicture(1)   =   "FrmValeCajaCerrar.frx":41C4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView2"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView ListView2 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   8
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7646
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4335
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7646
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "FrmValeCajaCerrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
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
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. VALE", 800
        .ColumnHeaders.Add , , "CLIENTE", 4200
        .ColumnHeaders.Add , , "TOTAL", 2000
        .ColumnHeaders.Add , , "FECHA", 2000
    End With
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. VALE", 800
        .ColumnHeaders.Add , , "CLIENTE", 4200
        .ColumnHeaders.Add , , "TOTAL", 2000
        .ColumnHeaders.Add , , "FECHA", 2000
    End With
    Buscar
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Buscar()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT ID_VALE, IMPORTE, FECHA, NOMBRE  FROM VsValeAplica WHERE APLICADO = 'S' ORDER BY ID_VALE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VALE"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("IMPORTE")) Then tLi.SubItems(2) = tRs.Fields("IMPORTE")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(3) = tRs.Fields("FECHA")
            tRs.MoveNext
        Loop
    End If
    ListView2.ListItems.Clear
    sBuscar = "SELECT ID_VALE, IMPORTE, FECHA, NOMBRE  FROM VsValeAplica WHERE APLICADO = 'N' ORDER BY ID_VALE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_VALE"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("IMPORTE")) Then tLi.SubItems(2) = tRs.Fields("IMPORTE")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(3) = tRs.Fields("FECHA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub imgLeer_Click()
    If Text2.Text <> "" Then
        Dim sBuscar As String
        sBuscar = "UPDATE VALE_CAJA SET APLICADO = 'F' WHERE ID_VALE = " & Text2.Text
        If MsgBox("ESTA SEGURO QUE DESEA FINALIZAR EL VALE?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
            cnn.Execute (sBuscar)
            Buscar
        End If
    Else
        MsgBox "NO HA SELECCIONADO UN VALE DE CAJA PARA CERRAR", vbInformation, "SACC"
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item.SubItems(1)
    Text2.Text = Item
End Sub
