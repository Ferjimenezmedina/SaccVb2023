VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmExisReales 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bajar Listados de Existencias Reales"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8160
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   5280
      Picture         =   "FrmExisReales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8040
      TabIndex        =   6
      Top             =   2640
      Width           =   975
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EXCEL"
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
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmExisReales.frx":29D2
         MousePointer    =   99  'Custom
         Picture         =   "FrmExisReales.frx":2CDC
         Top             =   240
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmExisReales.frx":481E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
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
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8040
      TabIndex        =   2
      Top             =   3840
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmExisReales.frx":483A
         MousePointer    =   99  'Custom
         Picture         =   "FrmExisReales.frx":4B44
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sucursal"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Almacen 3"
         Height          =   195
         Left            =   6840
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Almacen 2"
         Height          =   195
         Left            =   6840
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Almacen 1"
         Height          =   195
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "FrmExisReales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Command3_Click()
    If Option3.value Then
        FunAlm3
    End If
    If Option2.value Then
        FunAlm2
    End If
    If Option1.value Then
        FunAlm1
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
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
        .ColumnHeaders.Add , , "Clave del Producto", 4000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Real", 0
        .ColumnHeaders.Add , , "Sucursal", 2500
    End With
    'Combo1.Clear
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image10_Click()
    If ListView1.ListItems.COUNT > 0 Then
        Dim StrCopi As String
        Dim Con As Integer
        Dim Con2 As Integer
        Dim NumColum As Integer
        Dim Ruta As String
        Me.CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "Guardar como"
        CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
        Me.CommonDialog1.ShowSave
        Ruta = Me.CommonDialog1.FileName
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.COUNT
            For Con = 1 To ListView1.ListItems.COUNT
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            'archivo TXT
            Dim foo As Integer
        
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub FunAlm1()
    If Combo1.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim sBuscar As String
        Dim tLi As ListItem
        ListView1.ListItems.Clear
        sBuscar = "SELECT ID_PRODUCTO, CANTIDAD FROM VSEXISALMA1 WHERE SUCURSAL = '" & Combo1.Text & "' ORDER BY ID_PRODUCTO"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = tRs.Fields("CANTIDAD")
                tLi.SubItems(2) = ""
                tLi.SubItems(3) = Combo1.Text
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub FunAlm2()
    If Combo1.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim sBuscar As String
        Dim tLi As ListItem
        ListView1.ListItems.Clear
        sBuscar = "SELECT ID_PRODUCTO, CANTIDAD FROM VSEXISALMA2 WHERE SUCURSAL = '" & Combo1.Text & "' ORDER BY ID_PRODUCTO"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = tRs.Fields("CANTIDAD")
                tLi.SubItems(2) = ""
                tLi.SubItems(3) = Combo1.Text
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub FunAlm3()
    If Combo1.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim sBuscar As String
        Dim tLi As ListItem
        ListView1.ListItems.Clear
        If Combo1.Text = "BODEGA" Then
            sBuscar = "SELECT * FROM VsExisRealesBodega WHERE CANTIDAD >0 AND APARTADO >0 ORDER BY ID_PRODUCTO"
        Else
            sBuscar = "SELECT * FROM EXISTENCIAS WHERE SUCURSAL = '" & Combo1.Text & "' AND CANTIDAD > 0 ORDER BY ID_PRODUCTO"
        End If
        Set tRs = cnn.Execute(sBuscar)
        If Combo1.Text = "BODEGA" Then
            If Not (tRs.EOF And tRs.EOF) Then
                Do While Not tRs.EOF
                    Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                    If Not IsNull(tRs.Fields("APARTADO")) Then
                        tLi.SubItems(1) = tRs.Fields("APARTADO") + tRs.Fields("CANTIDAD")
                    Else
                       tLi.SubItems(1) = tRs.Fields("CANTIDAD")
                    End If
                    tLi.SubItems(2) = ""
                    tLi.SubItems(2) = Combo1.Text
                Loop
            End If
        Else
            If Not (tRs.EOF And tRs.EOF) Then
                Do While Not tRs.EOF
                    Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                    tLi.SubItems(1) = tRs.Fields("CANTIDAD")
                    tLi.SubItems(2) = ""
                    tLi.SubItems(2) = Combo1.Text
                    tRs.MoveNext
                Loop
            End If
        End If
    Else
        MsgBox "DEBE SELECCIONAR UNA SUCURSAL PARA EL LISTADO", vbInformation, "SACC"
    End If
End Sub
