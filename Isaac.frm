VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAStec 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Asistencias Tecnicas Pendientes"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9960
      TabIndex        =   42
      Top             =   2280
      Width           =   975
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
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
         TabIndex        =   43
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "Isaac.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Isaac.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9960
      TabIndex        =   39
      Top             =   3480
      Width           =   975
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "Isaac.frx":1EDC
         MousePointer    =   99  'Custom
         Picture         =   "Isaac.frx":21E6
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label14 
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
         TabIndex        =   40
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame23 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9960
      TabIndex        =   35
      Top             =   4680
      Width           =   975
      Begin VB.Image Image21 
         Height          =   735
         Left            =   120
         MouseIcon       =   "Isaac.frx":3D28
         MousePointer    =   99  'Custom
         Picture         =   "Isaac.frx":4032
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cerrar"
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
         TabIndex        =   36
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9960
      TabIndex        =   33
      Top             =   5880
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
         TabIndex        =   34
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Isaac.frx":5B40
         MousePointer    =   99  'Custom
         Picture         =   "Isaac.frx":5E4A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Ver"
      TabPicture(0)   =   "Isaac.frx":7F2C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label15"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Check1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListView2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text13"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Option1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Option2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Check2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Check3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Información"
      TabPicture(1)   =   "Isaac.frx":7F48
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "Label7"
      Tab(1).Control(7)=   "Label12"
      Tab(1).Control(8)=   "Text2"
      Tab(1).Control(9)=   "Text3"
      Tab(1).Control(10)=   "Text4"
      Tab(1).Control(11)=   "Text5"
      Tab(1).Control(12)=   "Text6"
      Tab(1).Control(13)=   "Text7"
      Tab(1).Control(14)=   "Text8"
      Tab(1).Control(15)=   "Frame1"
      Tab(1).Control(16)=   "Text12"
      Tab(1).ControlCount=   17
      Begin VB.CheckBox Check3 
         Caption         =   "Solo Gatantias"
         Height          =   255
         Left            =   8160
         TabIndex        =   45
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Solo Asistencias a Domicilio"
         Height          =   255
         Left            =   5760
         TabIndex        =   44
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Comentario"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Cliente"
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   1080
         Value           =   -1  'True
         Width           =   1095
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
         Left            =   8520
         Picture         =   "Isaac.frx":7F64
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   600
         Width           =   6855
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   3960
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H8000000F&
         Height          =   1215
         Left            =   -72480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   5040
         Width           =   6135
      End
      Begin VB.Frame Frame1 
         Caption         =   "Información del Articulo"
         Height          =   2175
         Left            =   -74760
         TabIndex        =   25
         Top             =   2760
         Width           =   9135
         Begin VB.TextBox Text11 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   1440
            Width           =   6855
         End
         Begin VB.TextBox Text10 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   960
            Width           =   7335
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   480
            Width           =   7935
         End
         Begin VB.Label Label11 
            Caption         =   "Descripcion de Piezas :"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Tipo de Articulo :"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Modelo : "
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -68400
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -71040
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -67200
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -69840
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   7695
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
         Left            =   6720
         Picture         =   "Isaac.frx":A936
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6120
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         MaxLength       =   500
         TabIndex        =   6
         Top             =   6120
         Width           =   6495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Marcar como Terminada"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   6480
         Width           =   2175
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   3836
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
      Begin VB.Label Label15 
         Caption         =   "Buscar Cliente :"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Cerradas :"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Pendientes :"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Comentarios del Tecnico :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   17
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Marca :"
         Height          =   255
         Left            =   -69120
         TabIndex        =   16
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha de domicilio :"
         Height          =   255
         Left            =   -72600
         TabIndex        =   15
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "A Domicilio? :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de Captura :"
         Height          =   255
         Left            =   -68760
         TabIndex        =   13
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Es Garantia ?"
         Height          =   255
         Left            =   -71160
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Sucursal :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   11
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del cliente :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10440
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAStec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim NOTA As Integer
Dim DelAST As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
On Error GoTo ManejaError
    If NOTA = 0 Then
        MsgBox "No ha seleccionado una asistencia para asignar el mensaje"
    Else
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        If Check1.Value = 1 Then
            If Not (Text1.Text = "") Then
                sBuscar = "UPDATE ASISTENCIA_TECNICA SET COMENTARIOS_TECNICOS = '" & Text1.Text & "' WHERE ID_AS_TEC = " & NOTA
                cnn.Execute (sBuscar)
            End If
        Else
            If Text1.Text = "" Then
                MsgBox "No puede asignar un valor nulo por mensaje"
            Else
                sBuscar = "UPDATE ASISTENCIA_TECNICA SET COMENTARIOS_TECNICOS = '" & Text1.Text & "' WHERE ID_AS_TEC = " & NOTA
                cnn.Execute (sBuscar)
            End If
        End If
        Buscar
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command2_Click()
    Buscar
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
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. Asistencia", 800
        .ColumnHeaders.Add , , "Nombre del Cliente", 5500
        .ColumnHeaders.Add , , "Sucursal", 1500
        .ColumnHeaders.Add , , "Garantia?", 1500
        .ColumnHeaders.Add , , "Fecha de Captura", 1500
        .ColumnHeaders.Add , , "A Domicilio?", 1000
        .ColumnHeaders.Add , , "Fecha de domicilio", 1000
        .ColumnHeaders.Add , , "Marca", 1500
        .ColumnHeaders.Add , , "Modelo", 1500
        .ColumnHeaders.Add , , "Tipo de Articulo", 1500
        .ColumnHeaders.Add , , "Descripcion de Piezas", 5500
        .ColumnHeaders.Add , , "Comentarios del Tecnico", 6000
        If VarMen.Text1(7).Text = "S" Then
            .ColumnHeaders.Add , , "Telefono", 6000
        End If
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. Asistencia", 800
        .ColumnHeaders.Add , , "Nombre del Cliente", 5500
        .ColumnHeaders.Add , , "Sucursal", 1500
        .ColumnHeaders.Add , , "Garantia?", 1500
        .ColumnHeaders.Add , , "Fecha de Captura", 1500
        .ColumnHeaders.Add , , "A Domicilio?", 1000
        .ColumnHeaders.Add , , "Fecha de domicilio", 1000
        .ColumnHeaders.Add , , "Marca", 1500
        .ColumnHeaders.Add , , "Modelo", 1500
        .ColumnHeaders.Add , , "Tipo de Articulo", 1500
        .ColumnHeaders.Add , , "Descripcion de Piezas", 5500
        .ColumnHeaders.Add , , "Comentarios del Tecnico", 6000
        If VarMen.Text1(7).Text = "S" Then
            .ColumnHeaders.Add , , "Telefono", 6000
        End If
    End With
    Buscar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image1_Click()
    FunImpATec
End Sub
Private Sub Image10_Click()
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Dim foo As Integer
    Me.CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
    Me.CommonDialog1.ShowSave
    Ruta = Me.CommonDialog1.FileName
    If MsgBox("Desea reporte de A.T. pendientes?", vbYesNo, "SACC") = vbYes Then
        If ListView1.ListItems.Count > 0 Then
            If Ruta <> "" Then
                NumColum = ListView1.ColumnHeaders.Count
                For Con = 1 To ListView1.ColumnHeaders.Count
                    StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
                For Con = 1 To ListView1.ListItems.Count
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                Next
                'archivo TXT
                foo = FreeFile
                Open Ruta For Output As #foo
                    Print #foo, StrCopi
                Close #foo
            End If
            ShellExecute Me.hWnd, "open", Ruta, "", "", 4
        End If
    Else
        If ListView2.ListItems.Count > 0 Then
            If Ruta <> "" Then
                NumColum = ListView2.ColumnHeaders.Count
                For Con = 1 To ListView2.ColumnHeaders.Count
                    StrCopi = StrCopi & ListView2.ColumnHeaders(Con).Text & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
                For Con = 1 To ListView2.ListItems.Count
                    StrCopi = StrCopi & ListView2.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView2.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                Next
                'archivo TXT
                foo = FreeFile
                Open Ruta For Output As #foo
                    Print #foo, StrCopi
                Close #foo
            End If
            ShellExecute Me.hWnd, "open", Ruta, "", "", 4
        End If
    End If
End Sub
Private Sub Image21_Click()
On Error GoTo ManejaError
    FrmCierreAsTec.Show vbModal
    Buscar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    NOTA = ListView1.SelectedItem
    DelAST = Item
    Text2.Text = Item.SubItems(1)
    Text3.Text = Item.SubItems(2)
    Text4.Text = Item.SubItems(3)
    Text5.Text = Item.SubItems(4)
    Text6.Text = Item.SubItems(5)
    Text7.Text = Item.SubItems(6)
    Text8.Text = Item.SubItems(7)
    Text9.Text = Item.SubItems(8)
    Text10.Text = Item.SubItems(9)
    Text11.Text = Item.SubItems(10)
    Text12.Text = Item.SubItems(11)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    If Option1.Value Then
        sBuscar = "SELECT * FROM ASISTENCIA_TECNICA WHERE NOMBRE LIKE '%" & Text13.Text & "%' AND ATENDIDO = '0'"
    Else
        sBuscar = "SELECT * FROM ASISTENCIA_TECNICA WHERE (COMENTARIOS_COTIZACION LIKE '%" & Text13.Text & "%' OR COMENTARIOS_TECNICOS LIKE '%" & Text13.Text & "%' OR MARCA LIKE '%" & Text13.Text & "%') AND ATENDIDO = '0'"
    End If
    If Check2.Value = 1 Then
        sBuscar = sBuscar & " AND A_DOMICILIO = '1'"
    End If
    If Check3.Value = 1 Then
        sBuscar = sBuscar & " AND GARANTIA = '1'"
    End If
    sBuscar = sBuscar & " ORDER BY ID_AS_TEC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        'With tRs
            ListView1.ListItems.Clear
            Do While Not tRs.EOF
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_AS_TEC"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(2) = tRs.Fields("SUCURSAL")
                If tRs.Fields("GARANTIA") = 1 Then
                    tLi.SubItems(3) = "SI"
                Else
                    tLi.SubItems(3) = "NO"
                End If
                If Not IsNull(tRs.Fields("FECHA_CAPTURA")) Then tLi.SubItems(4) = tRs.Fields("FECHA_CAPTURA")
                If tRs.Fields("A_DOMICILIO") & "" = 1 Then
                    tLi.SubItems(5) = "SI"
                Else
                    tLi.SubItems(5) = "NO"
                End If
                If Not IsNull(tRs.Fields("FECHA_DEBE_ATENDER")) Then tLi.SubItems(6) = tRs.Fields("FECHA_DEBE_ATENDER")
                If Not IsNull(tRs.Fields("MARCA")) Then tLi.SubItems(7) = tRs.Fields("MARCA")
                If Not IsNull(tRs.Fields("MODELO")) Then tLi.SubItems(8) = tRs.Fields("MODELO")
                If Not IsNull(tRs.Fields("TIPO_ARTICULO")) Then tLi.SubItems(9) = tRs.Fields("TIPO_ARTICULO")
                If Not IsNull(tRs.Fields("Descripcion_PIEZAS")) Then tLi.SubItems(10) = tRs.Fields("Descripcion_PIEZAS")
                If Not IsNull(tRs.Fields("COMENTARIOS_TECNICOS")) Then tLi.SubItems(11) = tRs.Fields("COMENTARIOS_TECNICOS")
                If VarMen.Text1(7).Text = "S" Then
                    If Not IsNull(tRs.Fields("TELEFONO")) Then tLi.SubItems(12) = tRs.Fields("TELEFONO")
                End If
                tRs.MoveNext
            Loop
        'End With
    End If
    ListView2.ListItems.Clear
    If Option1.Value Then
        sBuscar = "SELECT * FROM ASISTENCIA_TECNICA WHERE NOMBRE LIKE '%" & Text13.Text & "%' AND ATENDIDO = '3'"
    Else
        sBuscar = "SELECT * FROM ASISTENCIA_TECNICA WHERE (COMENTARIOS_COTIZACION LIKE '%" & Text13.Text & "%' OR COMENTARIOS_TECNICOS LIKE '%" & Text13.Text & "%' OR MARCA LIKE '%" & Text13.Text & "%') AND ATENDIDO = '3'"
    End If
    If Check2.Value = 1 Then
        sBuscar = sBuscar & " AND A_DOMICILIO = '1'"
    End If
    If Check3.Value = 1 Then
        sBuscar = sBuscar & " AND GARANTIA = '1'"
    End If
    sBuscar = sBuscar & " ORDER BY ID_AS_TEC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        With tRs
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_AS_TEC") & "")
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE") & ""
                If Not IsNull(.Fields("SUCURSAL")) Then tLi.SubItems(2) = .Fields("SUCURSAL") & ""
                If .Fields("GARANTIA") & "" = 1 Then
                    tLi.SubItems(3) = "SI"
                Else
                    tLi.SubItems(3) = "NO"
                End If
                If Not IsNull(.Fields("FECHA_CAPTURA")) Then tLi.SubItems(4) = .Fields("FECHA_CAPTURA") & ""
                If .Fields("A_DOMICILIO") & "" = 1 Then
                    tLi.SubItems(5) = "SI"
                Else
                    tLi.SubItems(5) = "NO"
                End If
                If Not IsNull(.Fields("FECHA_DEBE_ATENDER")) Then tLi.SubItems(6) = .Fields("FECHA_DEBE_ATENDER") & ""
                If Not IsNull(.Fields("MARCA")) Then tLi.SubItems(7) = .Fields("MARCA") & ""
                If Not IsNull(.Fields("MODELO")) Then tLi.SubItems(8) = .Fields("MODELO") & ""
                If Not IsNull(.Fields("TIPO_ARTICULO")) Then tLi.SubItems(9) = .Fields("TIPO_ARTICULO") & ""
                If Not IsNull(.Fields("Descripcion_PIEZAS")) Then tLi.SubItems(10) = .Fields("Descripcion_PIEZAS") & ""
                If Not IsNull(.Fields("COMENTARIOS_TECNICOS")) Then tLi.SubItems(11) = .Fields("COMENTARIOS_TECNICOS") & ""
                If VarMen.Text1(7).Text = "S" Then
                    If Not IsNull(.Fields("TELEFONO")) Then tLi.SubItems(12) = .Fields("TELEFONO") & ""
                End If
                .MoveNext
            Loop
        End With
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text2.Text = Item.SubItems(1)
    Text3.Text = Item.SubItems(2)
    Text4.Text = Item.SubItems(3)
    Text5.Text = Item.SubItems(4)
    Text6.Text = Item.SubItems(5)
    Text7.Text = Item.SubItems(6)
    Text8.Text = Item.SubItems(7)
    Text9.Text = Item.SubItems(8)
    Text10.Text = Item.SubItems(9)
    Text11.Text = Item.SubItems(10)
    Text12.Text = Item.SubItems(11)
End Sub
Private Sub FunImpATec()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim ConPag As Integer
    ConPag = 1
    Dim sBuscar As String
    sBuscar = "SELECT * FROM ASISTENCIA_TECNICA WHERE ID_AS_TEC = " & DelAST & ""
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\AsTec.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA"
        Set tRs1 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 20, 170, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 205, 20, 170, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "No. Asistencia : " & tRs.Fields("ID_AS_TEC"), "F3", 8, hCenter
        oDoc.WTextBox 70, 340, 20, 250, "Fecha :" & Format(tRs.Fields("FECHA_CAPTURA"), "dd/mm/yyyy"), "F3", 8, hCenter
        
        
        'CAJA1
        'sBuscar = "SELECT * FROM CLIENTE WHERE ID_CLIENTE = " & tRs1.Fields("ID_CLIENTE")
        'Set tRs2 = cnn.Execute(sBuscar)
        oDoc.WTextBox 110, 20, 100, 585, "CLIENTE : " & tRs.Fields("NOMBRE"), "F3", 8, hLeft
        oDoc.WTextBox 120, 20, 100, 585, "TELEFONO : " & tRs.Fields("TELEFONO"), "F3", 8, hLeft
        'If Not (tRs.EOF And tRs.BOF) Then
        '    If Not IsNull(tRs.Fields("NOMBRE")) Then oDoc.WTextBox 110, 20, 100, 400, tRs.Fields("NOMBRE"), "F3", 8, hCenter
        '    If Not IsNull(tRs.Fields("TELEFONO")) Then oDoc.WTextBox 120, 20, 100, 400, tRs.Fields("TELEFONO"), "F3", 8, hCenter
        'End If
        Posi = 150
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 10, 50, "MODELO", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 55, 10, 80, "MARCA", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 135, 10, 280, "TIPO DE ARTICULO", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 415, 10, 60, "GARANTIA", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 475, 10, 50, "DOMICILIO", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 525, 10, 65, "F. COMPROMISO", "F2", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 10
        oDoc.WTextBox Posi, 5, 20, 50, tRs.Fields("MODELO"), "F3", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 55, 20, 80, tRs.Fields("MARCA"), "F3", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 135, 20, 280, tRs.Fields("TIPO_ARTICULO"), "F3", 8, hCenter, , , 1, vbCyan
        If tRs.Fields("GARANTIA") = "1" Then
            oDoc.WTextBox Posi, 415, 20, 60, "SI", "F3", 8, hCenter, , , 1, vbCyan
        Else
            oDoc.WTextBox Posi, 415, 20, 60, "NO", "F3", 8, hCenter, , , 1, vbCyan
        End If
        If tRs.Fields("A_DOMICILIO") = "1" Then
            oDoc.WTextBox Posi, 475, 20, 50, "SI", "F3", 8, hCenter, , , 1, vbCyan
        Else
            oDoc.WTextBox Posi, 475, 20, 50, "NO", "F3", 8, hCenter, , , 1, vbCyan
        End If
        oDoc.WTextBox Posi, 525, 20, 65, Format(tRs.Fields("FECHA_DEBE_ATENDER"), "dd/mm/yyyy"), "F3", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 20
        oDoc.WTextBox Posi, 5, 10, 585, "DESCRIPCION DE PIEZAS", "F2", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 10
        oDoc.WTextBox Posi, 5, 30, 585, tRs.Fields("DESCRIPCION_PIEZAS"), "F3", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 30
        oDoc.WTextBox Posi, 5, 10, 585, "COMENTARIOS TECNICOS", "F2", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 10
        oDoc.WTextBox Posi, 5, 30, 585, tRs.Fields("COMENTARIOS_TECNICOS"), "F3", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 30
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 600
        oDoc.WLineTo 580, 600
        oDoc.LineStroke
        Posi = Posi + 6
        ' cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "No se encontrò la asistencia tècnica solicitada", vbExclamation, "SACC"
    End If
End Sub

