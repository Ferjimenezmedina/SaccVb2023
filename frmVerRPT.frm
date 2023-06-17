VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmVerRPT 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   9840
      TabIndex        =   11
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9840
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmVerRPT.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin MSComctlLib.ListView ListView1 
         Height          =   6375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   11245
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   9375
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   5
      Top             =   4920
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
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmVerRPT.frx":001C
         MousePointer    =   99  'Custom
         Picture         =   "frmVerRPT.frx":0326
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10080
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   3
      Top             =   6120
      Width           =   975
      Begin VB.Image cmdCancelar 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmVerRPT.frx":1E68
         MousePointer    =   99  'Custom
         Picture         =   "frmVerRPT.frx":2172
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label9 
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox txtSQL 
      Height          =   375
      Index           =   2
      Left            =   9840
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtSQL 
      Height          =   375
      Index           =   1
      Left            =   9840
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   10080
      Top             =   2880
   End
   Begin VB.TextBox txtSQL 
      Height          =   375
      Index           =   0
      Left            =   9840
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmVerRPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
'FUNCION PARA MANEJAR EL ANCHO DE LAS COLUMNAS DEL LISTVIEW1
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_SETCOLUMNWIDTH = &H101E
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
End Sub
Private Sub Image10_Click()
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
    If ListView1.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.Count
            For Con = 1 To ListView1.ColumnHeaders.Count
                StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
            Next
            ProgressBar1.Value = 0
            ProgressBar1.Visible = True
            ProgressBar1.Min = 0
            ProgressBar1.Max = ListView1.ListItems.Count
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView1.ListItems.Count
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
                ProgressBar1.Value = Con
            Next
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ProgressBar1.Visible = False
        ProgressBar1.Value = 0
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
Private Sub Timer1_Timer()
    Dim tLi As ListItem
    Dim tRs As ADODB.Recordset
    Dim Cont As Integer
    Dim i As Integer
    If txtSQL(0).Text <> "" Then
        If Reportes1.Combo1.Text = "" Then
            txtSQL(0).Text = Replace(txtSQL(0).Text, "order by sucursal", "")
        End If
        MsgBox txtSQL(0).Text
        Set tRs = cnn.Execute(txtSQL(0).Text)
    ElseIf txtSQL(1).Text <> "" Then
        Set tRs = cnn.Execute(txtSQL(1).Text)
    ElseIf txtSQL(2).Text <> "" Then
        Set tRs = cnn.Execute(txtSQL(2).Text)
    End If
    If Not (tRs.EOF And tRs.BOF) Then
        With ListView1
            .View = lvwReport
            .Gridlines = True
            .LabelEdit = lvwManual
            .HideSelection = False
            .HotTracking = False
            .HoverSelection = False
            For Cont = 0 To tRs.Fields.Count - 1
                .ColumnHeaders.Add , , tRs.Fields.Item(Cont).Name, 1000
            Next Cont
        End With
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields(0))
            For Cont = 1 To tRs.Fields.Count - 1
                If Not IsNull(tRs.Fields(Cont)) Then
                    tLi.SubItems(Cont) = tRs.Fields(Cont)
                End If
                Next Cont
            tRs.MoveNext
        Loop
        tRs.Close
        If txtSQL(1).Text <> "" Then
            Set tRs = cnn.Execute(txtSQL(1).Text)
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not (tRs.EOF)
                    Set tLi = ListView1.ListItems.Add(, , tRs.Fields(0))
                    For Cont = 1 To tRs.Fields.Count - 1
                        tLi.SubItems(Cont) = tRs.Fields(Cont)
                    Next Cont
                    tRs.MoveNext
                Loop
                tRs.Close
            End If
        End If
        If txtSQL(2).Text <> "" Then
            Set tRs = cnn.Execute(txtSQL(2).Text)
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not (tRs.EOF)
                    Set tLi = ListView1.ListItems.Add(, , tRs.Fields(0))
                    For Cont = 1 To tRs.Fields.Count - 1
                        If Cont <> 0 Then
                            If Not IsNull(tRs.Fields(Cont)) Then tLi.SubItems(Cont) = tRs.Fields(Cont)
                        End If
                    Next Cont
                    tRs.MoveNext
                Loop
                tRs.Close
            End If
        End If
    End If
    If ListView1.ListItems.Count = 0 Then
        With ListView1
            .View = lvwReport
            .Gridlines = True
            .LabelEdit = lvwManual
            .ColumnHeaders.Add , , "RESULTADOS", 4000
        End With
        Set tLi = ListView1.ListItems.Add(, , "BUSQUEDA SIN RESULTADOS")
    End If
    Timer1.Enabled = False
'CAMBIO DEL ANCHO DE COLUMNAS DEL LISTVIEW1
    For i = 0 To ListView1.ColumnHeaders.Count - 1
        SendMessage ListView1.hWnd, LVM_SETCOLUMNWIDTH, i, -3
    Next i
End Sub
