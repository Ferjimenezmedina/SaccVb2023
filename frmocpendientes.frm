VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmordpendiente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Orde de Compra"
   ClientHeight    =   4590
   ClientLeft      =   1530
   ClientTop       =   -1560
   ClientWidth     =   12375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11280
      TabIndex        =   16
      Top             =   3240
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
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmocpendientes.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmocpendientes.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Detalle de Orde de Compra"
      TabPicture(0)   =   "frmocpendientes.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fecha"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "art"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "oc"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtNumArticulo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txttipo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtNoSirve"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3480
         TabIndex        =   15
         Top             =   3840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Excel"
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
         Left            =   9600
         Picture         =   "frmocpendientes.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   9840
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   10200
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtNoSirve 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   3720
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txttipo 
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   3720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame1 
         Height          =   2895
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   10815
         Begin MSComctlLib.ListView lvwJR 
            Height          =   2535
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   4471
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.TextBox txtNumArticulo 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   3720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "O.C"
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
         Left            =   240
         TabIndex        =   11
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Proveedor"
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
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label oc 
         Caption         =   "Comanda"
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
         Left            =   1440
         TabIndex        =   9
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label art 
         Caption         =   "Articulo"
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
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   7095
      End
      Begin VB.Label fecha 
         Caption         =   "Cantidad"
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
         Left            =   1440
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha"
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
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11400
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmordpendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim StrRep As String
Dim StrRep2 As String
Dim StrRep3 As String
Dim StrRep4 As String
Dim StrRep5 As String
Dim orden As Integer
Private Sub Command1_Click()
    If lvwJR.ListItems.Count > 0 Then
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
        StrCopi = "O.C" & Chr(9) & "ID_PRODUCTO" & Chr(9) & "Descripcion" & Chr(9) & "CANTIDAD" & Chr(9) & " PRECIO" & Chr(9) & "SURTIDO" & Chr(9) & " STATUS" & Chr(13)
        If Ruta <> "" Then
            NumColum = lvwJR.ColumnHeaders.Count
            For Con = 1 To lvwJR.ListItems.Count
                StrCopi = StrCopi & lvwJR.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & lvwJR.ListItems.Item(Con).SubItems(Con2) & Chr(9)
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
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Text1.Text = Frmrepcomprass.Text5.Text
    Text2.Text = Frmrepcomprass.Text6.Text
    With lvwJR
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "O.C", 1000
        .ColumnHeaders.Add , , "ID_PRODUCTO", 1500
        .ColumnHeaders.Add , , "Descripcion", 2500
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "PRECIO", 1000
        .ColumnHeaders.Add , , "SURTIDO", 1000
        .ColumnHeaders.Add , , "STATUS", 3000
    End With
    proordenes
End Sub
Private Sub proordenes()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs3 As ADODB.Recordset
    Dim orde As Integer
    Dim tip As String
    Dim pro As String
    lvwJR.ListItems.Clear
    Dim ordeee As Integer
    sBuscar = "SELECT * FROM vsordencom WHERE  TIPO='" & Text2.Text & "' AND  NUM_ORDEN= '" & Text1.Text & "'     "
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = lvwJR.ListItems.Add(, , tRs.Fields("NUM_ORDEN"))
            orde = tRs.Fields("NUM_ORDEN")
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(2) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(3) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("PRECIO")) Then tLi.SubItems(4) = tRs.Fields("PRECIO")
            If Not IsNull(tRs.Fields("SURTIDO")) Then tLi.SubItems(5) = tRs.Fields("SURTIDO")
            If tRs.Fields("CONFIRMADA") = "N" Then
                tLi.SubItems(6) = "PRE-ORDEN"
            End If
            If tRs.Fields("CONFIRMADA") = "P" Then
                tLi.SubItems(6) = "PENDIENTE DE AUTORIZAR"
            End If
            If tRs.Fields("CONFIRMADA") = "S" Then
                tLi.SubItems(6) = "PENDIENTE DE IMPRIMIR"
            End If
            sBuscar = "SELECT * FROM vsordpende WHERE NUM_ORDEN= '" & tRs.Fields("NUM_ORDEN") & "' AND  TIPO= '" & tRs.Fields("TIPO") & "' AND  ID_PRODUCTO= '" & tRs.Fields("ID_PRODUCTO") & "'"
            Set tRs3 = cnn.Execute(sBuscar)
                Dim catpe  As Double
                catpe = CDbl(tRs3.Fields("CANTIDAD")) - CDbl(tRs3.Fields("SURTIDO"))
            If tRs.Fields("CONFIRMADA") = "X" And tRs3.Fields("SURTIDO") = 0 Then
                tLi.SubItems(6) = "PENDIENTE DE  LLEGAR "
            End If
            If tRs.Fields("CONFIRMADA") = "X" And catpe = 0 Then
                tLi.SubItems(6) = "PENDIENTE DE PAGO/EN ALMACEN"
            End If
            If tRs.Fields("CONFIRMADA") = "X" And catpe < tRs3.Fields("CANTIDAD") And tRs3.Fields("SURTIDO") < 0 Then
          
                tLi.SubItems(6) = "PENDIENTE DE PAGO/LLEGADA PARCIAL"
            End If
            If tRs.Fields("CONFIRMADA") = "Y" Then
                tLi.SubItems(6) = "PAGADA"
            End If
            catpe = 0
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
