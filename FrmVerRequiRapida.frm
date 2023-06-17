VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmVerRequiRapida 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Requisiciones de Orden Rapida"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   12
      Top             =   4440
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmVerRequiRapida.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmVerRequiRapida.frx":030A
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
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   4
      Top             =   3240
      Width           =   975
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   9
         Top             =   1320
         Width           =   975
         Begin VB.Image Image15 
            Height          =   720
            Left            =   120
            MouseIcon       =   "FrmVerRequiRapida.frx":23EC
            MousePointer    =   99  'Custom
            Picture         =   "FrmVerRequiRapida.frx":26F6
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
            TabIndex        =   10
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   7
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
            TabIndex        =   8
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "FrmVerRequiRapida.frx":40B8
            MousePointer    =   99  'Custom
            Picture         =   "FrmVerRequiRapida.frx":43C2
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   5
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
            TabIndex        =   6
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image17 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmVerRequiRapida.frx":5BEC
            MousePointer    =   99  'Custom
            Picture         =   "FrmVerRequiRapida.frx":5EF6
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image18 
         Height          =   750
         Left            =   120
         MouseIcon       =   "FrmVerRequiRapida.frx":79A8
         MousePointer    =   99  'Custom
         Picture         =   "FrmVerRequiRapida.frx":7CB2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   2
      Top             =   2040
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmVerRequiRapida.frx":99DC
         MousePointer    =   99  'Custom
         Picture         =   "FrmVerRequiRapida.frx":9CE6
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reporte"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmVerRequiRapida.frx":A275
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.ListView ListView1 
         Height          =   4935
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8705
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
   Begin VB.Image Image1 
      Height          =   375
      Left            =   8640
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "FrmVerRequiRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Form_Load()
    On Error GoTo ManejaError
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    Dim sBuscar As String
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
        .Checkboxes = True
        .ColumnHeaders.Add , , "Requisicion", 0
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Notas", 5500
        .ColumnHeaders.Add , , "Fecha de Alta", 1500
        .ColumnHeaders.Add , , "Usuario", 2000
    End With
    Buscar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Imprime()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate(App.Path & "\RequiRapidaPendientes.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
    ' Encabezado del reporte
    ' asi se agregan los logos... solo te falto poner un control IMAGE1 para cargar la imagen en el
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    oDoc.LoadImage Image1, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 70, 40, 43, 161, "Logo"
    sBuscar = "SELECT * FROM EMPRESA  "
    Set tRs1 = cnn.Execute(sBuscar)
    oDoc.WTextBox 40, 205, 100, 175, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
    oDoc.WTextBox 60, 224, 100, 175, tRs1.Fields("DIRECCION"), "F3", 8, hLeft
    oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs1.Fields("COLONIA"), "F3", 8, hLeft
    oDoc.WTextBox 70, 205, 100, 175, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
    oDoc.WTextBox 80, 205, 100, 175, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
    oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
    ' ENCABEZADO DEL DETALLE
    oDoc.WTextBox 100, 205, 100, 175, "Requisicion de Orden Rapida", "F2", 12, hCenter
    Posi = 120
    oDoc.WTextBox Posi, 10, 20, 500, "Usuario : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text, "F2", 8, hLeft
    Posi = Posi + 12
    ' Linea
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, Posi
    oDoc.WLineTo 580, Posi
    oDoc.LineStroke
    Posi = Posi + 6
    ' DETALLE
    For Cont = 1 To ListView1.ListItems.Count
        oDoc.WTextBox Posi, 10, 60, 570, "Descripcion", "F2", 8, hLeft, , , 1
        oDoc.WTextBox Posi, 250, 60, 200, ListView1.ListItems(Cont).SubItems(4) & " el dia " & ListView1.ListItems(Cont).SubItems(3), "F2", 8, hLeft
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 50, 530, ListView1.ListItems(Cont).SubItems(1), "F3", 8, hLeft
        Posi = Posi + 50
        oDoc.WTextBox Posi, 10, 60, 570, "Notas", "F2", 8, hLeft, , , 1
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 50, 530, ListView1.ListItems(Cont).SubItems(2), "F3", 8, hLeft
        Posi = Posi + 60
        Posi = Posi + 12
        If Posi >= 700 Then
            oDoc.NewPage A4_Vertical
            oDoc.WImage 70, 40, 43, 161, "Logo"
            sBuscar = "SELECT * FROM EMPRESA"
            Set tRs1 = cnn.Execute(sBuscar)
            oDoc.WTextBox 40, 205, 100, 175, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
            oDoc.WTextBox 60, 224, 100, 175, tRs1.Fields("DIRECCION"), "F3", 8, hLeft
            oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs1.Fields("COLONIA"), "F3", 8, hLeft
            oDoc.WTextBox 70, 205, 100, 175, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 80, 205, 100, 175, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
            oDoc.WTextBox 40, 380, 20, 250, "Fecha de Impresion: " & Date, "F3", 8, hCenter
            ' ENCABEZADO DEL DETALLE
            oDoc.WTextBox 100, 205, 100, 175, "Requisicion de Orden Rapida", "F3", 8, hCenter
            Posi = 120
            oDoc.WTextBox Posi, 10, 20, 500, "Usuario : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text, "F2", 8, hLeft
            Posi = Posi + 12
            ' Linea
            oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, Posi
            oDoc.WLineTo 580, Posi
            oDoc.LineStroke
            Posi = Posi + 6
        End If
    Next Cont
    ' Linea
    Posi = Posi + 6
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, Posi
    oDoc.WLineTo 580, Posi
    oDoc.LineStroke
     Posi = Posi + 16
    ' TEXTO ABAJO
    Posi = Posi + 16
    oDoc.WTextBox Posi, 205, 100, 175, "COMENTARIOS", "F3", 8, hCenter
    Posi = Posi + 20
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, Posi
    oDoc.WLineTo 580, Posi
    oDoc.LineStroke
    Posi = Posi + 16
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, Posi
    oDoc.WLineTo 580, Posi
    oDoc.LineStroke
    Posi = Posi + 16
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, Posi
    oDoc.WLineTo 580, Posi
    oDoc.LineStroke
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub Image18_Click()
    Dim Con As Integer
    Dim sBuscar As String
    For Con = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(Con).Checked Then
            sBuscar = "UPDATE REQUISICION_ORDEN_RAPIDA SET ACTIVO = 'N' WHERE ID_REQUISICION = " & ListView1.ListItems(Con)
            cnn.Execute (sBuscar)
        End If
    Next Con
    Buscar
End Sub
Private Sub Image26_Click()
    Imprime
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Buscar()
    Dim tRs As ADODB.Recordset
    Dim sBsucar As String
    Dim tLi As ListItem
    sBuscar = "SELECT * FROM REQUISICION_ORDEN_RAPIDA WHERE ACTIVO = 'S' ORDER BY ID_REQUISICION"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_REQUISICION"))
            tLi.SubItems(1) = tRs.Fields("DESCRIPCION")
            tLi.SubItems(2) = tRs.Fields("NOTAS")
            tLi.SubItems(3) = tRs.Fields("FECHA_ALTA")
            tLi.SubItems(4) = tRs.Fields("USUARIO")
            tRs.MoveNext
        Loop
    End If
End Sub
