VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRequiRapida 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requisición de Orden Rápida"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9600
      TabIndex        =   7
      Top             =   3960
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
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRequiRapida.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRequiRapida.frx":030A
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9600
      TabIndex        =   1
      Top             =   5160
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRequiRapida.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "FrmRequiRapida.frx":1FD6
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label26 
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
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmRequiRapida.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command12"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton Command1 
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
         Height          =   375
         Left            =   8040
         Picture         =   "FrmRequiRapida.frx":40D4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Eliminara todos los articulos marcados con el recuadro"
         Top             =   5640
         Width           =   975
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Agregar"
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
         Left            =   8040
         Picture         =   "FrmRequiRapida.frx":6AA6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Eliminara todos los articulos marcados con el recuadro"
         Top             =   2160
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   360
         TabIndex        =   9
         Top             =   2640
         Width           =   8655
         _ExtentX        =   15266
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
      Begin VB.TextBox Text2 
         Height          =   735
         Left            =   1320
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1260
         Width           =   7695
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   1320
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   420
         Width           =   7695
      End
      Begin VB.Label Label2 
         Caption         =   "Notas :"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción :"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   420
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   9600
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "FrmRequiRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    Dim NoReg As Integer
    Dim Con As Integer
    Con = 1
    NoReg = ListView1.ListItems.Count
    Do While Con <= NoReg
        If ListView1.ListItems(Con).Checked Then
            ListView1.ListItems.Remove (Con)
            NoReg = ListView1.ListItems.Count
        Else
            Con = Con + 1
        End If
    Loop
End Sub
Private Sub Command12_Click()
    Dim tLi As ListItem
    Set tLi = ListView1.ListItems.Add(, , Text1.Text)
    tLi.SubItems(1) = Text2.Text
End Sub
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
        .ColumnHeaders.Add , , "Descripcion", 5500
        .ColumnHeaders.Add , , "Notas", 5500
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image8_Click()
    Dim Con As Integer
    Dim sBuscar As String
    For Con = 1 To ListView1.ListItems.Count
        sBuscar = "INSERT INTO REQUISICION_ORDEN_RAPIDA (DESCRIPCION, NOTAS, ACTIVO, USUARIO) VALUES ('" & Text1.Text & "', '" & Text2.Text & "', 'S', '" & VarMen.Text1(0).Text & "')"
        cnn.Execute (sBuscar)
    Next Con
    Imprime
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ.,;$/_?- *+()&$@1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ.,;$/_?- *+()&$@1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Imprime()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate(App.Path & "\RequiRapida.pdf") Then
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
        oDoc.WTextBox Posi, 10, 60, 570, "Descripcion", "F2", 8, hCenter, , , 1
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 50, 530, ListView1.ListItems(Cont), "F3", 8, hLeft
        Posi = Posi + 50
        oDoc.WTextBox Posi, 10, 60, 570, "Notas", "F2", 8, hCenter, , , 1
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 50, 530, ListView1.ListItems(Cont).SubItems(1), "F3", 8, hLeft
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
