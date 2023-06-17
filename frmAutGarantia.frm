VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmAutGarantia 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Autorizar Garantia"
   ClientHeight    =   5895
   ClientLeft      =   3255
   ClientTop       =   2625
   ClientWidth     =   9870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   9870
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8640
      TabIndex        =   21
      Top             =   4440
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmAutGarantia.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmAutGarantia.frx":030A
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
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmAutGarantia.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCosto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTipo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label13"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label12"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lvwGarantia"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Check2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Check1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtComent"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   4200
         TabIndex        =   25
         Top             =   4800
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   195
         Left            =   8130
         TabIndex        =   23
         Top             =   4920
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.TextBox txtComent 
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   3840
         Width           =   3375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Trae mas del 50% de su capacidad de Tinta/Toner"
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   3360
         Width           =   3975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Trae etiqueta/marcado con numero de comanda"
         Height          =   195
         Left            =   3960
         TabIndex        =   3
         Top             =   3720
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Autorizar"
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
         Left            =   4320
         Picture         =   "frmAutGarantia.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Denegar"
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
         Left            =   5880
         Picture         =   "frmAutGarantia.frx":4DDA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3960
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwGarantia 
         Height          =   1935
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   3413
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label12 
         Caption         =   "Comentario de Produccion"
         Height          =   255
         Left            =   4800
         TabIndex        =   24
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   8055
      End
      Begin VB.Label Label2 
         Caption         =   "# VENTA:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Venta:"
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "IDC"
         Height          =   255
         Left            =   6480
         TabIndex        =   17
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "IDC"
         Height          =   255
         Left            =   7320
         TabIndex        =   16
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Codigo del Articulo"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Comentario"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Cantidad del Articulo"
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   6000
         TabIndex        =   12
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label8 
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblTipo 
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   3120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblCosto 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   1335
      Left            =   8640
      TabIndex        =   26
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2355
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmAutGarantia.frx":77AC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label23"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label24"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label25"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label26"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label28"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label29"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
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
         Left            =   120
         Picture         =   "frmAutGarantia.frx":77C8
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   195
         Left            =   8130
         TabIndex        =   27
         Top             =   4920
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label14 
         Caption         =   "Venta :"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label29 
         Height          =   375
         Left            =   2760
         TabIndex        =   33
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label28 
         Height          =   255
         Left            =   2160
         TabIndex        =   32
         Top             =   3120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label26 
         Height          =   255
         Left            =   1560
         TabIndex        =   31
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label25 
         Height          =   255
         Left            =   960
         TabIndex        =   30
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label24 
         Height          =   255
         Left            =   3240
         TabIndex        =   29
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label23 
         Height          =   255
         Left            =   6000
         TabIndex        =   28
         Top             =   3000
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmAutGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    Dim sqlQuery As String
     Dim sBuscar As String
    Dim NoRe As Integer
    Dim Cont As Integer
    Dim nComanda As Integer
    Dim cTipo As String
    Dim IDREP As String
    Dim Tipo As String
    Dim sql As String
    Dim catser As Integer
    If (Check1 = 1) Or (Check2 = 1) Then
        If Text2.Text = "" Then
            MsgBox ("!DEBE  INGRESAR UN COMENTARIO PRODUCCION PARA PODER CONTINUAR¡")
            MsgBox ("!DEBE  INGRESAR UN COMENTARIO PRODUCCION PARA PODER CONTINUAR¡")
        Else
            catser = InputBox("Ingrese la Cantidad que  Se Acepto:", "")
            If Check1 = 1 Then
                NoRe = Me.lvwGarantia.ListItems.Count
                For Cont = 1 To NoRe
                    If Me.lvwGarantia.ListItems.Item(Cont).Checked = True Then
                        sqlQuery = "UPDATE GARANTIAS SET ESTADO = 'A' WHERE ID_GARANTIA='" & Label13.Caption & "'"
                        cnn.Execute (sqlQuery)
                        sBuscar = "UPDATE GARANTIAS SET CANT_ACEP = '" & catser & "' WHERE ID_GARANTIA='" & Label13.Caption & "'"
                        cnn.Execute (sBuscar)
                        sql = "UPDATE GARANTIAS SET REVUNO = 'Trae mas del 50% de su capacidad de Tinta/Toner' , COMEN='" & Text2.Text & "' WHERE ID_GARANTIA='" & Label13.Caption & "'"
                        cnn.Execute (sql)
                    End If
                Next Cont
            End If
            If Check2 = 1 Then
                NoRe = Me.lvwGarantia.ListItems.Count
                For Cont = 1 To NoRe
                    If Me.lvwGarantia.ListItems.Item(Cont).Checked = True Then
                        sqlQuery = "UPDATE GARANTIAS SET ESTADO = 'A' WHERE ID_GARANTIA='" & Label13.Caption & "' "
                        cnn.Execute (sqlQuery)
                        sBuscar = "UPDATE GARANTIAS SET CANT_ACEP = '" & catser & "' WHERE ID_GARANTIA='" & Label13.Caption & "'"
                        cnn.Execute (sBuscar)
                        sql = "UPDATE GARANTIAS SET REVDOS = 'Trae etiqueta/marcado con numero de comanda', COMEN='" & Text2.Text & "' WHERE ID_GARANTIA='" & Label13.Caption & "'"
                        cnn.Execute (sql)
                    End If
                Next Cont
            End If
            If (Check2 = 1) And (Check1 = 1) Then
                NoRe = Me.lvwGarantia.ListItems.Count
                For Cont = 1 To NoRe
                    If Me.lvwGarantia.ListItems.Item(Cont).Checked = True Then
                        sqlQuery = "UPDATE GARANTIAS SET ESTADO = 'A' WHERE ID_GARANTIA='" & Label13.Caption & "'"
                        cnn.Execute (sqlQuery)
                        sql = "UPDATE GARANTIAS SET REVUNO = 'Trae mas del 50% de su capacidad de Tinta/Toner' WHERE ID_GARANTIA='" & Label13.Caption & "'"
                        cnn.Execute (sql)
                        sql = "UPDATE GARANTIAS SET REVDOS = 'Trae etiqueta/marcado con numero de comanda',COMEN='" & Text2.Text & "' WHERE ID_GARANTIA='" & Label13.Caption & "'"
                        cnn.Execute (sql)
                        sBuscar = "UPDATE GARANTIAS SET CANT_ACEP = '" & catser & "' WHERE ID_GARANTIA='" & Label13.Caption & "'"
                        cnn.Execute (sBuscar)
                    End If
                Next Cont
            End If
            listadodecomandas.Text2 = Label13.Caption
            listadodecomandas.Text1 = Label8.Caption
            listadodecomandas.Text3 = catser
            Busca
            listadodecomandas.Show vbModal
        End If
    End If
End Sub
Private Sub Command2_Click()
Dim tRs As ADODB.Recordset
    Dim sBus As String
    Dim tLi As ListItem
    lvwGarantia.ListItems.Clear
    If Text3.Text = "" Then
        MsgBox "INGRESE UN NUMERO DE VENTA!", vbInformation, "SACC"
    Else
        sBus = "SELECT * FROM GARANTIAS WHERE ID_VENTA='" & Text3.Text & "' AND ESTADO = 'P'"
        Set tRs = cnn.Execute(sBus)
        If tRs.EOF And tRs.BOF Then
            MsgBox "No hay garantias pendientes"
        Else
            With tRs
                Do While Not (.EOF)
                    Set tLi = lvwGarantia.ListItems.Add(, , .Fields("ID_VENTA") & "")
                    tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                    tLi.SubItems(2) = .Fields("ID_PRODUCTO") & ""
                    tLi.SubItems(3) = .Fields("FECHA") & ""
                    tLi.SubItems(4) = .Fields("IDDET") & ""
                    tLi.SubItems(5) = .Fields("COMENTARIO") & ""
                    tLi.SubItems(6) = .Fields("TIPO") & ""
                    tLi.SubItems(7) = .Fields("ID_GARANTIA") & ""
                    .MoveNext
                Loop
            End With
        End If
    End If
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim sql As String
    Dim sqlQuery As String
    NoRe = Me.lvwGarantia.ListItems.Count
    For Cont = 1 To NoRe
        If Me.lvwGarantia.ListItems.Item(Cont).Checked = True Then
            sqlQuery = "UPDATE GARANTIAS SET ESTADO = 'N'  WHERE ID_GARANTIA='" & Label13.Caption & "'"
            cnn.Execute (sqlQuery)
            Label1.Caption = ""
            Label3.Caption = ""
            Label5.Caption = ""
            Label6.Caption = ""
            Label8.Caption = ""
            Label10.Caption = ""
            If Check1.Value = 1 Then
              sql = "UPDATE GARANTIAS SET REVUNO = 'Trae mas del 50% de su capacidad de Tinta/Toner',COMEN='" & Text2.Text & "' WHERE ID_GARANTIA='" & Label13.Caption & "'"
              cnn.Execute (sql)
            End If
            If Check2.Value = 1 Then
              sql = "UPDATE GARANTIAS SET REVDOS = 'Trae etiqueta/marcado con numero de comanda',COMEN='" & Text2.Text & "' WHERE ID_GARANTIA='" & Label13.Caption & "'"
              cnn.Execute (sql)
            End If
            Busca
        End If
    Next Cont
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
On Error GoTo ManejaError
    Dim tRs As ADODB.Recordset
    Dim sBus As String
    Dim tLi As ListItem
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    With lvwGarantia
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "#VENTA", 1200
        .ColumnHeaders.Add , , "CANTIDAD", 1200
        .ColumnHeaders.Add , , "ID_PROD", 2200
        .ColumnHeaders.Add , , "FECHA", 1300
        .ColumnHeaders.Add , , "IDDET", 0
        .ColumnHeaders.Add , , "Coment", 2400
        .ColumnHeaders.Add , , "TIPO", 0
        .ColumnHeaders.Add , , "ID_G", 0
    End With
    Set cnn = New ADODB.Connection
    With cnn
          .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Busca
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Busca()
Dim tRs As ADODB.Recordset
    Dim sBus As String
    Dim tLi As ListItem
    
     sBus = "SELECT * FROM GARANTIAS WHERE ESTADO = 'P' ORDER BY FECHA DESC "
    Set tRs = cnn.Execute(sBus)
    If tRs.EOF And tRs.BOF Then
        MsgBox "No hay garantias pendientes"
    Else
        With tRs
            Do While Not (.EOF)
                Set tLi = lvwGarantia.ListItems.Add(, , .Fields("ID_VENTA") & "")
                tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                tLi.SubItems(2) = .Fields("ID_PRODUCTO") & ""
                tLi.SubItems(3) = .Fields("FECHA") & ""
                tLi.SubItems(4) = .Fields("IDDET") & ""
                tLi.SubItems(5) = .Fields("COMENTARIO") & ""
                tLi.SubItems(6) = .Fields("TIPO") & ""
                tLi.SubItems(7) = .Fields("ID_GARANTIA") & ""
                .MoveNext
            Loop
        End With
    End If


End Sub

Private Sub Label21_Click()
End Sub
Private Sub lvwGarantia_DblClick()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim Dia As Date
    sBuscar = "SELECT * FROM VSFACTURA WHERE ID_VENTA = '" & lvwGarantia.SelectedItem & "' AND IDDET = '" & lvwGarantia.SelectedItem.ListSubItems(4) & "'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        Label1.Caption = .Fields("Nombre")
        Label3.Caption = .Fields("ID_VENTA")
        Label5.Caption = .Fields("FECHA")
        Label6.Caption = .Fields("ID_CLIENTE")
        Label8.Caption = lvwGarantia.SelectedItem.ListSubItems(2) 'ID PRODUCTO
        Label10.Caption = lvwGarantia.SelectedItem.ListSubItems(1) 'CANTIDAD
        Label13.Caption = lvwGarantia.SelectedItem.ListSubItems(7) 'ID_GARANTIA
        txtComent.Text = lvwGarantia.SelectedItem.ListSubItems(5) 'COMENTARIO
        lblTipo.Caption = lvwGarantia.SelectedItem.ListSubItems(6) 'TIPO
        lblCosto.Caption = .Fields("PRECIO_VENTA")
        Command1.Enabled = True
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If Text3.Text = "" Then
        MsgBox "INGRESE UN NUMERO DE VENTA!", vbInformation, "SACC"
    Else
        If KeyAscii = 13 Then
            Me.Command2.Value = True
        End If
    End If
End Sub
Private Sub comand()
    Dim oDoc  As cPDF
    ListView3.ListItems(Con).SubItems (3)
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim loca As Integer
    Dim sBuscar As String
    Dim tRs1 As ADODB.Recordset
    Dim PosVer As Integer
    Dim Posabo As Integer
    Dim tRs  As ADODB.Recordset
    Dim tRs2  As ADODB.Recordset
    Dim tRs4  As ADODB.Recordset
    Set oDoc = New cPDF
    Dim sumdeuda As Double
    Dim sumIndi As Double
    Dim sNombre As String
    Dim sumabonos As Double
    Dim sumIndiabonos As Double
    Dim ConPag As Integer
    ConPag = 1
    If Not oDoc.PDFCreate(App.Path & "\Juegoreparacion.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
' Encabezado del reporte
    oDoc.NewPage A4_Vertical
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 8, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 8, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 8, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 8, hCenter
    oDoc.WTextBox 90, 200, 20, 250, "Juego de Reparacion Para Comandas", "F2", 10, hCenter
    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
' Encabezado de pagina
    oDoc.WTextBox 100, 90, 40, 200, "Cantidad en Juego", "F2", 10, hCenter
    oDoc.WTextBox 100, 250, 40, 300, "Cantidad  en comanda", "F2", 10, hLeft
' Cuerpo del reporte
    If MsgBox("ESTA SEGURO QUE  QUIERE IMPRIMIR EL JUEGO DE REPARACION. " & ListView3.ListItems(Con).SubItems(2) & "?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
        sBuscar = "SELECT * FROM  IMPCOMAN WHERE ID_COMANDA='" & Text1.Text & "'"
        Set tRs4 = cnn.Execute(sBuscar)
        If Not (tRs4.EOF And tRs4.BOF) Then
            oDoc.WTextBox 100, 500, 40, 300, "Reimpresion", "F2", 10, hLeft
        Else
            sBuscar = "INSERT INTO IMPCOMAN (ID_COMANDA,FECHA) VALUES ('" & Text1.Text & "','" & Format(Date, "dd/mm/yyyy") & "');"
            cnn.Execute (sBuscar)
            oDoc.WTextBox 100, 500, 40, 300, "Impresion", "F2", 10, hLeft
        End If
        sumdeuda = 0
        sBuscar = "SELECT ID_PRODUCTO,CANTIDAD,ID_COMANDA FROM VENTAS_DETALLE WHERE ID_COMANDA='" & Text1.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        oDoc.WTextBox 130, 30, 30, 60, "COMANDA :", "F2", 10, hLeft
        oDoc.WTextBox 130, 100, 30, 40, tRs.Fields("ID_COMANDA"), "F2", 10, hLeft
        Posi = 130
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 100
        oDoc.WLineTo 580, 100
        oDoc.LineStroke
        oDoc.MoveTo 10, 125
        oDoc.WLineTo 580, 125
        oDoc.LineStroke
        Posi = Posi + 15
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                sBuscar = "SELECT ID_REPARACION,ID_PRODUCTO,CANTIDAD FROM JUEGO_REPARACION  WHERE ID_REPARACION='" & tRs.Fields("ID_PRODUCTO") & "' GROUP BY  ID_REPARACION,ID_PRODUCTO,CANTIDAD"
                Set tRs1 = cnn.Execute(sBuscar)
                If Not (tRs1.EOF And tRs1.BOF) Then
                    Do While Not (tRs1.EOF)
                        If sNombre <> tRs1.Fields("ID_REPARACION") Then
                            Posi = Posi + 15
                            oDoc.WTextBox Posi, 30, 40, 200, tRs1.Fields("ID_REPARACION"), "F2", 11, hLeft
                            oDoc.WTextBox Posi, 150, 50, 100, tRs.Fields("CANTIDAD"), "F2", 11, hLeft
                        End If
                        Posi = Posi + 10
                        oDoc.WTextBox Posi, 30, 40, 300, tRs1.Fields("ID_PRODUCTO"), "F2", 9, hLeft
                        oDoc.WTextBox Posi, 200, 40, 80, tRs1.Fields("CANTIDAD"), "F2", 9, hLeft
                        sumdeuda = CDbl(tRs.Fields("CANTIDAD") * tRs1.Fields("CANTIDAD"))
                        oDoc.WTextBox Posi, 300, 40, 300, sumdeuda, "F2", 10, hLeft
                        sNombre = tRs1.Fields("ID_REPARACION")
                        sumdeuda = 0
                        tRs1.MoveNext
                    Loop
                End If
                Posi = Posi + 15
                tRs.MoveNext
                If Posi >= 740 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    oDoc.NewPage A4_Vertical
                    ' Encabezado del reporte
                    Posi = 140
                    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                    oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Pagar", "F2", 10, hCenter
                    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
                    oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
                    ' Encabezado de pagina
                    oDoc.WTextBox 100, 20, 30, 40, "Id", "F2", 10, hCenter
                    oDoc.WTextBox 100, 30, 30, 80, "Factura", "F2", 10, hCenter
                    oDoc.WTextBox 100, 50, 50, 160, "Fecha", "F2", 10, hCenter
                    oDoc.WTextBox 100, 90, 40, 200, "Total-Fac", "F2", 10, hCenter
                    oDoc.WTextBox 100, 250, 40, 300, "Abono", "F2", 10, hLeft
                    oDoc.WTextBox 100, 350, 40, 200, "Pendiente", "F2", 10, hLeft
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 10, 100
                    oDoc.WLineTo 580, 100
                    oDoc.LineStroke
                    oDoc.MoveTo 10, 125
                    oDoc.WLineTo 580, 125
                    oDoc.LineStroke
                End If
            Loop
            Posi = Posi + 30
            Cont = Cont + 1
        End If
    End If
    oDoc.PDFClose
    oDoc.Show
End Sub
