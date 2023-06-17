VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmProduccion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producción"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10680
      TabIndex        =   7
      Top             =   6480
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
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmProduccion.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmProduccion.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmProduccion.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvwTinta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtNumArticulo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCantidad"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtArticulo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtComanda"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSeleccionar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar"
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
         Left            =   8880
         Picture         =   "frmProduccion.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7080
         Width           =   1215
      End
      Begin VB.TextBox txtComanda 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   7200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   7200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   7200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtNumArticulo 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   7200
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ListView lvwTinta 
         Height          =   6495
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   11456
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Private Sub cmdSeleccionar_Click()
On Error GoTo ManejaError
    If Me.txtArticulo.Text = "" Then
        MsgBox "SELECCIONE UN ARTICULO", vbInformation, "SACC"
    Else
        If Me.lvwTinta.SelectedItem.Selected Then
            frmProduccion2.Show vbModal
            Llenar_Lista_Tinta
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
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
    With lvwTinta
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "", 0
        .ColumnHeaders.Add , , "COMANDA", 1200
        .ColumnHeaders.Add , , "ARTICULO", 0
        .ColumnHeaders.Add , , "TIPO", 900
        .ColumnHeaders.Add , , "PRODUCTO", 2440
        .ColumnHeaders.Add , , "Descripcion", 3880
        .ColumnHeaders.Add , , "CANTIDAD", 1200
        .ColumnHeaders.Add , , "ESTADO", 1440
    End With
    Llenar_Lista_Tinta
    'Colorear_Items
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Tinta()
On Error GoTo ManejaError
    Dim nComanda As Integer
    Dim cTipo As String
    sqlQuery = "SELECT * FROM COMANDAS_DETALLES_2 WHERE (ESTADO_ACTUAL = 'R' OR ESTADO_ACTUAL = 'S') AND CANTIDAD <> 0 ORDER BY ID_COMANDA, TIPO"
    Set tRs = cnn.Execute(sqlQuery)
    Me.lvwTinta.ListItems.Clear
    With tRs
        While Not .EOF
            Set tLi = Me.lvwTinta.ListItems.Add(, , .Fields("ID_COMANDA"))
            'PARA NO REPLICAR EN EL LIST VIEW EL NUMERO DE COMANDA
            If nComanda = .Fields("ID_COMANDA") Then
                If Not IsNull(.Fields("ARTICULO")) Then tLi.SubItems(2) = .Fields("ARTICULO")
                If Not IsNull(.Fields("TIPO")) Then
                    'PARA NO REPLICAR EN EL LIST VIEW EL TIPO
                    If Not (cTipo = .Fields("TIPO")) Then
                        cTipo = .Fields("TIPO")
                        If .Fields("TIPO") = "I" Then
                            tLi.SubItems(3) = "TINTA"
                        Else
                            tLi.SubItems(3) = "TONER"
                        End If
                    End If
                End If
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(4) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(6) = .Fields("CANTIDAD")
                If Not IsNull(.Fields("ESTADO_ACTUAL")) Then
                    If .Fields("ESTADO_ACTUAL") = "Z" Then
                        tLi.SubItems(7) = "PENDIENTE"
                    ElseIf .Fields("ESTADO_ACTUAL") = "S" Then
                        tLi.SubItems(7) = "AUTORIZADA"
                    Else
                        tLi.SubItems(7) = ""
                    End If
                End If
            Else
                nComanda = .Fields("ID_COMANDA")
                If Not IsNull(.Fields("ID_COMANDA")) Then tLi.SubItems(1) = .Fields("ID_COMANDA")
                If Not IsNull(.Fields("ARTICULO")) Then tLi.SubItems(2) = .Fields("ARTICULO")
                If Not IsNull(.Fields("TIPO")) Then
                    cTipo = .Fields("TIPO")
                    If .Fields("TIPO") = "I" Then
                        tLi.SubItems(3) = "TINTA"
                    Else
                        tLi.SubItems(3) = "TONER"
                    End If
                End If
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(4) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(6) = .Fields("CANTIDAD")
            End If
            .MoveNext
        Wend
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Colorear_Items()
On Error GoTo ManejaError
    Dim ItMx As ListItem
    Dim lvSI As ListSubItem
    Dim intIndex As Integer
    Dim NoRe As Integer
    Dim Cont As Integer
    Dim nComanda As Integer
    NoRe = Me.lvwTinta.ListItems.Count
    nComanda = Me.lvwTinta.ListItems.Item(Cont)
    For Cont = 1 To NoRe
        If nComanda = Me.lvwTinta.ListItems.Item(Cont) Then
            Set ItMx = Me.lvwTinta.ListItems(Cont)
            ItMx.ForeColor = vbBlue
            For intIndex = 1 To Me.lvwTinta.ColumnHeaders.Count - 3
                Set lvSI = ItMx.ListSubItems(intIndex)
                lvSI.ForeColor = vbBlue
            Next
            DoEvents
        Else
            Set ItMx = Me.lvwTinta.ListItems(Cont)
            ItMx.ForeColor = vbRed
            For intIndex = 1 To Me.lvwTinta.ColumnHeaders.Count - 3
                Set lvSI = ItMx.ListSubItems(intIndex)
                lvSI.ForeColor = vbRed
            Next
            DoEvents
        End If
    Next
    Set ItMx = Nothing
    Set lvSI = Nothing
    Exit Function
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwTinta_DblClick()
On Error GoTo ManejaError
    Me.cmdSeleccionar.Value = True
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwTinta_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Me.txtArticulo.Text = Me.lvwTinta.SelectedItem.SubItems(4)
    Me.txtCantidad.Text = Me.lvwTinta.SelectedItem.SubItems(6)
    Me.txtComanda.Text = Me.lvwTinta.SelectedItem
    Me.txtNumArticulo.Text = Me.lvwTinta.SelectedItem.SubItems(2)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
