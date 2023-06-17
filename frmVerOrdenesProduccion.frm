VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmVerOrdenesProduccion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Ordenes de Producción"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   -2147483644
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmVerOrdenesProduccion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblEstado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvwOrdenes"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin MSComctlLib.ListView lvwOrdenes 
         Height          =   3615
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6376
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lblEstado 
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
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3720
         Width           =   4215
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4920
      TabIndex        =   3
      Top             =   1920
      Width           =   975
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Image cmdTraer 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmVerOrdenesProduccion.frx":001C
         MousePointer    =   99  'Custom
         Picture         =   "frmVerOrdenesProduccion.frx":0326
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4920
      TabIndex        =   1
      Top             =   3120
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmVerOrdenesProduccion.frx":1F28
         MousePointer    =   99  'Custom
         Picture         =   "frmVerOrdenesProduccion.frx":2232
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
End
Attribute VB_Name = "frmVerOrdenesProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Private Sub cmdTraer_Click()
    frmReviComa.txtCantidadComanda.Text = Me.lvwOrdenes.SelectedItem
    frmReviComa.cmdTraer.Value = True
    Unload Me
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
    With lvwOrdenes
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Orden", 1000
        .ColumnHeaders.Add , , "Fecha", 1440
        .ColumnHeaders.Add , , "Agente", 1440
    End With
    If Hay_Ordenes Then
        Llenar_Lista_Ordenes
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Hay_Ordenes() As Boolean
On Error GoTo ManejaError
    Me.lblEstado.Caption = "Buscando ordenes de producción"
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    sqlQuery = "SELECT COUNT(ID_COMANDA)ID_COMANDA FROM COMANDAS_2 WHERE ESTADO_ACTUAL = 'A' AND TIPO = 'P'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_COMANDA") <> 0 Then
            Hay_Ordenes = True
            Me.lblEstado.Caption = "Se encontraron " & .Fields("ID_COMANDA") & " ordenes"
            Me.lblEstado.ForeColor = vbBlue
            DoEvents
        Else
            Hay_Ordenes = False
            Me.lblEstado.Caption = "No se encontraron ordenes"
            Me.lblEstado.ForeColor = vbRed
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Lista_Ordenes()
On Error GoTo ManejaError
    sqlQuery = "SELECT C.ID_COMANDA, C.FECHA_INICIO, U.NOMBRE + ' ' + U.APELLIDOS AS NOMBRE FROM COMANDAS_2 AS C JOIN USUARIOS AS U ON C.ID_AGENTE = U.ID_USUARIO WHERE C.ESTADO_ACTUAL = 'A' AND C.TIPO = 'P' ORDER BY C.ID_COMANDA"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        Me.lvwOrdenes.ListItems.Clear
        Do While Not .EOF
            Set tLi = Me.lvwOrdenes.ListItems.Add(, , .Fields("ID_COMANDA"))
            If Not IsNull(.Fields("FECHA_INICIO")) Then tLi.SubItems(1) = .Fields("FECHA_INICIO")
            If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(2) = .Fields("NOMBRE")
            .MoveNext
        Loop
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwOrdenes_DblClick()
    If Me.lvwOrdenes.SelectedItem.Selected Then
        frmReviComa.txtCantidadComanda.Text = Me.lvwOrdenes.SelectedItem
        frmReviComa.cmdTraer.Value = True
        Unload Me
    End If
End Sub
