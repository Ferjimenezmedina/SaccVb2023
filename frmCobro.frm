VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCobro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobrar Mercancia"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmCobro.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblEstado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdImprimir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtComents"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtNombre"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   0
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox txtComents 
         Appearance      =   0  'Flat
         Height          =   1365
         Left            =   120
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1320
         Width           =   4815
      End
      Begin VB.CommandButton cmdImprimir 
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
         Height          =   375
         Left            =   2400
         Picture         =   "frmCobro.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
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
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentarios"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblEstado 
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
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   3615
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5280
      TabIndex        =   3
      Top             =   2400
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmCobro.frx":29EE
         MousePointer    =   99  'Custom
         Picture         =   "frmCobro.frx":2CF8
         Top             =   120
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As Recordset
Dim cont As Integer
Dim NoRe As Integer
Dim nID_COBRO As Integer
Private Sub cmdImprimir_Click()
On Error GoTo ManejaError
    If Puede_Imprimir Then
        sqlQuery = "INSERT INTO COBRO_MCIA (NOMBRE, COMENTARIO, FECHA, ID_INVENTARIO) VALUES ('" & Me.TxtNOMBRE.Text & "', '" & Me.txtComents.Text & "', '" & Format(Date, "dd/mm/yyyy") & "', " & frmPerdidas.lblInv.Caption & ")"
        'InputBox "", "", sqlQuery
        cnn.Execute (sqlQuery)
        sqlQuery = "SELECT TOP 1 ID_COBRO FROM COBRO_MCIA ORDER BY ID_COBRO DESC"
        'InputBox "", "", sqlQuery
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            nID_COBRO = .Fields("ID_COBRO")
        End With
        NoRe = frmPerdidas.lvwInventario.ListItems.Count
        For cont = 1 To NoRe
            If frmPerdidas.lvwInventario.ListItems.Item(cont).Checked = True Then
                sqlQuery = "INSERT INTO COBRO_MCIA_DETALLE (ID_COBRO, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_UNITARIO) VALUES (" & nID_COBRO & ", '" & frmPerdidas.lvwInventario.ListItems.Item(cont) & "', '" & frmPerdidas.lvwInventario.ListItems.Item(cont).SubItems(1) & "', " & frmPerdidas.lvwInventario.ListItems.Item(cont).SubItems(6) & ", " & Replace(FormatNumber(frmPerdidas.lvwInventario.ListItems.Item(cont).SubItems(4), 2, vbFalse, vbFalse, vbFalse), ",", ".") & ")"
                InputBox "", "", sqlQuery
                cnn.Execute (sqlQuery)
            End If
        Next cont
        Set crReport = crApplication.OpenReport(App.Path & "\REPORTES\COBRO_MCIA.rpt")
        crReport.ParameterFields.Item(1).ClearCurrentValueAndRange
        crReport.ParameterFields.Item(1).AddCurrentValue nID_COBRO
        crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        crReport.PrintOut
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Puede_Imprimir() As Boolean
On Error GoTo ManejaError
    If Me.TxtNOMBRE.Text = "" Then
        Me.lblEstado.Caption = "Escriba el nombre"
        Me.lblEstado.ForeColor = vbRed
        Me.TxtNOMBRE.SetFocus
        Puede_Imprimir = False
        Exit Function
    End If
    If Me.txtComents.Text = "" Then
        Me.lblEstado.Caption = "Escriba el comentario"
        Me.lblEstado.ForeColor = vbRed
        Me.txtComents.SetFocus
        Puede_Imprimir = False
        Exit Function
    End If
    Puede_Imprimir = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub txtComents_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.cmdImprimir.Value = True
    End If
End Sub
Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.txtComents.SetFocus
    End If
End Sub
