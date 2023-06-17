VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmDeshacer2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deshacer (Regresar a Existencia)"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6000
      TabIndex        =   16
      Top             =   2640
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmDeshacer2.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmDeshacer2.frx":030A
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
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmDeshacer2.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label8"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label9"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.CommandButton Command1 
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
         Height          =   375
         Left            =   3120
         Picture         =   "FrmDeshacer2.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label Label8 
         Caption         =   "Cantidad en Existencia :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad Surtida :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Numero de Orden :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Clave del Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Cantidad Pendiente :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Cantidad Pedida :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   2760
         Width           =   3375
      End
   End
End
Attribute VB_Name = "FrmDeshacer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim Pend As Double
    Dim NuevaExis As Double
    Dim sqlComanda As String
    Dim tRs As ADODB.Recordset
    If CDbl(Text1.Text) > CDbl(Label7.Caption) Then
        MsgBox "NO ESTA PERMITIDO REGRESAR CANTIDAD MAYOR A LA APARTADA!", vbInformation, "SACC"
    Else
        NuevaExis = CDbl(Label9.Caption) + CDbl(Text1.Text)
        Pend = CDbl(Label12.Caption) - CDbl(Text1.Text)
        sqlComanda = "UPDATE PED_CLIEN_DETALLE SET CANTIDAD_PENDIENTE = " & Pend & " WHERE ID_PRODUCTO = '" & Label6.Caption & "' AND NO_PEDIDO = " & Label5.Caption
        Set tRs = cnn.Execute(sqlComanda)
        sqlComanda = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Label6.Caption & "' AND SUCURSAL = 'BODEGA'"
        Set tRs = cnn.Execute(sqlComanda)
        If Not (tRs.EOF And tRs.BOF) Then
            sqlComanda = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs.Fields("CANTIDAD")) + CDbl(Text1.Text) & " WHERE ID_PRODUCTO = '" & Label6.Caption & "' AND SUCURSAL = 'BODEGA'"
            Set tRs = cnn.Execute(sqlComanda)
        Else
            sqlComanda = "INSERT INTO EXISTENCIAS (CANTIDAD, ID_PRODUCTO, SUCURSAL) VALUES(" & Text1.Text & ", '" & Label6.Caption & "', 'BODEGA');"
            cnn.Execute (sqlComanda)
        End If
        Command1.Enabled = False
        Unload Me
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Label5.Caption = FrmAjusteManual.Text1(2)
    Label6.Caption = FrmAjusteManual.Text1(0)
    Label7.Caption = FrmAjusteManual.Text1(1)
    Label12.Caption = FrmAjusteManual.Text1(3)
    Label13.Caption = FrmAjusteManual.Text1(4)
    Command1.Enabled = False
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE SUCURSAL = 'BODEGA' AND ID_PRODUCTO = '" & Label6.Caption & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF Or tRs.EOF) Then
        Label9.Caption = tRs.Fields("CANTIDAD") & ""
    Else
        Label9.Caption = "0"
    End If
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
Private Sub Text1_Change()
On Error GoTo ManejaError
    If Text1.Text = "" Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_GotFocus()
On Error GoTo ManejaError
    Text1.BackColor = &HFFE1E1
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_LostFocus()
      Text1.BackColor = &H80000005
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
