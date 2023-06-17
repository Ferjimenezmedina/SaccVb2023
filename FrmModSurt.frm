VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmModSurt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Surtir Producto de Pedido"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6120
      TabIndex        =   12
      Top             =   1800
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
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmModSurt.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmModSurt.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmModSurt.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   2400
         Width           =   1335
      End
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
         Left            =   4200
         Picture         =   "FrmModSurt.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Clave del Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Numero de Orden :"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad Pediente :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label8 
         Caption         =   "Cantidad en Existencia :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   1920
         Width           =   3375
      End
   End
End
Attribute VB_Name = "FrmModSurt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim Pend As Double
    Dim NuevaExis As Double
    Dim sqlComanda As String
    Dim tRs As ADODB.Recordset
    If CDbl(Text1.Text) > CDbl(Label9.Caption) Then
        MsgBox "EL INVENTARIO NO CUENTA CON CANTIDAD SUFICIENTE!", vbInformation, "SACC"
    Else
        If CDbl(Text1.Text) > CDbl(Label7.Caption) Then
            MsgBox "NO PUEDE SURTIR CANTIDAD MAYOR A LA PENDIENTE!", vbInformation, "SACC"
        Else
            NuevaExis = CDbl(Text1.Text)
            Pend = CDbl(Label7.Caption) - CDbl(Text1.Text)
            sqlComanda = "UPDATE PED_CLIEN_DETALLE SET CANTIDAD_PENDIENTE = " & Pend & " WHERE ID_PRODUCTO = '" & Label6.Caption & "' AND NO_PEDIDO = " & Label5.Caption
            Set tRs = cnn.Execute(sqlComanda)
            sqlComanda = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Label6.Caption & "' AND SUCURSAL = '" & frmShowPediC.Combo1.Text & "'"
            Set tRs = cnn.Execute(sqlComanda)
            If Not (tRs.EOF And tRs.BOF) Then
                sqlComanda = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs.Fields("CANTIDAD")) - CDbl(NuevaExis) & " WHERE ID_PRODUCTO = '" & Label6.Caption & "' AND SUCURSAL = '" & frmShowPediC.Combo1.Text & "'"
                cnn.Execute (sqlComanda)
            Else
                MsgBox "DIFERENCIA EN EXISTENCIA! COMUNIQUELO AL ADMINISTRADOR DEL SISTEMA!", vbInformation, "SACC"
                Exit Sub
            End If
            Command1.Enabled = False
            'CODIGO DE RECORRIDO DE LA TABLA DE PEDIDOS PARA VER SI ESTE PEDIDO
            'YA QUEDO COMPLETO O NO Y CERRARLO EN CASO DE QUE TODOS LOS PENDIENTES
            'SEAN CEROS
            sqlComanda = "SELECT CANTIDAD_PENDIENTE FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO = " & Label5.Caption & " AND CANTIDAD_PENDIENTE > 0"
            Set tRs = cnn.Execute(sqlComanda)
            If (tRs.BOF And tRs.EOF) Then
                sqlComanda = "UPDATE PED_CLIEN SET ESTADO = 'C' WHERE NO_PEDIDO = " & Label5.Caption
                Set tRs = cnn.Execute(sqlComanda)
            End If
            Unload Me
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
    Label5.Caption = frmShowPediC.Text1(2)
    Label6.Caption = frmShowPediC.Text1(0)
    Label7.Caption = frmShowPediC.Text1(1)
    Label10.Caption = frmShowPediC.Text1(5)
    Text1.Text = frmShowPediC.Text1(1)
    Command1.Enabled = False
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE SUCURSAL = '" & frmShowPediC.Combo1.Text & "' AND ID_PRODUCTO = '" & Label6.Caption & "'"
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
    Valido = "1234567890."
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
