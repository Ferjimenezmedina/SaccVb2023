VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEmpresa 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos de la Empresa"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   28
      Top             =   4800
      Width           =   6855
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Factura electrónica"
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
      Left            =   4200
      TabIndex        =   27
      Top             =   1320
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1350
      Left            =   120
      Picture         =   "frmEmpresa.frx":0000
      ScaleHeight     =   1350
      ScaleWidth      =   3915
      TabIndex        =   26
      Top             =   120
      Width           =   3915
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
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
      Left            =   4200
      Picture         =   "frmEmpresa.frx":1129
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   7080
      TabIndex        =   23
      Top             =   2640
      Width           =   975
      Begin VB.Image Image1 
         Height          =   705
         Left            =   120
         MouseIcon       =   "frmEmpresa.frx":3AFB
         MousePointer    =   99  'Custom
         Picture         =   "frmEmpresa.frx":3E05
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label3 
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
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   7080
      TabIndex        =   21
      Top             =   3840
      Width           =   975
      Begin VB.Label Label2 
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
      Begin VB.Image Image6 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmEmpresa.frx":58B7
         MousePointer    =   99  'Custom
         Picture         =   "frmEmpresa.frx":5BC1
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   19
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7800
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   4560
      TabIndex        =   8
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   3960
      TabIndex        =   7
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   2400
      TabIndex        =   6
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   3600
      TabIndex        =   3
      Top             =   4200
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   6855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "* Quien Autoriza los Pagos?"
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
      Index           =   10
      Left            =   120
      TabIndex        =   29
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "* C.P."
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
      Index           =   9
      Left            =   2040
      TabIndex        =   20
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pais"
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
      Index           =   8
      Left            =   4560
      TabIndex        =   17
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "* RFC"
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
      Index           =   7
      Left            =   3960
      TabIndex        =   16
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "* Estado"
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
      Index           =   6
      Left            =   2400
      TabIndex        =   15
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "* Ciudad"
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
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "* Colonia"
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
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
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
      Index           =   3
      Left            =   3600
      TabIndex        =   12
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "* Telefono"
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
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "* Direccion"
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
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "* Nombre de la empreza"
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
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   4335
   End
End
Attribute VB_Name = "frmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Command9_Click()
    Dim Ruta As String
    CommonDialog1.DialogTitle = "Abrir"
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "Imagen JPG|*.jpg|Imagen BMP|*.bmp|Imagen PNG|*.png"
    CommonDialog1.ShowOpen
    Ruta = CommonDialog1.FileName
    If Ruta <> "" Then
        Picture1.Picture = LoadPicture(Ruta)
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    sBuscar = "SELECT * FROM EMPRESA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Text2.Text = "S"
        Text1(0).Text = tRs.Fields("NOMBRE")
        Text1(1).Text = tRs.Fields("DIRECCION")
        Text1(2).Text = tRs.Fields("TELEFONO")
        Text1(3).Text = tRs.Fields("FAX")
        Text1(4).Text = tRs.Fields("COLONIA")
        Text1(5).Text = tRs.Fields("CD")
        Text1(6).Text = tRs.Fields("ESTADO")
        Text1(8).Text = tRs.Fields("PAIS")
        Text1(7).Text = tRs.Fields("RFC")
        Text1(9).Text = tRs.Fields("CP")
        'If Not IsNull(tRs.Fields("IMAGEN")) Or tRs.Fields("IMAGEN") = "" Then
            Picture1.Picture = LoadPicture(App.Path & "/REPORTES/LOGO.jpg")
        'End If
        If tRs.Fields("FACTURA_ELECTRONICA") = "S" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    End If
End Sub
Private Sub Image1_Click()
    Dim sBuscar As String
    Dim Ruta As String
    Dim Fac_elec As String
    Ruta = App.Path & "\REPORTES\" & Mid(Text1(0).Text, 1, 3) & ".jpg"
    SavePicture Picture1.Image, Ruta
    If Check1.Value Then
        Fac_elec = "S"
    Else
        Fac_elec = "N"
    End If
    If Text2.Text <> "" Then
        sBuscar = "UPDATE EMPRESA SET NOMBRE = '" & Text1(0).Text & "', DIRECCION = '" & Text1(1).Text & "', TELEFONO = '" & Text1(2).Text & "', FAX = '" & Text1(3).Text & "', COLONIA = '" & Text1(4).Text & "', CD = '" & Text1(5).Text & "', ESTADO = '" & Text1(6).Text & "', PAIS = '" & Text1(8).Text & "', RFC = '" & Text1(7).Text & "', CP = '" & Text1(9).Text & "', FACTURA_ELECTRONICA = '" & Fac_elec & "', REPRESENTANTE = '" & Text1(10).Text & "'"
    Else
        sBuscar = "INSERT INTO EMPRESA(NOMBRE, DIRECCION, TELEFONO, FAX, COLONIA, CD, ESTADO, PAIS, RFC, CP, FACTURA_ELECTRONICA, REPRESENTANTE) VALUES('" & Text1(0).Text & "', '" & Text1(1).Text & "', '" & Text1(2).Text & "', '" & Text1(3).Text & "', '" & Text1(4).Text & "', '" & Text1(5).Text & "', '" & Text1(6).Text & "', '" & Text1(8).Text & "', '" & Text1(7).Text & "', '" & Text1(9).Text & "', '" & Fac_elec & "', '" & Text1(10).Text & "');"
    End If
    cnn.Execute (sBuscar)
    Text2.Text = "S"
    MsgBox "INFORMACIÓN GUARDADA", vbInformation, "SACC"
End Sub
Private Sub Image6_Click()
    Unload Me
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).BackColor = &HFFE1E1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Valido As String
    Select Case Index
        Case Is = 0: Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -.,"
        Case Is = 1: Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -.,1234567890"
        Case Is = 2: Valido = "1234567890- "
        Case Is = 3: Valido = "1234567890- "
        Case Is = 4: Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -.,"
        Case Is = 5: Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -.,"
        Case Is = 6: Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -.,"
        Case Is = 7: Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -.,1234567890"
        Case Is = 8: Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -.,"
        Case Is = 9: Valido = "1234567890"
        Case Is = 10: Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -.,1234567890"
    End Select
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = &H80000005
End Sub
