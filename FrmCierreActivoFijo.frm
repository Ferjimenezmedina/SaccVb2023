VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCierreActivoFijo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cerrar como..."
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmCierreActivoFijo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.Frame Frame1 
         Caption         =   "Detalle"
         Height          =   2295
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   6855
         Begin VB.TextBox Text3 
            Height          =   1215
            Left            =   960
            MaxLength       =   500
            TabIndex        =   13
            Top             =   960
            Width           =   5775
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   960
            MaxLength       =   50
            TabIndex        =   12
            Top             =   480
            Width           =   5775
         End
         Begin VB.Label Label3 
            Caption         =   "Notas"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "No. Serie"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frame2"
         Height          =   2295
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   6735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Funcional"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Asistencia"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Scrap"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Contador del equipo"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   2
      Top             =   2640
      Width           =   975
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCierreActivoFijo.frx":001C
         MousePointer    =   99  'Custom
         Picture         =   "FrmCierreActivoFijo.frx":0326
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   0
      Top             =   1440
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmCierreActivoFijo.frx":2408
         MousePointer    =   99  'Custom
         Picture         =   "FrmCierreActivoFijo.frx":2712
         Top             =   240
         Width           =   675
      End
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmCierreActivoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Public IdPrestamo As String
Public IdProducto As String
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
           "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
End Sub
Private Sub Image8_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If Option1.Value = False Then
        If Option2.Value Then
            sBuscar = "INSERT INTO ASISTENCIA_SERVICIO (ID_PRESTAMO, ID_PRODUCTO, NO_SERIE, NOTA, ESTADO_ASISTENCIA) VALUES ('" & IdPrestamo & "', '" & IdProducto & "', '" & Text2.Text & "', '" & Text3.Text & "', 'I')"
        Else
            sBuscar = "INSERT INTO SCRAP_PRESTAMOS (ID_PRESTAMO, ID_PRODUCTO, NO_SERIE, NOTA) VALUES ('" & IdPrestamo & "', '" & IdProducto & "', '" & Text2.Text & "', '" & Text3.Text & "')"
        End If
    Else
        sBuscar = "INSERT INTO DEVOLUCIONES_RESTAMOS (ID_PRESTAMO, ID_PRODUCTO, NO_SERIE, NOTA) VALUES ('" & IdPrestamo & "', '" & IdProducto & "', '" & Text2.Text & "', '" & Text3.Text & "')"
    End If
    cnn.Execute (sBuscar)
    Unload Me
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Option1_Click()
    If Option1.Value = True Then
        Frame1.Visible = False
    Else
        Frame1.Visible = True
    End If
End Sub
Private Sub Option2_Click()
    If Option2.Value = True Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False
    End If
End Sub
Private Sub Option3_Click()
    If Option3.Value = True Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
