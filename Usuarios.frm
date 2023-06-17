VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Usuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALTA USUARIOS"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   7455
      Begin VB.TextBox Text1 
         DataField       =   "ID_USUARIO"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   360
         MaxLength       =   4
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         DataField       =   "NOMBRE"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   1
         Left            =   360
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1200
         Width           =   6735
      End
      Begin VB.TextBox Text1 
         DataField       =   "APELLIDO"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   2
         Left            =   360
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1800
         Width           =   6735
      End
      Begin VB.TextBox Text1 
         DataField       =   "PUESTO"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   3
         Left            =   360
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2400
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         DataField       =   "PASSWORD"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   4
         Left            =   360
         MaxLength       =   30
         TabIndex        =   7
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clave"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   2760
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Puesto"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   2160
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Apellidos"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   960
         Width           =   555
      End
   End
   Begin VB.CommandButton btnSalir 
      BackColor       =   &H80000009&
      Caption         =   "Salir"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   4560
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por Nombre"
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Clave"
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   330
      Width           =   4815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5760
      Top             =   4560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "USUARIOS"
      Caption         =   "Siguiente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Buscar"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Buscar(Optional ByVal Siguiente As Boolean = False)
    Dim vBookmark As Variant
    Dim sADOBuscar As String
    On Error Resume Next
    If Option1.Value Then
        sADOBuscar = "ID_USUARIO Like '" & Text2.Text & "'"
    End If
    If Option2.Value Then
        sADOBuscar = "NOMBRE Like '" & Text2.Text & "'"
    End If
    vBookmark = Adodc1.Recordset.Bookmark
    If Siguiente = False Then
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset.Find sADOBuscar
    Else
        Adodc1.Recordset.Find sADOBuscar, 1
    End If
    If Err.Number Or Adodc1.Recordset.BOF Or Adodc1.Recordset.EOF Then
        Err.Clear
        MsgBox "No existe el dato buscado o ya no hay más datos que mostrar."
        Adodc1.Recordset.Bookmark = vBookmark
    End If
End Sub
Private Sub btnSalir_Click()
    Unload Me
End Sub
Private Sub cmdAdd_Click()
    If Text1(1).Text <> "" And Text1(2).Text <> "" And Text1(3).Text <> "" Then
        cmdAdd.Caption = "Guardar/Nuevo"
        Adodc1.Recordset.AddNew
        If Text2.Enabled = True Then
            Me.Text2.Enabled = False
        End If
    End If
End Sub
Private Sub Form_Load()
    Text2 = ""
    Option2.Value = True
    Dim I As Long
    For I = 0 To 4
        Set Text1(I).DataSource = Adodc1
    Next
    Text1(0).DataField = "ID_USUARIO"
    Text1(1).DataField = "NOMBRE"
    Text1(2).DataField = "APELLIDO"
    Text1(3).DataField = "PUESTO"
    Text1(4).DataField = "PASSWORD"
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    If Index = 0 Or Index = 7 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Buscar
    End If
End Sub
