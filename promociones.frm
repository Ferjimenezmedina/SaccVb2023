VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Promos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALTA PROMOCIONES"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      DataField       =   "ACTIVA"
      DataSource      =   "Adodc1"
      Height          =   195
      Index           =   5
      Left            =   1680
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Activa"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   "TERMINA"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   4
      Left            =   3120
      TabIndex        =   9
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID_PROMOCION"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "EMPIEZA"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "DESCUENTO"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      DataField       =   "DESCRIPCION"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   5295
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   3480
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por Nombre"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Numero"
      Height          =   195
      Left            =   4200
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton btnSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton btnBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   3480
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "PROMOCION"
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
   Begin VB.Label Label7 
      Caption         =   "Clave"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Termina"
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Empieza"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Descuento"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar "
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Promos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Text1_Change(Index As Integer)
    If Text1(5).Text = "1" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If Index = 3 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 25 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub btnBuscar_Click()
    Buscar
End Sub
Private Sub Buscar(Optional ByVal Siguiente As Boolean = False)
    Dim nReg As Long
    Dim vBookmark As Variant
    Dim sADOBuscar As String
    On Error Resume Next
    If Option1.Value Then
        nReg = Val(Text2)
        sADOBuscar = "ID_PROMOCION = " & nReg
    End If
    If Option2.Value Then
        sADOBuscar = "DESCRIPCION Like '" & Text2.Text & "'"
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
Private Sub Form_Load()
    Text2 = ""
    Option2.Value = True
    Dim i As Long
    For i = 0 To 5
        Set Text1(i).DataSource = Adodc1
    Next
    Text1(0).DataField = "ID_PROMOCION"
    Text1(1).DataField = "DESCRIPCION"
    Text1(2).DataField = "DESCUENTO"
    Text1(3).DataField = "EMPIEZA"
    Text1(4).DataField = "TERMINA"
    Text1(5).DataField = "ACTIVA"
End Sub
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Text1(5).Text = "1"
    Else
        Text1(5).Text = "0"
    End If
End Sub
