VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Agente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AGENTE"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4800
      Top             =   1080
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
      RecordSource    =   "AGENTE"
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
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton btnBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton btnSalir 
      BackColor       =   &H80000009&
      Caption         =   "Salir"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Numero"
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por Nombre"
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   3120
      TabIndex        =   7
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar "
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Direccion"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Telefono"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Comision"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Clave"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "Agente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalir_Click()
    Unload Me
End Sub
Private Sub cmdAdd_Click()
    If Text1(1).Text <> "" And Text1(2).Text <> "" And Text1(3).Text <> "" Then
        Adodc1.Recordset.Update
        If Text2.Enabled = True Then
            Me.Text2.Enabled = False
        End If
    Else
        MsgBox "Falta informacion necesaria para registrar"
    End If
End Sub
Private Sub Command1_Click()
    Text1(0).Enabled = True
    Text1(1).Enabled = True
    Text1(2).Enabled = True
    Text1(3).Enabled = True
    Text1(4).Enabled = True
    Me.cmdAdd.Enabled = True
    Adodc1.Recordset.AddNew
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If Index = 4 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 25 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub Text2_Change()
    If Text2.Text = "" Then
        btnBuscar.Enabled = False
    Else
        btnBuscar.Enabled = True
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Me.Text1(4).Visible = True
    If KeyAscii = vbKeyReturn Then
        On Error Resume Next
        KeyAscii = 0
        Buscar
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
        sADOBuscar = "ID_AGENTE = " & nReg
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
Private Sub Form_Load()
    Text2 = ""
    Option2.Value = True
    Dim I As Long
    For I = 0 To 4
        Set Text1(I).DataSource = Adodc1
    Next
    Text1(0).DataField = "ID_AGENTE"
    Text1(1).DataField = "NOMBRE"
    Text1(2).DataField = "DIRECCION"
    Text1(3).DataField = "TELEFONO"
    Text1(4).DataField = "COMISION"
    Text1(0).Enabled = False
    Text1(1).Enabled = False
    Text1(2).Enabled = False
    Text1(3).Enabled = False
    Text1(4).Enabled = False
    Me.cmdAdd.Enabled = False
End Sub
