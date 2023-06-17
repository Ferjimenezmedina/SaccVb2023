VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form AltaSucu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALTAS SUCURSAL"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7440
      TabIndex        =   17
      Top             =   1560
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "Sucursales.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Sucursales.frx":030A
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label7 
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
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7440
      TabIndex        =   15
      Top             =   2760
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Sucursales.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "Sucursales.frx":1FD6
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label26 
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
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Sucursales"
      TabPicture(0)   =   "Sucursales.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1(4)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         DataField       =   "TELEFONO"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   5
         Left            =   4560
         MaxLength       =   20
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         DataField       =   "ESTADO"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   4
         Left            =   3600
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         DataField       =   "COLONIA"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   3
         Left            =   240
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2160
         Width           =   6735
      End
      Begin VB.TextBox Text1 
         DataField       =   "CALLE"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   2
         Left            =   240
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1560
         Width           =   6735
      End
      Begin VB.TextBox Text1 
         DataField       =   "NOMBRE"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   1
         Left            =   240
         MaxLength       =   25
         TabIndex        =   0
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   3360
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "* % de Impuesto"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "* Nombre"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Colonia"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   3600
         TabIndex        =   11
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "* Telefono"
         Height          =   195
         Left            =   4560
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "* Distin. Factura"
         Height          =   195
         Left            =   3360
         TabIndex        =   9
         Top             =   720
         Width           =   1125
      End
   End
End
Attribute VB_Name = "AltaSucu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image8_Click()
On Error GoTo ManejaError
If Text1(1).Text <> "" And Text1(6).Text <> "" And Text1(5).Text <> "" And Text3.Text <> "" Then
    Dim sqlComanda As String
    sqlComanda = "INSERT INTO SUCURSALES (NOMBRE, CALLE, COLONIA, CIUDAD, ESTADO, TELEFONO, ELIMINADO, DISTINTIVO, IVA) VALUES ('" & Text1(1).Text & "', '" & Text1(2).Text & "', '" & Text1(3).Text & "', '" & Text2.Text & "', '" & Text1(4).Text & "', '" & Text1(5).Text & "', '0', '" & Text1(6).Text & "', " & Replace(Text3.Text, ",", "") & " );"
    cnn.Execute (sqlComanda)
    MsgBox sqlComanda
    sqlComanda = "INSERT INTO FOLIOSUC (SUCURSAL, FOLIO) VALUES ('" & Text1(1).Text & "', 0);"
    cnn.Execute (sqlComanda)
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    Text1(5).Text = ""
    Text1(6).Text = ""
    Text2.Text = ""
    Text3.Text = ""
Else
    MsgBox "FALTA INFORMACIÓN NECESARIA PARA EL REGISTRO!", vbInformation, "SACC"
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
Private Sub Text1_GotFocus(Index As Integer)
On Error GoTo ManejaError
    Text1(Index).BackColor = &HFFE1E1
    Text1(Index).SetFocus
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ManejaError
    Dim Valido As String
    If Index = 5 Then
        Valido = "1234567890-"
    ElseIf Index = 6 Then
        Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ"
    Else
        Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ _ç-,#~<>?¿!¡$@()/&%@!?*+"
    End If
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
Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = &H80000005
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ _ç-,#~<>?¿!¡$@()/&%*+"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
