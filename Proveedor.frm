VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Proveedor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALTA PROVEEDORES"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   47
      Top             =   3120
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Proveedor.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Proveedor.frx":030A
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
         TabIndex        =   48
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   43
      Top             =   1920
      Width           =   975
      Begin VB.Label Label21 
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
         TabIndex        =   44
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "Proveedor.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "Proveedor.frx":26F6
         Top             =   240
         Width           =   675
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "Proveedor.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label12"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label13"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label14"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label11"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(7)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(8)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(11)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(10)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(9)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(19)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Bancos"
      TabPicture(1)   =   "Proveedor.frx":40D4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(12)"
      Tab(1).Control(1)=   "Text1(13)"
      Tab(1).Control(2)=   "Text1(14)"
      Tab(1).Control(3)=   "Text1(15)"
      Tab(1).Control(4)=   "Text1(16)"
      Tab(1).Control(5)=   "Text1(17)"
      Tab(1).Control(6)=   "Label15"
      Tab(1).Control(7)=   "Label16"
      Tab(1).Control(8)=   "Label17"
      Tab(1).Control(9)=   "Label18"
      Tab(1).Control(10)=   "Label19"
      Tab(1).Control(11)=   "Label20"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Notas"
      TabPicture(2)   =   "Proveedor.frx":40F0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "Text1(18)"
      Tab(2).Control(2)=   "Label7"
      Tab(2).ControlCount=   3
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   120
         MaxLength       =   100
         TabIndex        =   11
         Top             =   3720
         Width           =   6735
      End
      Begin VB.Frame Frame1 
         Caption         =   "Almacen"
         Height          =   3135
         Left            =   -70200
         TabIndex        =   46
         Top             =   720
         Width           =   2055
         Begin VB.CheckBox Check3 
            Caption         =   "Almacen 3"
            Height          =   255
            Left            =   480
            TabIndex        =   21
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Almacen 2"
            Height          =   255
            Left            =   480
            TabIndex        =   20
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Almacen 1"
            Height          =   255
            Left            =   480
            TabIndex        =   19
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.TextBox Text1 
         Height          =   2325
         Index           =   18
         Left            =   -74760
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   12
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   -72480
         MaxLength       =   100
         TabIndex        =   13
         Top             =   1560
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   14
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   -72240
         MaxLength       =   100
         TabIndex        =   15
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   17
         Top             =   2760
         Width           =   6615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   -70080
         MaxLength       =   100
         TabIndex        =   16
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   120
         MaxLength       =   20
         TabIndex        =   8
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   9
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   4680
         MaxLength       =   20
         TabIndex        =   10
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   4560
         TabIndex        =   7
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   120
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   120
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   23
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   100
         TabIndex        =   0
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   120
         MaxLength       =   100
         TabIndex        =   1
         Top             =   1320
         Width           =   6735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "* E-mail"
         Height          =   195
         Left            =   240
         TabIndex        =   49
         Top             =   3480
         Width           =   525
      End
      Begin VB.Label Label7 
         Caption         =   "Notas"
         Height          =   255
         Left            =   -74640
         TabIndex        =   42
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   -74640
         TabIndex        =   41
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Left            =   -72360
         TabIndex        =   40
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
         Height          =   195
         Left            =   -74640
         TabIndex        =   39
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Routing"
         Height          =   195
         Left            =   -72120
         TabIndex        =   38
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
         Height          =   195
         Left            =   -74640
         TabIndex        =   37
         Top             =   2520
         Width           =   510
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Clave Swift"
         Height          =   195
         Left            =   -70080
         TabIndex        =   36
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
         Height          =   195
         Left            =   4800
         TabIndex        =   35
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Telefono 2"
         Height          =   195
         Left            =   2520
         TabIndex        =   34
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "* Telefono 1"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   2880
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "* R.F.C."
         Height          =   195
         Left            =   4680
         TabIndex        =   32
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "* C.P."
         Height          =   195
         Left            =   3120
         TabIndex        =   31
         Top             =   2280
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pais"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   2280
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   4920
         TabIndex        =   29
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "* Ciudad"
         Height          =   195
         Left            =   2760
         TabIndex        =   28
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "* Colonia"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "* Direccion"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "* Nombre"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clave"
         Height          =   195
         Left            =   5160
         TabIndex        =   24
         Top             =   480
         Width           =   405
      End
   End
   Begin VB.Label Label22 
      Caption         =   "Label22"
      Height          =   495
      Left            =   3600
      TabIndex        =   45
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "Proveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Image8_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim ALM1 As String
    Dim ALM2 As String
    Dim ALM3 As String
    If Check1.Value = 1 Then
        ALM1 = "S"
    Else
        ALM1 = "N"
    End If
    If Check2.Value = 1 Then
        ALM2 = "S"
    Else
        ALM2 = "N"
    End If
    If Check3.Value = 1 Then
        ALM3 = "S"
    Else
        ALM3 = "N"
    End If
    If Text1(1).Text <> "" And Text1(2).Text <> "" And Text1(3).Text <> "" And Text1(4).Text <> "" And Text1(5).Text <> "" And Text1(6).Text <> "" And Text1(8).Text <> "" And Text1(9).Text <> "" And Text1(19).Text <> "" Then
        sBuscar = "INSERT INTO PROVEEDOR (NOMBRE, DIRECCION, COLONIA, CIUDAD, ESTADO, PAIS, CP, RFC, TELEFONO1, TELEFONO2, TELEFONO3, TRANS_BANCO, TRANS_DIRECCION, TRANS_CIUDAD, TRANS_ROUTING, TRANS_CUENTA, TRANS_CLAVE_SWIFT, NOTAS, ALMACEN1, ALMACEN2, ALMACEN3, ELIMINADO, EMAIL) VALUES ('" & Text1(1).Text & "' , '" & Text1(2).Text & "' , '" & Text1(3).Text & "' , '" & Text1(4).Text & "' , '" & Text1(5).Text & "' , '" & Text1(6).Text & "', '" & Text1(7).Text & "', '" & Text1(8).Text & "' , '" & Text1(9).Text & "' , '" & Text1(10).Text & "' , '" & Text1(11).Text & "' , '" & Text1(12).Text & "' , '" & Text1(13).Text & "' , '" & Text1(14).Text & "' , '" & Text1(15).Text & "' , '" & Text1(16).Text & "' , '" & Text1(17).Text & "' , '" & Text1(18).Text & "', '" & ALM1 & "', '" & ALM2 & "', '" & ALM3 & "', 'N', '" & Text1(19).Text & "')"
        cnn.Execute (sBuscar)
        Text1(0).Text = ""
        Text1(1).Text = ""
        Text1(2).Text = ""
        Text1(3).Text = ""
        Text1(4).Text = ""
        Text1(5).Text = ""
        Text1(6).Text = ""
        Text1(7).Text = ""
        Text1(8).Text = ""
        Text1(9).Text = ""
        Text1(10).Text = ""
        Text1(11).Text = ""
        Text1(12).Text = ""
        Text1(13).Text = ""
        Text1(14).Text = ""
        Text1(15).Text = ""
        Text1(16).Text = ""
        Text1(17).Text = ""
        Text1(18).Text = ""
        Text1(19).Text = ""
        Check1.Value = 0
        Check2.Value = 0
        Check3.Value = 0
    Else
        MsgBox "FALTA INFORMACION NECESARIA PARA EL REGISTRO!", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
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
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    If Index <> 0 Then
        Text1(Index).BackColor = &HFFE1E1
    End If
    Text1(Index).SetFocus
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Valido As String
    If Index = 7 Or Index = 9 Or Index = 10 Or Index = 11 Then
        Valido = "1234567890-()"
    Else
        If Index = 18 Then
            Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz -()/&%@!?*+"
        Else
            If Index = 19 Then
                Valido = "1234567890.abcdefghijklmnñopqrstuvwxyz@-_"
            Else
                Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+"
            End If
        End If
    End If
    If Index = 18 Or Index = 19 Then
        KeyAscii = Asc(Chr(KeyAscii))
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    If Index <> 0 Then
        Text1(Index).BackColor = &H80000005
    End If
End Sub
