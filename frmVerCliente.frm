VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmVerCliente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del Cliente"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6600
      Top             =   1680
   End
   Begin VB.TextBox Te1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "ID_CLIENTE"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6120
      TabIndex        =   24
      Top             =   2640
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmVerCliente.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmVerCliente.frx":030A
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
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6588
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Datos Personales"
      TabPicture(0)   =   "frmVerCliente.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label23"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label15"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label18"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label21"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1(6)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(8)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(12)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(15)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(18)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Dirección"
      TabPicture(1)   =   "frmVerCliente.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(0)"
      Tab(1).Control(1)=   "Text1(7)"
      Tab(1).Control(2)=   "Text1(9)"
      Tab(1).Control(3)=   "Text1(10)"
      Tab(1).Control(4)=   "Text1(11)"
      Tab(1).Control(5)=   "Text1(13)"
      Tab(1).Control(6)=   "Text1(14)"
      Tab(1).Control(7)=   "Text1(17)"
      Tab(1).Control(8)=   "Text1(19)"
      Tab(1).Control(9)=   "Text1(20)"
      Tab(1).Control(10)=   "Label8"
      Tab(1).Control(11)=   "Label11"
      Tab(1).Control(12)=   "Label29"
      Tab(1).Control(13)=   "Label30"
      Tab(1).Control(14)=   "Label31"
      Tab(1).Control(15)=   "Label32"
      Tab(1).Control(16)=   "Label33"
      Tab(1).Control(17)=   "Label34"
      Tab(1).Control(18)=   "Label35"
      Tab(1).Control(19)=   "Label36"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "Credito"
      TabPicture(2)   =   "frmVerCliente.frx":2424
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label14"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label24"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Comentarios"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text1(16)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text1(23)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Text1(21)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Text1(22)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Text1(24)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   24
         Left            =   -74760
         MaxLength       =   8
         TabIndex        =   61
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   22
         Left            =   -74760
         MaxLength       =   8
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   -74760
         MaxLength       =   8
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -74880
         MaxLength       =   9
         TabIndex        =   9
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Height          =   2085
         Index           =   23
         Left            =   -72840
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   -74760
         MaxLength       =   8
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   -72840
         MaxLength       =   20
         TabIndex        =   15
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   -74880
         MaxLength       =   30
         TabIndex        =   14
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   -70440
         MaxLength       =   9
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   -74880
         MaxLength       =   100
         TabIndex        =   11
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   -70320
         MaxLength       =   9
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   -71520
         MaxLength       =   9
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   -74880
         MaxLength       =   50
         TabIndex        =   18
         Top             =   3120
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   -74880
         MaxLength       =   100
         TabIndex        =   17
         Top             =   2520
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   -71040
         MaxLength       =   20
         TabIndex        =   16
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   18
         Left            =   4080
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   0
         Top             =   840
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   120
         MaxLength       =   20
         TabIndex        =   2
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   3
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2760
         MaxLength       =   20
         TabIndex        =   7
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   6
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   120
         MaxLength       =   30
         TabIndex        =   5
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   4
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   100
         TabIndex        =   1
         Top             =   1560
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   6
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   27
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   -74760
         TabIndex        =   62
         Top             =   2760
         Width           =   465
      End
      Begin VB.Label Comentarios 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios"
         Height          =   195
         Left            =   -72840
         TabIndex        =   59
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Limite de credito"
         Height          =   195
         Left            =   -74760
         TabIndex        =   58
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descuento"
         Height          =   195
         Left            =   -74760
         TabIndex        =   57
         Top             =   2160
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Dias Crédito"
         Height          =   195
         Left            =   -74760
         TabIndex        =   56
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pais"
         Height          =   195
         Left            =   -71040
         TabIndex        =   55
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "* Dirección"
         Height          =   195
         Left            =   -74880
         TabIndex        =   54
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
         Height          =   195
         Left            =   -74640
         TabIndex        =   53
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Colonia"
         Height          =   195
         Left            =   -74640
         TabIndex        =   52
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Numero Exterior"
         Height          =   195
         Left            =   -70320
         TabIndex        =   51
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Numero Interior"
         Height          =   195
         Left            =   -68760
         TabIndex        =   50
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Postal"
         Height          =   195
         Left            =   -69240
         TabIndex        =   49
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Dirección de Correo Electronico"
         Height          =   195
         Left            =   -71160
         TabIndex        =   48
         Top             =   2640
         Width           =   2250
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   -71880
         TabIndex        =   47
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Direccion Envio"
         Height          =   195
         Left            =   -74640
         TabIndex        =   46
         Top             =   2640
         Width           =   1125
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña de Web"
         Height          =   195
         Left            =   4080
         TabIndex        =   45
         Top             =   2760
         Width           =   1425
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "CURP"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
         Height          =   195
         Left            =   1920
         TabIndex        =   43
         Top             =   2040
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
         Height          =   195
         Left            =   2760
         TabIndex        =   42
         Top             =   2760
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tel. Trabajo"
         Height          =   195
         Left            =   1440
         TabIndex        =   41
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "* Tel. Casa"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   2760
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "* R.F.C"
         Height          =   195
         Left            =   4080
         TabIndex        =   39
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Comercial"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "* Nombre "
         Height          =   195
         Left            =   1320
         TabIndex        =   37
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Clave Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   930
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "* Ciudad"
         Height          =   195
         Left            =   -74880
         TabIndex        =   35
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "* Colonia"
         Height          =   195
         Left            =   -74880
         TabIndex        =   34
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Num Ext"
         Height          =   195
         Left            =   -70320
         TabIndex        =   33
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Num Int"
         Height          =   195
         Left            =   -71520
         TabIndex        =   32
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "* C.P."
         Height          =   195
         Left            =   -70440
         TabIndex        =   31
         Top             =   480
         Width           =   405
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         Height          =   195
         Left            =   -74880
         TabIndex        =   30
         Top             =   2880
         Width           =   435
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   -72840
         TabIndex        =   29
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Direccion Envio"
         Height          =   195
         Left            =   -74880
         TabIndex        =   28
         Top             =   2280
         Width           =   1125
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clave :"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   6120
      TabIndex        =   60
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmVerCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
           "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Timer1_Timer()
    'Te1.Text = Item
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT * FROM CLIENTE WHERE ID_CLIENTE = " & Te1.Text
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Text1(6).Text = Te1.Text
        If Not IsNull(tRs.Fields("NOMBRE")) Then
            Text1(15).Text = tRs.Fields("NOMBRE")
        Else
            Text1(15).Text = ""
        End If
        If Not IsNull(tRs.Fields("NOMBRE_COMERCIAL")) Then
            Text1(1).Text = tRs.Fields("NOMBRE_COMERCIAL")
        Else
            Text1(1).Text = ""
        End If
        If Not IsNull(tRs.Fields("CURP")) Then
            Text1(12).Text = tRs.Fields("CURP")
        Else
            Text1(12).Text = ""
        End If
        If Not IsNull(tRs.Fields("CONTACTO")) Then
            Text1(8).Text = tRs.Fields("CONTACTO")
        Else
            Text1(8).Text = ""
        End If
        If Not IsNull(tRs.Fields("RFC")) Then
            Text1(2).Text = tRs.Fields("RFC")
        Else
            Text1(2).Text = ""
        End If
        If Not IsNull(tRs.Fields("TELEFONO_CASA")) Then
            Text1(3).Text = tRs.Fields("TELEFONO_CASA")
        Else
            Text1(3).Text = ""
        End If
        If Not IsNull(tRs.Fields("TELEFONO_TRABAJO")) Then
            Text1(4).Text = tRs.Fields("TELEFONO_TRABAJO")
        Else
            Text1(4).Text = ""
        End If
        If Not IsNull(tRs.Fields("FAX")) Then
            Text1(5).Text = tRs.Fields("FAX")
        Else
            Text1(5).Text = ""
        End If
        If Not IsNull(tRs.Fields("WEB_PASSWORD")) Then
            Text1(18).Text = tRs.Fields("WEB_PASSWORD")
        Else
            Text1(18).Text = ""
        End If
        If Not IsNull(tRs.Fields("COLONIA")) Then
            Text1(0).Text = tRs.Fields("COLONIA")
        Else
            Text1(0).Text = ""
        End If
        If Not IsNull(tRs.Fields("CP")) Then
            Text1(10).Text = tRs.Fields("CP")
        Else
            Text1(10).Text = ""
        End If
        If Not IsNull(tRs.Fields("DIRECCION")) Then
            Text1(11).Text = tRs.Fields("DIRECCION")
        Else
            Text1(11).Text = ""
        End If
        If Not IsNull(tRs.Fields("NUMERO_EXTERIOR")) Then
            Text1(13).Text = tRs.Fields("NUMERO_EXTERIOR")
        Else
            Text1(13).Text = ""
        End If
        If Not IsNull(tRs.Fields("NUMERO_INTERIOR")) Then
            Text1(14).Text = tRs.Fields("NUMERO_INTERIOR")
        Else
            Text1(14).Text = ""
        End If
        If Not IsNull(tRs.Fields("CIUDAD")) Then
            Text1(9).Text = tRs.Fields("CIUDAD")
        Else
            Text1(9).Text = ""
        End If
        If Not IsNull(tRs.Fields("ESTADO")) Then
            Text1(7).Text = tRs.Fields("ESTADO")
        Else
            Text1(7).Text = ""
        End If
        If Not IsNull(tRs.Fields("PAIS")) Then
            Text1(20).Text = tRs.Fields("PAIS")
        Else
            Text1(20).Text = ""
        End If
        If Not IsNull(tRs.Fields("DIRECCION_ENVIO")) Then
            Text1(19).Text = tRs.Fields("DIRECCION_ENVIO")
        Else
            Text1(19).Text = ""
        End If
        If Not IsNull(tRs.Fields("EMAIL")) Then
            Text1(17).Text = tRs.Fields("EMAIL")
        Else
            Text1(17).Text = ""
        End If
        If Not IsNull(tRs.Fields("LIMITE_CREDITO")) Then
            Text1(16).Text = tRs.Fields("LIMITE_CREDITO")
        Else
            Text1(16).Text = ""
        End If
        If Not IsNull(tRs.Fields("DIAS_CREDITO")) Then
            Text1(21).Text = tRs.Fields("DIAS_CREDITO")
        Else
            Text1(21).Text = ""
        End If
        If Not IsNull(tRs.Fields("DESCUENTO")) Then
            Text1(22).Text = tRs.Fields("DESCUENTO")
        Else
            Text1(22).Text = ""
        End If
        If Not IsNull(tRs.Fields("COMENTARIOS")) Then
            Text1(23).Text = tRs.Fields("COMENTARIOS")
        Else
            Text1(23).Text = ""
        End If
        If Not IsNull(tRs.Fields("NUM_CUENTA_PAGO_CLIENTE")) Then
            Text1(24).Text = tRs.Fields("NUM_CUENTA_PAGO_CLIENTE")
        Else
            Text1(24).Text = ""
        End If
    End If
    Timer1.Enabled = False
End Sub
