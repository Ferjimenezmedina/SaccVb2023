VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form AltaClien 
   Caption         =   "CLIENTES"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7800
      TabIndex        =   62
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Altas clientes.frx":0000
      DataField       =   "ID_AGENTE"
      DataSource      =   "Adodc2"
      Height          =   315
      Left            =   7800
      TabIndex        =   61
      Top             =   3600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2520
      TabIndex        =   59
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   21233665
      CurrentDate     =   38664
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   7680
      TabIndex        =   58
      Top             =   6000
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6000
      Top             =   1200
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6480
      Top             =   960
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
      RecordSource    =   "CLIENTE"
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
   Begin VB.OptionButton Option2 
      Caption         =   "Por Nombre"
      Height          =   375
      Left            =   7800
      TabIndex        =   57
      Top             =   480
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Clave"
      Height          =   495
      Left            =   7800
      TabIndex        =   56
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   25
      Left            =   7800
      MaxLength       =   40
      TabIndex        =   55
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton btnSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9120
      TabIndex        =   54
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton btnBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   9240
      TabIndex        =   53
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dirección"
      Height          =   3495
      Left            =   120
      TabIndex        =   32
      Top             =   3480
      Width           =   7335
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   120
         MaxLength       =   40
         TabIndex        =   42
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   3000
         MaxLength       =   20
         TabIndex        =   41
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   120
         MaxLength       =   30
         TabIndex        =   40
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   4440
         MaxLength       =   9
         TabIndex        =   39
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   11
         Left            =   120
         MaxLength       =   100
         TabIndex        =   38
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   6000
         MaxLength       =   9
         TabIndex        =   37
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   4440
         MaxLength       =   9
         TabIndex        =   36
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   35
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   120
         MaxLength       =   100
         TabIndex        =   34
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   5520
         MaxLength       =   20
         TabIndex        =   33
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pais"
         Height          =   195
         Left            =   5760
         TabIndex        =   52
         Top             =   1680
         Width           =   300
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Left            =   360
         TabIndex        =   51
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
         Height          =   195
         Left            =   360
         TabIndex        =   50
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Colonia"
         Height          =   195
         Left            =   360
         TabIndex        =   49
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Numero Exterior"
         Height          =   195
         Left            =   6000
         TabIndex        =   48
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Numero Interior"
         Height          =   195
         Left            =   4440
         TabIndex        =   47
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Postal"
         Height          =   195
         Left            =   4680
         TabIndex        =   46
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Dirección de Correo Electronico"
         Height          =   195
         Left            =   4080
         TabIndex        =   45
         Top             =   2400
         Width           =   2250
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   3240
         TabIndex        =   44
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Direccion Envio"
         Height          =   195
         Left            =   360
         TabIndex        =   43
         Top             =   2400
         Width           =   1125
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   24
      Left            =   7800
      MaxLength       =   15
      TabIndex        =   31
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   1485
      Index           =   23
      Left            =   120
      TabIndex        =   30
      Top             =   7320
      Width           =   10455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   7800
      MaxLength       =   10
      TabIndex        =   26
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   21
      Left            =   2520
      TabIndex        =   25
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   8760
      MaxLength       =   15
      TabIndex        =   24
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   7800
      MaxLength       =   8
      TabIndex        =   22
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   120
      MaxLength       =   100
      TabIndex        =   21
      Top             =   1680
      Width           =   7695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   8160
      MaxLength       =   20
      TabIndex        =   18
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   8160
      MaxLength       =   100
      TabIndex        =   17
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   6480
      MaxLength       =   20
      TabIndex        =   16
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   4440
      MaxLength       =   30
      TabIndex        =   15
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   14
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   120
      MaxLength       =   30
      TabIndex        =   13
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   120
      MaxLength       =   100
      TabIndex        =   12
      Top             =   2280
      Width           =   7695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   7335
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label26 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   240
      TabIndex        =   60
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Comentarios 
      AutoSize        =   -1  'True
      Caption         =   "Comentarios"
      Height          =   195
      Left            =   4920
      TabIndex        =   29
      Top             =   7080
      Width           =   870
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "Limite de credito"
      Height          =   195
      Left            =   7920
      TabIndex        =   28
      Top             =   3960
      Width           =   1155
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Fecha alta"
      Height          =   195
      Left            =   2760
      TabIndex        =   27
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña de vtas Web"
      Height          =   195
      Left            =   8760
      TabIndex        =   23
      Top             =   2760
      Width           =   1770
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "CURP"
      Height          =   195
      Left            =   8280
      TabIndex        =   20
      Top             =   1440
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Agente"
      Height          =   195
      Left            =   7920
      TabIndex        =   19
      Top             =   3360
      Width           =   510
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Descuento"
      Height          =   195
      Left            =   7920
      TabIndex        =   11
      Top             =   5160
      Width           =   780
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Contacto"
      Height          =   195
      Left            =   8280
      TabIndex        =   10
      Top             =   2040
      Width           =   645
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Dias Crédito"
      Height          =   195
      Left            =   7920
      TabIndex        =   9
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "FAX"
      Height          =   195
      Left            =   7320
      TabIndex        =   8
      Top             =   2760
      Width           =   300
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Telefono Trabajo"
      Height          =   195
      Left            =   4680
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Telefono Casa"
      Height          =   195
      Left            =   2640
      TabIndex        =   6
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "R.F.C"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   2760
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nombre Comercial"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nombre "
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Clave Cliente"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   930
   End
End
Attribute VB_Name = "AltaClien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalir_Click()
Unload Me
End Sub
Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    Adodc1.Caption = "Registro actual: " & pRecordset.AbsolutePosition
    If Err Or pRecordset.BOF Or pRecordset.EOF Then
        Adodc1.Caption = "Ningún registro activo"
    End If
    Err = 0
End Sub
Private Sub Adodc2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    Adodc2.Caption = "Registro actual: " & pRecordset.AbsolutePosition
    If Err Or pRecordset.BOF Or pRecordset.EOF Then
        Adodc2.Caption = "Ningún registro activo"
    End If
    Err = 0
End Sub
Private Sub Combo1_LostFocus()
   Combo1.Text = Text1(22).Text
End Sub
Private Sub cmdAdd_Click()
    If Text1(1).Text <> "" And Text1(2).Text <> "" And Text1(3).Text <> "" Then
        cmdAdd.Caption = "Guardar/Nuevo"
        Adodc1.Recordset.AddNew
        If Text2.Enabled = True Then
            Me.Text2.Enabled = False
            Me.Text1(22).Visible = False
            Me.Combo1.Visible = True
        End If
    End If
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
Text1(22).Text = Combo1.Text
End Sub
Private Sub DataCombo1_Change()
Text1(22).Text = DataCombo1.Text
End Sub
Private Sub DTPicker1_Change()
Text1(21).Text = DTPicker1.Value
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim Valido As String
Valido = "1234567890."
If Index = 16 Or Index = 14 Or Index = 13 Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 25 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
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
Private Sub DataCombo2_Change()
Text1(24).Text = DataCombo2.Text
End Sub
Private Sub DataCombo3_Change()
Text1(25).Text = DataCombo3.Text
End Sub
Private Sub Buscar(Optional ByVal Siguiente As Boolean = False)
    Dim nReg As Long
    Dim vBookmark As Variant
    Dim sADOBuscar As String
    On Error Resume Next
    If Option1.Value Then
        nReg = Val(Text2)
        sADOBuscar = "ID_CLIENTE = " & nReg
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
    Text1(21).Text = DTPicker1.Value
    Text1(22).Visible = True
    Text2 = ""
    Option2.Value = True
    Dim i As Long
    For i = 0 To 25
        Set Text1(i).DataSource = Adodc1
    Next
    Text1(0).DataField = "ID_CLIENTE"
    Text1(1).DataField = "NOMBRE_COMERCIAL"
    Text1(2).DataField = "RFC"
    Text1(3).DataField = "TELEFONO_CASA"
    Text1(4).DataField = "TELEFONO_TRABAJO"
    Text1(5).DataField = "FAX"
    Text1(6).DataField = "COLONIA"
    Text1(7).DataField = "ESTADO"
    Text1(8).DataField = "CONTACTO"
    Text1(9).DataField = "CIUDAD"
    Text1(10).DataField = "CP"
    Text1(11).DataField = "DIRECCION"
    Text1(12).DataField = "CURP"
    Text1(13).DataField = "NUMERO_EXTERIOR"
    Text1(14).DataField = "NUMERO_INTERIOR"
    Text1(15).DataField = "NOMBRE"
    Text1(16).DataField = "LIMITE_CREDITO"
    Text1(17).DataField = "EMAIL"
    Text1(18).DataField = "WEB_PASSWORD"
    Text1(19).DataField = "DIRECCION_ENVIO"
    Text1(20).DataField = "PAIS"
    Text1(21).DataField = "FECHA_ALTA"
    Text1(22).DataField = "ID_AGENTE"
    Text1(23).DataField = "COMENTARIOS"
    Text1(24).DataField = "DIAS_CREDITO"
    Text1(25).DataField = "DESCUENTO"
End Sub
