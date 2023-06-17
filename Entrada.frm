VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   Caption         =   "Registro de Entradas"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   LinkTopic       =   "Form4"
   ScaleHeight     =   3735
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc AdoEnt 
      Height          =   330
      Left            =   7680
      Top             =   3240
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AP Toner;Data Source=VENTAS"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AP Toner;Data Source=VENTAS"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ENTRADAS"
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
   Begin VB.CommandButton Salir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox AyuTxt 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.ComboBox CboBUs 
      Height          =   315
      ItemData        =   "Entrada.frx":0000
      Left            =   1800
      List            =   "Entrada.frx":0002
      TabIndex        =   14
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox AuxTxt 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSAdodcLib.Adodc AdoUsu 
      Height          =   330
      Left            =   120
      Top             =   3240
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AP Toner;Data Source=VENTAS"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AP Toner;Data Source=VENTAS"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "USUARIOS"
      Caption         =   "Adodc1"
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
   Begin MSComCtl2.DTPicker DTFecha 
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   21364737
      CurrentDate     =   38667
   End
   Begin VB.TextBox MuesTxt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7920
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox BusTxt 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   240
      Width           =   6375
   End
   Begin MSComctlLib.ListView ListProv 
      Height          =   1455
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   2566
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Guardar 
      Caption         =   "Guardar/Nuevo"
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox MaTxt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   4680
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox MaTxt 
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   10
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox MaTxt 
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   11
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox MaTxt 
      Height          =   285
      Index           =   2
      Left            =   9720
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox MaTxt 
      Height          =   285
      Index           =   1
      Left            =   9720
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox MaTxt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   7320
      TabIndex        =   16
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label usu 
      Caption         =   "Usuario que capturo"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label num 
      Caption         =   "Numero de Factura"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label tot 
      Caption         =   "Total"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label fech 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   7320
      TabIndex        =   3
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label prov 
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label ent 
      Caption         =   "Numero de Entrada"
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Private Sub AdoUsu_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    If Err Or pRecordset.BOF Or pRecordset.EOF Then
        AdoUsu.Caption = "Ningún registro activo"
    End If
    Err = 0
End Sub
Private Sub AdoEnt_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    AdoEnt.Caption = "Registro actual: " & pRecordset.AbsolutePosition
    If Err Or pRecordset.BOF Or pRecordset.EOF Then
        AdoEnt.Caption = "Ningún registro activo"
    End If

    Err = 0
End Sub
Private Sub CboBUs_LostFocus()
    CboBUs.Text = AyuTxt.Text
    Dim BQue As String
    BQue = "NOMBRE Like '" & CboBUs.Text & "'"
    AdoUsu.Recordset.MoveFirst
    AdoUsu.Recordset.Find BQue
    Set AuxTxt.DataSource = AdoUsu
    AuxTxt.DataField = "ID_USUARIO"
    MaTxt(5).Text = AuxTxt.Text
End Sub
Private Sub CboBUs_Validate(Cancel As Boolean)
    AyuTxt.Text = CboBUs.Text
End Sub
Private Sub Guardar_Click()
    MaTxt(2).Text = DTFecha.Value
    If MaTxt(1).Text <> "" And MaTxt(2).Text <> "" And MaTxt(3).Text <> "" Then
        AdoEnt.Recordset.Update
        AdoEnt.Recordset.AddNew
    End If
End Sub

Private Sub Salir_Click()
    Unload Me
End Sub
Private Sub DTFecha_CloseUp()
    MaTxt(2).Text = DTFecha.Value
End Sub
Private Sub Form_Load()
    Const sPathBase As String = "VENTAS"
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AP Toner;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    rst.Open "SELECT * FROM PROVEEDOR", cnn, adOpenDynamic, adLockOptimistic
    With ListProv
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "CLAVE DEL PROVEEDOR", 2400
        .ColumnHeaders.Add , , "NOMBRE", 6100
        .ColumnHeaders.Add , , "CIUDAD", 2300
    End With
    With Me.AdoUsu
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AP Toner;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "USUARIOS"
    End With
    With Me.AdoEnt
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AP Toner;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "ENTRADAS"
    End With
    Dim i As Integer
    For i = 0 To 5
        Set MaTxt(i).DataSource = AdoEnt
    Next
    MaTxt(0).DataField = "ID_ENTRADA"
    MaTxt(1).DataField = "ID_PROVEEDOR"
    MaTxt(2).DataField = "FECHA"
    MaTxt(3).DataField = "TOTAL"
    MaTxt(4).DataField = "FACTURA"
    MaTxt(5).DataField = "ID_USUARIO"
    Set CboBUs.DataSource = AdoUsu
    CboBUs.DataField = "NOMBRE"
    AdoEnt.Recordset.AddNew
End Sub
Private Sub ListProv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    MaTxt(1).Text = Item
    MuesTxt.Text = Item
    BusTxt.Text = Item.SubItems(1)
End Sub
Private Sub BusTxt_Change()
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = BusTxt
    sBuscar = Replace(sBuscar, "*", "%")
    sBuscar = Replace(sBuscar, "?", "_")

    BusTxt = sBuscar
    sBuscar = "SELECT * FROM PROVEEDOR WHERE NOMBRE LIKE '%" & sBuscar & "%' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
            ListProv.ListItems.Clear
            Do While Not .EOF
                Set tLi = ListProv.ListItems.Add(, , .Fields("ID_PROVEEDOR") & "")
                tLi.SubItems(1) = .Fields("NOMBRE") & ""
                tLi.SubItems(2) = .Fields("CIUDAD") & ""
                .MoveNext
            Loop
    End With
End Sub
Private Sub CboBUs_DropDown()
    CboBUs.Clear
    Buscarcbo
End Sub
Private Sub Buscarcbo(Optional ByVal Siguiente As Boolean = False)
    Dim nRegcbo As Long
    Dim vBookmarkcbo As Variant
    Dim sADOBuscarcbo As String
    On Error Resume Next
    sADOBuscarcbo = "NOMBRE Like '" & "%" & "'"
    vBookmarkcbo = AdoUsu.Recordset.Bookmark
        AdoUsu.Recordset.MoveFirst
        AdoUsu.Recordset.Find sADOBuscarcbo
    If AdoUsu.Recordset.BOF Or AdoUsu.Recordset.EOF Then
        Err.Clear
        MsgBox "No existe el dato buscado o ya no hay más datos que mostrar."
        AdoUsu.Recordset.Bookmark = vBookmarkcbo
    End If
    If AdoUsu.Recordset.EOF = False Then
    Do While AdoUsu.Recordset.EOF = False
        CboBUs.AddItem AdoUsu.Recordset.Fields("NOMBRE")
        AdoUsu.Recordset.MoveNext
    Loop
    End If
End Sub
Private Sub CboBUs_KeyPress(KeyAscii As Integer)
     KeyAscii = 0
End Sub
