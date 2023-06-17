VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRevComSuc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PEDIDOS DE SUCURSALES"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCantNoOk 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtCantNo 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      MaxLength       =   15
      TabIndex        =   2
      Top             =   4200
      Width           =   3015
   End
   Begin VB.TextBox txtCantidad_No 
      Height          =   285
      Left            =   7425
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   6705
      TabIndex        =   22
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtComanda2 
      Height          =   285
      Left            =   8145
      TabIndex        =   21
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtArticulo2 
      Height          =   195
      Left            =   8865
      TabIndex        =   20
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtLlego 
      Height          =   375
      Left            =   7080
      TabIndex        =   19
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtComanda 
      Height          =   375
      Left            =   5520
      TabIndex        =   18
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtArticulo 
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Listo"
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox txtNoCom2 
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adcCom 
      Height          =   330
      Left            =   120
      Top             =   7080
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
      CommandType     =   8
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
      RecordSource    =   ""
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
   Begin MSComctlLib.ListView lvwComDet 
      Height          =   1935
      Left            =   4680
      TabIndex        =   10
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ARTICULO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID_PRODUCTO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CANTIDAD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "LLEGO"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdNoComOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox txtNoCom 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      MaxLength       =   15
      TabIndex        =   5
      Top             =   8040
      Width           =   2895
   End
   Begin MSComctlLib.ListView lvwCom 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   11456
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "COMANDA"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FECHA"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CLIENTE"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "USUARIO"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ACTIVO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "REVISADO"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvwComDetNo 
      Height          =   1935
      Left            =   4680
      TabIndex        =   14
      Top             =   5160
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ARTICULO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID_PRODUCTO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CANTIDAD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "LLEGO"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSAdodcLib.Adodc adcComDet 
      Height          =   330
      Left            =   4680
      Top             =   2520
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
      CommandType     =   8
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
      RecordSource    =   ""
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
   Begin VB.Label Label7 
      Caption         =   "5)  Escriba aquì la cantidad que se resta."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   24
      Top             =   3840
      Width           =   4575
   End
   Begin VB.Line Line4 
      X1              =   4680
      X2              =   9120
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
      Caption         =   "3)  Si la lista es correcta de click en LISTO."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   2880
      Width           =   4695
   End
   Begin VB.Line Line2 
      X1              =   4680
      X2              =   9120
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label6 
      Caption         =   "6) Revise la lista."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   4800
      Width           =   4575
   End
   Begin VB.Label Label5 
      Caption         =   "4)  Si no, seleccione el articulo."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   3360
      Width           =   4575
   End
   Begin VB.Line Line3 
      X1              =   4680
      X2              =   9120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label4 
      Caption         =   "...o escriba aquí el numero de comanda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   7680
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "1)   Seleccione la comanda de la lista..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "2)  Revice la lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   240
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   9120
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "frmRevComSuc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim sqlCom As String
Dim sqlComDet As String
Dim tRs As Recordset
Dim tLi As ListItem
Dim Conta As Integer, NoRe As Integer
Dim COM As Integer, COM_DETA As Integer
Private Sub cmdActualizar_Click()
    TRAER_COMANDA
End Sub
Private Sub cmdCantNoOk_Click()
            Set Me.txtLlego.DataSource = Me.adcComDet
            Set Me.txtArticulo2.DataSource = Me.adcComDet
            Set Me.txtComanda2.DataSource = Me.adcComDet
            Set Me.txtCantidad.DataSource = Me.adcComDet
            Set Me.txtCantidad_No.DataSource = Me.adcComDet
            Me.txtLlego.DataField = "LLEGO"
            Me.txtArticulo2.DataField = "ARTICULO"
            Me.txtComanda2.DataField = "COMANDA"
            Me.txtCantidad.DataField = "CANTIDAD"
            Me.txtCantidad_No.DataField = "CANTIDAD_NO"
            Dim C As Integer
            Me.adcComDet.Recordset.MoveFirst
            Do While C = 0
                If Me.txtArticulo2.Text = Me.txtArticulo.Text And Me.txtComanda.Text = Me.txtComanda2.Text Then
                    If Val(Me.txtCantNo.Text) = 0 Then
                        MsgBox "ERROR EN LA CANTIDAD QUE NO LLEGO", vbExclamation, "AVISO DEL SISTEMA"
                        Me.txtCantNo.SetFocus
                        C = 1
                    Else
                        If Val(Me.txtCantNo.Text) > Me.txtCantidad.Text Then
                            MsgBox "ERROR EN LA CANTIDAD QUE NO LLEGO", vbExclamation, "AVISO DEL SISTEMA"
                            Me.txtCantNo.SetFocus
                            C = 1
                        Else
                            Me.txtLlego.Text = "No"
                            Me.txtCantidad_No.Text = Val(Me.txtCantNo.Text)
                            Me.adcComDet.Recordset.MovePrevious
                            Me.adcComDet.Recordset.MoveNext
                            TRAER_COMANDA_DETALLE
                            TRAER_COMANDA_DETALLE_NO
                            Me.txtCantNo.Text = ""
                            Me.lvwComDet.SetFocus
                            C = 1
                        End If
                    End If
                Else
                    adcComDet.Recordset.MoveNext
                    If Me.adcComDet.Recordset.EOF = True Then
                        C = 1
                        MsgBox "NO SE ENCONTRO", vbInformation, "MENSAJE DEL SISTEMA"
                        Me.txtNoCom.SetFocus
                    End If
                End If
            Loop
End Sub
Private Sub cmdNoComOk_Click()
    TRAER_COMANDA_DETALLE
    TRAER_COMANDA_DETALLE_NO
End Sub
Private Sub cmdOk_Click()
    ENVIAR_COMANDA_DETALLE
    ENVIAR_COMANDA_DETALLE_NO
    ENVIAR_COMANDA
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
    Const sPathBase As String = "LINUX"
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    TRAER_COMANDA
End Sub
Private Sub lvwCom_DblClick()
    If Me.lvwCom.ListItems.Count = 0 Then
        MsgBox "NO HAY ARTICULOS", vbExclamation, "AVISO DEL SISTEMA"
    Else
        Me.txtNoCom.Text = Me.lvwCom.SelectedItem
        Me.lvwComDet.SetFocus
        TRAER_COMANDA_DETALLE
        TRAER_COMANDA_DETALLE_NO
    End If
End Sub
Private Sub lvwCom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.lvwCom.ListItems.Count = 0 Then
            MsgBox "NO HAY ARTICULOS", vbExclamation, "AVISO DEL SISTEMA"
        Else
            Me.txtNoCom.Text = Me.lvwCom.SelectedItem
            Me.lvwComDet.SetFocus
            TRAER_COMANDA_DETALLE
            TRAER_COMANDA_DETALLE_NO
        End If
    End If
End Sub
Private Sub lvwComDet_DblClick()
        If Me.lvwComDet.ListItems.Count = 0 Then
            MsgBox "NO HAY ARTICULOS", vbExclamation, "AVISO DEL SISTEMA"
        Else
            Me.txtArticulo.Text = Me.lvwComDet.SelectedItem
            Me.txtCantNo.SetFocus
        End If
End Sub
Private Sub lvwComDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.lvwComDet.ListItems.Count = 0 Then
            MsgBox "NO HAY ARTICULOS", vbExclamation, "AVISO DEL SISTEMA"
        Else
            Me.txtArticulo.Text = Me.lvwComDet.SelectedItem
            Me.txtCantNo.SetFocus
        End If
    End If
End Sub
Private Sub lvwComDetNo_DblClick()
        If Me.lvwComDetNo.ListItems.Count = 0 Then
            MsgBox "NO HAY ARTICULOS", vbExclamation, "AVISO DEL SISTEMA"
        Else
            Me.txtArticulo.Text = Me.lvwComDetNo.SelectedItem
            Set Me.txtLlego.DataSource = Me.adcComDet
            Set Me.txtArticulo2.DataSource = Me.adcComDet
            Set Me.txtComanda2.DataSource = Me.adcComDet
            Set Me.txtCantidad.DataSource = Me.adcComDet
            Set Me.txtCantidad_No.DataSource = Me.adcComDet
            Me.txtLlego.DataField = "LLEGO"
            Me.txtArticulo2.DataField = "ARTICULO"
            Me.txtComanda2.DataField = "COMANDA"
            Me.txtCantidad.DataField = "CANTIDAD"
            Me.txtCantidad_No.DataField = "CANTIDAD_NO"
            adcComDet.Recordset.MoveFirst
            Dim C As Integer
            Do While C = 0
                If Me.txtArticulo2.Text = Me.txtArticulo.Text And Me.txtComanda.Text = Me.txtComanda2.Text Then
                    Me.txtLlego.Text = "Si"
                    Me.txtCantidad_No.Text = 0
                    Me.adcComDet.Recordset.MovePrevious
                    Me.adcComDet.Recordset.MoveNext
                    TRAER_COMANDA_DETALLE
                    TRAER_COMANDA_DETALLE_NO
                    C = 1
                Else
                    adcComDet.Recordset.MoveNext
                    If Me.adcComDet.Recordset.EOF = True Then
                        C = 1
                        MsgBox "NO SE ENCONTRO", vbInformation, "MENSAJE DEL SISTEMA"
                    End If
                End If
            Loop
        End If
End Sub
Private Sub txtCantNo_GotFocus()
    Me.txtCantNo.SelStart = 0
    Me.txtCantNo.SelLength = Len(Me.txtCantNo.Text)
End Sub
Private Sub txtCantNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdCantNoOk.Value = True
    End If
End Sub
Sub TRAER_COMANDA()
    Const sPathBase As String = "LINUX"
    With Me.adcCom
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "COMANDAS"
    End With
    sqlCom = "SELECT * FROM COMANDAS WHERE CLIENTE = 0 AND REVISADO='No' ORDER BY COMANDA"
    Set tRs = cnn.Execute(sqlCom)
    With tRs
            Me.lvwCom.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwCom.ListItems.Add(, , .Fields("COMANDA"))
                tLi.SubItems(1) = .Fields("FECHA")
                tLi.SubItems(2) = .Fields("CLIENTE")
                tLi.SubItems(3) = .Fields("USUARIO")
                tLi.SubItems(4) = .Fields("ACTIVO")
                tLi.SubItems(5) = .Fields("REVISADO")
                .MoveNext
            Loop
    End With
End Sub
Sub TRAER_COMANDA_DETALLE()
    Const sPathBase As String = "LINUX"
    With Me.adcComDet
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "COMANDAS_DETALLES"
    End With
    sqlComDet = "SELECT * FROM COMANDAS_DETALLES WHERE COMANDA = " & Val(Me.txtNoCom.Text) & " AND LLEGO = 'SI' ORDER BY ARTICULO"
    Set tRs = cnn.Execute(sqlComDet)
    With tRs
            Me.lvwComDet.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwComDet.ListItems.Add(, , .Fields("ARTICULO"))
                tLi.SubItems(1) = .Fields("ID_PRODUCTO")
                tLi.SubItems(2) = .Fields("CANTIDAD")
                tLi.SubItems(3) = .Fields("LLEGO")
                Me.txtComanda.Text = .Fields("COMANDA")
                .MoveNext
            Loop
    End With
End Sub
Sub TRAER_COMANDA_DETALLE_NO()
    sqlComDet = "SELECT * FROM COMANDAS_DETALLES WHERE COMANDA = " & Val(Me.txtNoCom.Text) & " AND LLEGO = 'NO' ORDER BY ARTICULO"
    Set tRs = cnn.Execute(sqlComDet)
    With tRs
            Me.lvwComDetNo.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwComDetNo.ListItems.Add(, , .Fields("ARTICULO"))
                tLi.SubItems(1) = .Fields("ID_PRODUCTO")
                tLi.SubItems(2) = .Fields("CANTIDAD_NO")
                tLi.SubItems(3) = .Fields("LLEGO")
                Me.txtComanda.Text = .Fields("COMANDA")
                .MoveNext
            Loop
    End With
End Sub
Private Sub txtNoCom_GotFocus()
    Me.txtNoCom.SelStart = 0
    Me.txtNoCom.SelLength = Len(Me.txtNoCom.Text)
End Sub
Private Sub txtNoCom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdNoComOk.Value = True
        Me.lvwComDet.SetFocus
    End If
End Sub
Sub ENVIAR_COMANDA()
    Dim cComanda As Double
    Dim sqlComanda As String
    If Val(Me.txtNoCom.Text) = 0 Then
        cComanda = Val(Trim(InputBox("¿CUAL ES EL NUMERO DE LA COMANDA?", "REVICIÓN DE COMANDAS")))
    Else
        cComanda = Val(Me.txtNoCom.Text)
    End If
        sqlComanda = "COMANDA = " & cComanda
        Set Me.txtNoCom2.DataSource = Me.adcCom
        Me.txtNoCom2.DataField = "REVISADO"
        With Me.adcCom.Recordset
            If .EOF Or .BOF Then
                MsgBox "NO HAY ARTICULOS", vbExclamation, "AVISO DEL SISTEMA"
            Else
                adcCom.Recordset.MoveFirst
                adcCom.Recordset.Find sqlComanda
                Me.txtNoCom2.Text = "Si"
                Me.adcCom.Recordset.MovePrevious
                Me.adcCom.Recordset.MoveNext
                TRAER_COMANDA
                Me.lvwComDet.ListItems.Clear
                Me.lvwComDetNo.ListItems.Clear
                Me.lvwCom.SetFocus
                Me.txtNoCom.Text = ""
            End If
        End With
End Sub
Sub ENVIAR_COMANDA_DETALLE()
    NoRe = Me.lvwComDet.ListItems.Count
    COM = Me.txtComanda.Text
    For Conta = 1 To NoRe
        COM_DETA = Me.lvwComDet.ListItems(Conta)
        sqlCom = "UPDATE COMANDAS_DETALLES SET REVISADO=1 WHERE COMANDA=" & COM & " AND ARTICULO=" & COM_DETA
        Set tRs = cnn.Execute(sqlCom)
    Next Conta
End Sub
Sub ENVIAR_COMANDA_DETALLE_NO()
    NoRe = Me.lvwComDetNo.ListItems.Count
    COM = Me.txtComanda.Text
    For Conta = 1 To NoRe
        COM_DETA = Me.lvwComDetNo.ListItems(Conta)
        sqlCom = "UPDATE COMANDAS_DETALLES SET REVISADO=1 WHERE COMANDA=" & COM & " AND ARTICULO=" & COM_DETA
        Set tRs = cnn.Execute(sqlCom)
    Next Conta
End Sub
