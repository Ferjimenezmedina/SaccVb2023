VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRegAlm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REGRESAR A ALMACENES"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adcE 
      Height          =   330
      Left            =   3960
      Top             =   4200
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
      Connect         =   ""
      OLEDBString     =   ""
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
   Begin MSAdodcLib.Adodc adcJR 
      Height          =   330
      Left            =   5160
      Top             =   4200
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
      Connect         =   ""
      OLEDBString     =   ""
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
   Begin MSComctlLib.ListView lvwJR 
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "REPARACION"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "PRODUCTO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "CANTIDAD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "LISTO"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "REGRESAR..."
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2415
      Begin VB.OptionButton Option3 
         Caption         =   "Nada."
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Una Cantidad."
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todo."
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lvwE 
      Height          =   1815
      Left            =   3600
      TabIndex        =   8
      Top             =   2400
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "PRODUCTO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "CANTIDAD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "SUCURSAL"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblP 
      Caption         =   "..."
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
      Left            =   4800
      TabIndex        =   15
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "CANTIDAD:"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "PRODUCTO:"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "SELECIIONE EL PRODUCTO Y ESCRIBA LA CANTIDAD"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   120
      Width           =   4575
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   6720
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   3480
      X2              =   3480
      Y1              =   4680
      Y2              =   600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "EXISTENCIAS EN BODEGA"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label lblJR 
      Alignment       =   2  'Center
      Caption         =   "..."
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
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "JUEGO DE REPARACION DE:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   1920
      Width           =   2535
   End
End
Attribute VB_Name = "frmRegAlm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim tRs As Recordset
Dim tLi As ListItem
Dim SQL As String
Dim NR As Integer
Dim Cont As Integer
Dim PROD As String
Dim CANT As Double
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()
    If Me.txtCantidad.Text = "0" Then
        MsgBox "POR FAVOR, ESPECIFIQUE LA CANTIDAD QUE SE DEVOLVERA A ALMACEN", vbExclamation, "AVISO DEL SISTEMA"
        Me.txtCantidad.SetFocus
    Else
        If Me.Option1.Value = True Then
            NR = Me.lvwJR.ListItems.Count
            For Cont = 1 To NR
                PROD = Me.lvwJR.ListItems.Item(Cont).SubItems(1)
                CANT = Me.lvwJR.ListItems.Item(Cont).SubItems(2)
                SQL = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD + " & CANT & " WHERE ID_PRODUCTO = '" & PROD & "' AND SUCURSAL = 'BODEGA'"
                Set tRs = cnn.Execute(SQL)
            Next Cont
            REFRESHI
            LLENAR_LISTA_EXISTENCIAS
            Me.cmdOk.Enabled = False
        Else
            If Me.Option2.Value = True Then
                If Me.lvwJR.SelectedItem.SubItems(3) = "1" Then
                    MsgBox "YA FUE AUMENTADO", vbExclamation, "AVISO DEL SISTEMA"
                Else
                    PROD = Trim(Me.lblP.Caption)
                    CANT = Trim(Me.txtCantidad.Text) * Val(frmProd.txtNoSIr.Text)
                    SQL = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD + " & CANT & " WHERE ID_PRODUCTO = '" & PROD & "' AND SUCURSAL = 'BODEGA'"
                    Set tRs = cnn.Execute(SQL)
                    REFRESHI
                    LLENAR_LISTA_EXISTENCIAS
                    Me.lvwJR.SelectedItem.SubItems(3) = "1"
                End If
            Else
                Unload Me
            End If
        End If
        If BanLista = 0 Then
            frmProd.lvwComandas_Detalles_Cliente.SelectedItem.SubItems(11) = 1
        Else
            If BanLista = 1 Then
                frmProd.lvwComandas_Detalles_Sucursal.SelectedItem.SubItems(11) = 1
            Else
                MsgBox "ERROR", vbCritical, "AP TONER"
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
    Const sPathBase As String = "LINUX"
    With Me.adcJR
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "JUEGO_REPARACION"
    End With
    With Me.adcE
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "EXISTENCIAS"
    End With
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    SQL = "SELECT * FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & frmProd.txtProducto.Text & "' ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(SQL)
    With tRs
            Me.lvwJR.ListItems.Clear
            Do While Not .EOF
                Set tLi = Me.lvwJR.ListItems.Add(, , RTrim(.Fields("ID_REPARACION")))
                tLi.SubItems(1) = RTrim(.Fields("ID_PRODUCTO"))
                tLi.SubItems(2) = RTrim(.Fields("CANTIDAD"))
                .MoveNext
            Loop
    End With
    REFRESHI
    LLENAR_LISTA_EXISTENCIAS
    Me.lblJR.Caption = Trim(frmProd.txtProducto.Text)
End Sub
Sub LLENAR_LISTA_EXISTENCIAS()
    NR = Me.lvwJR.ListItems.Count
    Me.lvwE.ListItems.Clear
    For Cont = 1 To NR
        PROD = Me.lvwJR.ListItems.Item(Cont).SubItems(1)
        SQL = "SELECT * FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & PROD & "' AND SUCURSAL = 'BODEGA'ORDER BY ID_PRODUCTO"
        Set tRs = cnn.Execute(SQL)
        With tRs
                Do While Not .EOF
                    Set tLi = Me.lvwE.ListItems.Add(, , RTrim(.Fields("ID_EXISTENCIA")))
                    tLi.SubItems(1) = RTrim(.Fields("ID_PRODUCTO"))
                    tLi.SubItems(2) = RTrim(.Fields("CANTIDAD"))
                    tLi.SubItems(3) = RTrim(.Fields("SUCURSAL"))
                    .MoveNext
                Loop
        End With
    Next Cont
End Sub
Sub REFRESHI()
    adcJR.Refresh
    adcE.Refresh
End Sub
Private Sub lvwJR_DblClick()
    Me.lblP.Caption = Trim(Me.lvwJR.SelectedItem.SubItems(1))
    Me.txtCantidad.SetFocus
End Sub
Private Sub Option1_Click()
    Me.txtCantidad.Locked = True
    Me.txtCantidad.Text = "TODO"
    Me.cmdOk.Enabled = True
End Sub
Private Sub Option2_Click()
    Me.txtCantidad.Text = 0
    Me.txtCantidad.Locked = False
    Me.cmdOk.Enabled = True
End Sub
Private Sub Option3_Click()
    Me.txtCantidad.Text = 0
    Me.txtCantidad.Locked = True
    Me.cmdOk.Enabled = False
End Sub
Private Sub txtCantidad_GotFocus()
        Me.txtCantidad.SelStart = 0
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
End Sub
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdOk.Value = True
    Else
            Dim Valido As String
        Valido = "1234567890."
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
