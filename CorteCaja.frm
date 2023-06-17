VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form9 
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   LinkTopic       =   "Form9"
   ScaleHeight     =   6375
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8880
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   5880
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   6000
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
      RecordSource    =   "VENTAS"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   5880
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   4815
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8493
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   21168129
      CurrentDate     =   38701
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8280
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Chihuahua"
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
      Left            =   5760
      TabIndex        =   10
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Ciudad"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Americas"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label6 
      Caption         =   "Sucursal"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "El Churritos"
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
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label Label4 
      Caption         =   "Agente"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   6960
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn2 As ADODB.Connection
Attribute cnn2.VB_VarHelpID = -1
Private WithEvents rst2 As ADODB.Recordset
Attribute rst2.VB_VarHelpID = -1
Private Sub Command1_Click()
    Dim nReg As Long
    Dim vBookmark As Variant
    Dim sADOBuscar As String
    On Error Resume Next
    sADOBuscar = "FECHA = '" & Text1.Text & "'"
    vBookmark = Adodc1.Recordset.Bookmark
    With Adodc1.Recordset
        .MoveFirst
        Do While Not .EOF
            .Find sADOBuscar, 1
        Loop
    End With
    If Err.Number Or Adodc1.Recordset.BOF Or Adodc1.Recordset.EOF Then
        Err.Clear
        MsgBox "No existe el dato buscado o ya no hay más datos que mostrar."
        'Adodc1.Recordset.Bookmark = vBookmark
    End If
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
Printer.Print "Corte de Caja de Suc. " & Label7.Caption
Printer.Print "Ciudad : " & Label9.Caption
Printer.Print "Fecha : " & DTPicker1.Value
Dim POSY As Integer
Dim acum As String
acum = "0"
POSY = 800
Printer.CurrentY = POSY
Printer.CurrentX = 100
Printer.Print "Producto"
Printer.CurrentY = POSY
Printer.CurrentX = 1300
Printer.Print "Precio unitario"
Printer.CurrentY = POSY
Printer.CurrentX = 3000
Printer.Print "Total"
Printer.Print "--------------------------------------------------------------------------------"
Dim NumeroRegistros As Integer
NumeroRegistros = ListView2.ListItems.Count
Dim Conta As Integer
POSY = POSY + 200
For Conta = 1 To NumeroRegistros
    POSY = POSY + 200
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print ListView2.ListItems.Item(Conta)
    Printer.CurrentY = POSY
    Printer.CurrentX = 1300
    Printer.Print ListView2.ListItems.Item(Conta).SubItems(2)
    Printer.CurrentY = POSY
    Printer.CurrentX = 3000
    Printer.Print ListView2.ListItems.Item(Conta).SubItems(4)
    acum = CDbl(Format(acum, "0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(4), "0.00"))
Next Conta
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.CurrentY = POSY + 400
    Printer.CurrentX = 2500
    Printer.Print "TOTAL : " & Format(acum, "0.00")
End Sub
Private Sub Form_Load()
    Label5.Caption = Form8.Text1(1).Text
    Label7.Caption = Form8.Text4(0).Text
    Label9.Caption = Form8.Text4(3).Text
    DTPicker1.Value = Date
    Text1.Text = DTPicker1.Value
    Const sPathBase As String = "VENTAS"
    Set cnn2 = New ADODB.Connection
    Set rst2 = New ADODB.Recordset
    With cnn2
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AP Toner;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    rst2.Open "SELECT * FROM VENTAS", cnn2, adOpenDynamic, adLockOptimistic
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID_PRODUCTO", 2700
        .ColumnHeaders.Add , , "DESCRIPCION", 5000, lvwColumnCenter
        .ColumnHeaders.Add , , "CANTIDAD", 2000, lvwColumnCenter
        .ColumnHeaders.Add , , "PRECIO VENTA", 3000, lvwColumnCenter
        .ColumnHeaders.Add , , "TOTAL", 3000, lvwColumnCenter
    End With
    Set Text2.DataSource = Adodc1
    Text2.DataField = "ID_VENTA"
    ListView2.ListItems.Clear
End Sub
Private Sub Text2_Change()
    Dim sBuscar2 As String
    Dim tRs2 As Recordset
    Dim tLi2 As ListItem
    Dim exe As String
    sBuscar2 = Text2.Text
    sBuscar2 = Replace(sBuscar2, "*", "%")
    sBuscar2 = Replace(sBuscar2, "?", "_")
    Text2.Text = sBuscar2
    If sBuscar2 <> "" Then
        sBuscar2 = "SELECT * FROM VENTAS WHERE ID_VENTA = " & sBuscar2 '& ""
        Set tRs2 = cnn2.Execute(sBuscar2)
        With tRs2
            If Not (.BOF And .EOF) Then
                exe = .Fields("FACTURADO")
            End If
        End With
        If exe = "" Or exe = "0" Then
            sBuscar2 = Val(Text2.Text)
            sBuscar2 = "SELECT * FROM VENTAS_DETALLE WHERE ID_VENTA = " & sBuscar2 & ""
            Set tRs2 = cnn2.Execute(sBuscar2)
            With tRs2
                If Not (.BOF And .EOF) Then
                    'ListView2.ListItems.Clear
                    .MoveFirst
                    Do While Not .EOF
                        Set tLi2 = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        tLi2.SubItems(1) = .Fields("DESCRIPCION") & ""
                        tLi2.SubItems(2) = .Fields("CANTIDAD") & ""
                        tLi2.SubItems(3) = .Fields("PRECIO_VENTA") & ""
                        tLi2.SubItems(4) = Format(CDbl(.Fields("PRECIO_VENTA")), "0.00") * Format(CDbl(.Fields("CANTIDAD")), "0.00") & ""
                        .MoveNext
                    Loop
                End If
            End With
        End If
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
