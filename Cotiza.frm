VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Cotiza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cotizacion"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9840
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "EXCEL"
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Height          =   195
      Left            =   2280
      TabIndex        =   42
      Text            =   "Text10"
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Por Codigo de Barras"
      Height          =   255
      Left            =   8160
      TabIndex        =   39
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   1
      Top             =   600
      Width           =   7215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   10080
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   10200
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text2 
      Height          =   195
      Index           =   4
      Left            =   5640
      TabIndex        =   24
      Text            =   "Text2"
      Top             =   5520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   2640
      Width           =   6615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Nombre"
      Height          =   195
      Left            =   8160
      TabIndex        =   22
      Top             =   2400
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por Clave"
      Height          =   195
      Left            =   8160
      TabIndex        =   21
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   195
      Index           =   5
      Left            =   8040
      TabIndex        =   19
      Top             =   5280
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   720
      TabIndex        =   8
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Agregar producto"
      Height          =   375
      Left            =   8520
      TabIndex        =   15
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   840
      TabIndex        =   14
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   6840
      TabIndex        =   10
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9240
      TabIndex        =   13
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   10320
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2880
      Top             =   7320
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
      RecordSource    =   "COTIZACION_DETALLE"
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
      Left            =   4200
      Top             =   7320
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
      RecordSource    =   "COTIZACION"
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
   Begin MSComctlLib.ListView ListView3 
      Height          =   1335
      Left            =   120
      TabIndex        =   20
      Top             =   5880
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   2355
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1335
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   2355
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1335
      Left            =   120
      TabIndex        =   27
      Top             =   960
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   2355
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
   Begin VB.Frame Frame1 
      Caption         =   "Producto Seleccionado"
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   4680
      Width           =   10215
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   4200
         MaxLength       =   5
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   195
         Left            =   8040
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text2 
         Height          =   195
         Index           =   3
         Left            =   7800
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text2 
         Height          =   195
         Index           =   10
         Left            =   3120
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   7
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   5895
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   8
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   9
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   6120
         TabIndex        =   36
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   91226113
         CurrentDate     =   38681
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   3480
         TabIndex        =   38
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   5640
         TabIndex        =   37
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Precio de Venta"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Clave"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label13 
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
      Left            =   3960
      TabIndex        =   44
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label10 
      Caption         =   "Agente"
      Height          =   255
      Left            =   3240
      TabIndex        =   43
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label8 
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
      Left            =   960
      TabIndex        =   41
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Sucursal :"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Buscar producto"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "TOTAL  :"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "# DE COT."
      Height          =   255
      Left            =   8160
      TabIndex        =   28
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Cotiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Private cnn2 As ADODB.Connection
Private WithEvents rst2 As ADODB.Recordset
Attribute rst2.VB_VarHelpID = -1
Private cnn3 As ADODB.Connection
Private WithEvents rst3 As ADODB.Recordset
Attribute rst3.VB_VarHelpID = -1
Private elim As Double
Private elim2 As Double
Private tot As Double
Private Xind As Integer
Private Sub Command1_Click()
    Text2(11).Text = Replace(Text2(11).Text, ",", ".")
    Text2(4).Text = DTPicker1.Value
    Dim x1 As String
    Dim x2 As String
    Dim x3 As String
    Dim x4 As String
    Dim x5 As String
    x1 = Text2(0).Text
    x2 = Text2(1).Text
    x3 = Text2(2).Text
    x4 = Text2(4).Text
    x5 = Text2(11).Text
    If Text2(0).Text = "" Then
        MsgBox "No se a seleccionado cliente para facturacion"
    Else
        Adodc1.Recordset.AddNew
        Text2(0).Text = x1
        Text2(1).Text = x2
        Text2(2).Text = x3
        Text2(4).Text = x4
        Text2(11).Text = x5
        If Text2(0).Text <> "" And Text2(1).Text <> "" And Text2(4).Text <> "" And Text2(11).Text <> "" Then
            Adodc1.Recordset.Update
            Dim nRegcbo As Long
            Dim vBookma As Variant
            Dim sADOBus As String
            On Error Resume Next
            sADOBus = "ORDER BY ID_COTIZACION"
            vBookma = Adodc1.Recordset.Bookmark
            Adodc1.Recordset.MoveLast
            Adodc1.Recordset.Find sADOBus
            Dim rs As Recordset
            Set rs = Adodc1.Recordset
            Set Text8.DataSource = Adodc1
            Text8.DataField = "ID_COTIZACION"
            Dim NumeroRegistros As Integer
            NumeroRegistros = ListView3.ListItems.Count
            Dim Conta As Integer
            ListView3.SetFocus
            Dim Item As MSComctlLib.ListItem
            For Conta = 1 To NumeroRegistros
                Adodc2.Recordset.MoveLast
                Adodc2.Recordset.AddNew
                Text2(8).Text = ListView3.ListItems.Item(Conta)
                Text2(7).Text = ListView3.ListItems.Item(Conta).SubItems(1)
                Text2(10).Text = ListView3.ListItems.Item(Conta).SubItems(2)
                Text2(9).Text = ListView3.ListItems.Item(Conta).SubItems(3)
                Text2(3).Text = ListView3.ListItems.Item(Conta).SubItems(4)
                Text2(5).Text = ListView3.ListItems.Item(Conta).SubItems(5)
                Text2(12).Text = Text8.Text
                Adodc2.Recordset.Update
            Next Conta
            Me.Command1.Enabled = False
        Else
            MsgBox ("ARTICULO NO LISTO PARA VENTA")
        End If
    End If
End Sub
Private Sub Command3_Click()
    If Text5.Text = "" Then
        Text5.Text = "1"
    End If
    Text2(3).Text = Text4.Text
    Text6.Text = Text2(11).Text
    Text2(5).Text = Text5.Text
    Text2(4).Text = DTPicker1.Value
    Me.Command1.Enabled = True
    Dim Clave As String
    Dim PRODU As String
    Dim SUC As String
    SUC = Label8.Caption
    PRODU = Text2(8).Text
    Dim CANT As Double
    Dim prec As Double
    Dim desc As Double
    Dim TOTAL As Double
    Dim cadesc As String
    cadesc = Text2(2).Text
    If cadesc = "" Then
        cadesc = "0.00"
    End If
    Text5.Text = Replace(Text5.Text, ".", ",")
    Text2(9).Text = Replace(Text2(9).Text, ".", ",")
    cadesc = Replace(cadesc, ".", ",")
    CANT = CDbl(Format(Text5.Text, "0.00"))
    prec = CDbl(Format(Text2(9).Text, "0.00"))
    desc = CDbl(Format(cadesc, "0.00"))
    TOTAL = (CDbl(Format(prec, "0.00")) * CDbl(Format(CANT, "0.00"))) - (CDbl(Format(desc, "0.00")) / 100) * (CDbl(Format(prec, "0.00")) * CDbl(Format(CANT, "0.00"))) * 1.15
    If Text6.Text <> "" Then
        Dim ACUM As Double
        ACUM = CDbl(Format(Text6.Text, "0.00"))
        TOTAL = CDbl(Format(TOTAL, "0.00")) + CDbl(Format(ACUM, "0.00"))
    End If
    Text6.Text = CDbl(Format(TOTAL, "0.00"))
    Text2(11).Text = CDbl(Format(TOTAL, "0.00"))
    If Text2(7).Text = "" Or Text2(10).Text = "" Or Text2(9).Text = "" Or Text2(3).Text = "" Or Text2(5).Text = "" Then
        MsgBox ("ARTICULO NO LISTO PARA VENTA")
    Else
        Dim tRs As Recordset
        Dim tLi As ListItem
        If desc = 0 Then
            Set tLi = ListView3.ListItems.Add(, , Text2(8).Text)
                tLi.SubItems(1) = Text2(7).Text
                tLi.SubItems(2) = Text2(10).Text
                tLi.SubItems(3) = Text2(9).Text
                tLi.SubItems(4) = Text2(3).Text
                tLi.SubItems(5) = Text2(5).Text
        Else
            Dim preReal As Double
            preReal = CDbl(Text2(9).Text) - ((CDbl(Text2(2).Text) / 100) * CDbl(Text2(9).Text))
            Set tLi = ListView3.ListItems.Add(, , Text2(8).Text)
                tLi.SubItems(1) = Text2(7).Text
                tLi.SubItems(2) = Text2(10).Text
                tLi.SubItems(3) = preReal
                tLi.SubItems(4) = Text2(3).Text
                tLi.SubItems(5) = Text2(5).Text
            End If
    End If
End Sub
Private Sub Command4_Click()
    If Text10.Text <> "" And elim <> 0 And elim2 <> 0 And Xind <> 0 Then
        ListView3.ListItems.Remove (Xind)
        Command4.Enabled = False
    Else
        MsgBox "No ha seleccionado ningun articulo"
    End If
End Sub
Private Sub Command5_Click()
    Unload Me
End Sub
Private Sub Command6_Click()
    Dim FILE As String
    CommonDialog1.DialogTitle = "Guardar Como"
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "[Excel (*.xls)] |*.xls|"
    CommonDialog1.ShowOpen
    FILE = CommonDialog1.FileName
    Dim ApExcel As Excel.Application
    Set ApExcel = CreateObject("Excel.application")
    ApExcel.Workbooks.Add
    With ApExcel
        .Cells(1, 1) = "ACTITUD POSITIVA EN TONER S DE RL MI"
        .Cells(2, 1) = "ORTIZ DE CAMPOS 1308-1. COLONIA SAN FELIPE, CHIHUAHUA, CHIHUAHUA. CP 31203"
        .Cells(3, 1) = "(614) 414-82-41"
        .Cells(4, 1) = "Cotización " & Text8.Text
        .Cells(5, 1) = "Cliente " & Text1.Text
        .Cells(6, 1) = "Fecha " & Text2(4).Text
        .Cells(7, 1) = "Comentarios "
        .Cells(8, 1) = Text2(6).Text
        .Cells(9, 1) = "Cantidad"
        .Cells(9, 2) = "Producto"
        .Cells(9, 3) = "Precio"
        .Cells(9, 4) = "Importe"
        Dim Subtot As String
        Dim Conta As Integer
        Conta = 0
        Subtot = 0
        Dim NumeroRegistros As Integer
        NumeroRegistros = ListView3.ListItems.Count
        Dim X As Integer
        For X = 1 To NumeroRegistros
            Conta = 10 + X
            .Cells(Conta, 1) = ListView3.ListItems(X).SubItems(5)
            .Cells(Conta, 2) = ListView3.ListItems(X).SubItems(1)
            .Cells(Conta, 3) = ListView3.ListItems(X).SubItems(3)
            .Cells(Conta, 4) = CDbl(Format(ListView3.ListItems(X).SubItems(5), "0.00")) * CDbl(Format(ListView3.ListItems(X).SubItems(3), "0.00"))
            Subtot = CDbl(Format(Subtot, "0.00")) + CDbl(Format(ListView3.ListItems(X).SubItems(5), "0.00")) * CDbl(Format(ListView3.ListItems(X).SubItems(3), "0.00"))
        Next X
        .Cells(Conta + 1, 4) = "Subtotal " & Subtot
        .Cells(Conta + 2, 4) = "IVA " & CDbl(Format(Subtot, "0.00")) * 0.15
        .Cells(Conta + 3, 4) = "Total " & CDbl(Format(Subtot, "0.00")) * 1.15
        .Range("A1:A9").Font.Color = vbBlack
        .Range("A1:A9").VerticalAlignment = xlHAlignCenter
        .Range("A1:A9").Font.Name = "Arial"
        .Range("A1:A9").Font.Size = 10
        .Range("A1:A9").Font.Bold = True
        .Range("A1:A9").Borders.LineStyle = xlContinuous
    End With
    With ApExcel.Range("A1:B10")
    .HorizontalAlignment = xlHAlignLeft
    .VerticalAlignment = xlHAlignCenter
    .WrapText = True
    End With
    ApExcel.AlertBeforeOverwriting = False
    ApExcel.ActiveWorkbook.SaveAs "" & FILE
    ApExcel.Visible = True
    Set ApExcel = Nothing
End Sub
Private Sub Form_Load()
    Command4.Enabled = False
    Label13.Caption = MENU.Text1(1).Text
    Label8.Caption = MENU.Text4(0).Text
    Me.Command1.Enabled = False
    Me.DTPicker1.Value = Date
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Set cnn2 = New ADODB.Connection
    Set rst2 = New ADODB.Recordset
    Set cnn3 = New ADODB.Connection
    Set rst3 = New ADODB.Recordset
    Const sPathBase As String = "LINUX"
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    rst.Open "SELECT * FROM CLIENTE", cnn, adOpenDynamic, adLockOptimistic
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "# DEL CLIENTE", 2400
        .ColumnHeaders.Add , , "NOMBRE", 6100
        .ColumnHeaders.Add , , "DESCUENTO", 2300
    End With
    With cnn2
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    rst2.Open "SELECT * FROM ALMACEN1", cnn2, adOpenDynamic, adLockOptimistic
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "CLAVE DEL PRODUCTO", 2400
        .ColumnHeaders.Add , , "DESCRIPCION", 6800
        .ColumnHeaders.Add , , "% de GANANCIA", 0
        .ColumnHeaders.Add , , "PRECIO_COSTO", 0
    End With
    With cnn3
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    With ListView3
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "CLAVE DEL PRODUCTO", 2100
        .ColumnHeaders.Add , , "DESCRIPCION", 6000
        .ColumnHeaders.Add , , "% de GANANCIA", 0
        .ColumnHeaders.Add , , "PRECIO DE VENTA", 1800
        .ColumnHeaders.Add , , "PRECIO DE COSTO", 0
        .ColumnHeaders.Add , , "CANTIDAD", 1100
    End With
    Set Text2(0).DataSource = Adodc1
    Set Text2(1).DataSource = Adodc1
    Set Text2(2).DataSource = Adodc1
    Set Text2(4).DataSource = Adodc1
    Set Text2(11).DataSource = Adodc1

    Text2(0).DataField = "ID_CLIENTE"
    Text2(1).DataField = "NOMBRE"
    Text2(2).DataField = "DESCUENTO"
    Text2(4).DataField = "FECHA"
    Text2(11).DataField = "TOTAL"
    
    Set Text2(3).DataSource = Adodc2
    Set Text2(7).DataSource = Adodc2
    Set Text2(8).DataSource = Adodc2
    Set Text2(9).DataSource = Adodc2
    Set Text2(10).DataSource = Adodc2
    Set Text2(5).DataSource = Adodc2
    Set Text2(12).DataSource = Adodc2
    
    Text2(3).DataField = "PRECIO_COSTO"
    Text2(7).DataField = "DESCRIPCION"
    Text2(8).DataField = "ID_PRODUCTO"
    Text2(9).DataField = "PRECIO_VENTA"
    Text2(10).DataField = "GANANCIA"
    Text2(5).DataField = "CANTIDAD"
    Text2(12).DataField = "ID_COTIZACION"

    Text2(0).Text = ""
    Text2(1).Text = ""
    Text2(2).Text = ""
    Text2(4).Text = ""
    Text2(11).Text = ""
    Text2(3).Text = ""
    Text2(7).Text = ""
    Text2(8).Text = ""
    Text2(9).Text = ""
    Text2(10).Text = ""
    Text2(5).Text = ""
    Text2(12).Text = ""
    Text2(6).Text = Label8.Caption
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text2(0).Text = Item
    Text2(1).Text = Item.SubItems(1)
    Text2(2).Text = Item.SubItems(2)
    Text1.Text = Item.SubItems(1)
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text2(3).Text = Item.SubItems(3)
    If Item.SubItems(2) = "" Then
        Text2(10).Text = 0
    Else
        Text2(10).Text = Item.SubItems(2)
    End If
    Text4.Text = Item.SubItems(3)
    Text2(8).Text = Item
    Text2(7).Text = Item.SubItems(1)
    Dim cod As String
    cod = Text2(8).Text
    Dim Porci As Double
    Dim valor As Double
    If Item.SubItems(2) <> "" Or Val(Item.SubItems(2)) <> 0 Then
        Porci = Val(Item.SubItems(2))
        If Text2(3).Text <> "" Then
            valor = Text2(3).Text
            valor = Val(valor)
            Porci = (Porci / 100) + 1
            valor = valor * Porci
            Text2(9).Text = valor
        Else
            MsgBox ("Debe dar precio de venta")
            Text2(9).Text = ""
        End If
    Else
        MsgBox ("Debe dar precio de venta")
        Text2(9).Enabled = True
        Text2(9).Text = ""
        Label9.Enabled = True
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text10.Text = Item
    elim = Item.SubItems(5)
    elim2 = Item.SubItems(3)
    Xind = Item.Index
    Command4.Enabled = True
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As Recordset
        Dim tLi As ListItem
        sBuscar = Text1.Text
        sBuscar = Replace(sBuscar, "*", "%")
        sBuscar = Replace(sBuscar, "?", "_")
    
        Text1.Text = sBuscar
        sBuscar = "SELECT * FROM CLIENTE WHERE NOMBRE LIKE '%" & sBuscar & "%' ORDER BY NOMBRE"
        Set tRs = cnn.Execute(sBuscar)
        With tRs
                ListView1.ListItems.Clear
                Do While Not .EOF
                    Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                    tLi.SubItems(1) = .Fields("NOMBRE") & ""
                    tLi.SubItems(2) = .Fields("DESCUENTO") & ""
                    .MoveNext
                Loop
        End With
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text2_Change(Index As Integer)
    Text6.Text = Text2(11).Text
    If Index = 10 Then
        Text2(3).Text = Text4.Text
        Dim Porci As Double
        Dim valor As Double
        Text2(3).Text = Text4.Text
        If Text2(10).Text <> "" Or Val(Text2(10).Text) <> 0 Then
            Porci = Val(Text2(10).Text)
            If Text2(3).Text <> "" Then
                valor = Val(Text2(3).Text)
                Porci = (Porci / 100) + 1
                valor = valor * Porci
                Text2(9).Text = valor
            End If
        Else
            Porci = Val(Text2(2).Text)
            valor = Val(Text2(3).Text)
            Porci = (Porci / 100) + 1
            valor = valor * Porci
            Text2(9).Text = valor
        End If
    End If
    If Text2(11).Text < "0" Then
        Text2(11).Text = "0"
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Option3.Value Then
            Text3.Text = Replace(Text3.Text, ",", "")
            Text3.Text = Replace(Text3.Text, "-", "")
            Text3.Text = Replace(Text3.Text, "_", "")
            Text3.Text = Replace(Text3.Text, ".", "")
            Text3.Text = Replace(Text3.Text, "*", "")
            Text3.Text = Replace(Text3.Text, "%", "")
            Text3.Text = Replace(Text3.Text, "&", "")
            Text3.Text = Replace(Text3.Text, "/", "")
            Text3.Text = Replace(Text3.Text, "'", "")
            Text3.Text = Replace(Text3.Text, "$", "")
            Text3.Text = Replace(Text3.Text, "=", "")
            Text3.Text = Replace(Text3.Text, "@", "")
            Text3.Text = Replace(Text3.Text, "!", "")
            Text3.Text = Replace(Text3.Text, "?", "")
            Text3.Text = Replace(Text3.Text, "^", "")
            Text3.Text = Replace(Text3.Text, "#", "")
            Text3.Text = Replace(Text3.Text, " ", "")
            Text3.Text = Replace(Text3.Text, "+", "")
            Text3.Text = Replace(Text3.Text, ";", "")
            Text3.Text = Replace(Text3.Text, ":", "")
        End If
        Dim tRs As Recordset
        Dim tLi As ListItem
        Const sPathBase As String = "LINUX"
        Set cnn = New ADODB.Connection
        Set rst = New ADODB.Recordset
        With cnn
            .ConnectionString = _
                "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
                "Data Source=" & sPathBase & ";"
            .Open
        End With
        Dim Query As String
        Dim bus As String
        Dim sBus As String
        Query = Text3.Text
        If Option2.Value = True Then
            sBus = "SELECT * FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Query & "%'"
            Set tRs = cnn.Execute(sBus)
            With tRs
                ListView2.ListItems.Clear
                Do While Not .EOF
                    Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        tLi.SubItems(1) = .Fields("DESCRIPCION") & ""
                        tLi.SubItems(2) = .Fields("GANANCIA") & ""
                        tLi.SubItems(3) = .Fields("PRECIO_COSTO")
                        .MoveNext
                Loop
            End With
        End If
        If Option1.Value = True Then
            sBus = "SELECT * FROM ALMACEN3 WHERE DESCRIPCION LIKE '%" & Query & "%'"
            Set tRs = cnn.Execute(sBus)
            With tRs
                ListView2.ListItems.Clear
                Do While Not .EOF
                    Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        tLi.SubItems(1) = .Fields("DESCRIPCION") & ""
                        tLi.SubItems(2) = .Fields("GANANCIA") & ""
                        tLi.SubItems(3) = .Fields("PRECIO_COSTO")
                        .MoveNext
                Loop
            End With
        End If
        If Option3.Value = True Then
            sBus = "SELECT * FROM ENTRADA_PRODUCTO WHERE CODIGO_BARAS LIKE '%" & Query & "%'"
            Set tRs = cnn.Execute(sBus)
            With tRs
                ListView2.ListItems.Clear
                Do While Not .EOF
                    Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    tLi.SubItems(1) = .Fields("CODIGO_BARAS") & ""
                    .MoveNext
                Loop
            End With
            On Error Resume Next
            If ListView2.ListItems.Item(1) <> "" Then
                Dim NEWBUS As String
                NEWBUS = ListView2.ListItems.Item(1)
                sBus = "SELECT * FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & NEWBUS & "%'"
                Set tRs = cnn.Execute(sBus)
                With tRs
                    ListView2.ListItems.Clear
                    Do While Not .EOF
                        Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        tLi.SubItems(1) = .Fields("DESCRIPCION") & ""
                        tLi.SubItems(2) = .Fields("GANANCIA") & ""
                        tLi.SubItems(3) = .Fields("PRECIO_COSTO")
                        .MoveNext
                    Loop
                End With
            End If
        End If
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text4_Change()
    Text2(3).Text = Text4.Text
    Dim Porci As Double
    Dim valor As Double
    Text2(3).Text = Text4.Text
    If Text2(10).Text <> "" Or Val(Text2(10).Text) <> 0 Then
        Porci = Val(Text2(10).Text)
        If Text2(3).Text <> "" Then
            valor = Val(Text2(3).Text)
            Porci = (Porci / 100) + 1
            valor = valor * Porci
            Text2(9).Text = valor
        Else
            MsgBox ("Debe dar precio de venta")
            Text2(9).Text = ""
        End If
    Else
        Porci = Val(Text2(2).Text)
        valor = Val(Text2(3).Text)
        Porci = (Porci / 100) + 1
        valor = valor * Porci
        Text2(9).Text = valor
    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
         If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text5_Change()
    Text2(5).Text = Text5.Text
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
