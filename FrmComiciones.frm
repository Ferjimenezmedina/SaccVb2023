VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmComiciones 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comiciónes de ventas"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   855
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   46
      Text            =   "FrmComiciones.frx":0000
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Solo Facturas"
      Height          =   255
      Left            =   7440
      TabIndex        =   45
      Top             =   1800
      Width           =   1455
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usuario del Cliente"
      Height          =   255
      Left            =   5040
      TabIndex        =   44
      Top             =   1800
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usuario de la Venta"
      Height          =   255
      Left            =   5040
      TabIndex        =   43
      Top             =   1560
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1080
      TabIndex        =   41
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Servicios"
      Height          =   255
      Left            =   6960
      TabIndex        =   40
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   8160
      TabIndex        =   37
      Top             =   120
      Width           =   1215
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Totales"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A Detalle"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   6960
      TabIndex        =   33
      Top             =   120
      Width           =   1095
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contado"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Credito"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todo"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9600
      TabIndex        =   31
      Top             =   3840
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmComiciones.frx":0006
         MousePointer    =   99  'Custom
         Picture         =   "FrmComiciones.frx":0310
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reporte"
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
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Originales"
      Height          =   255
      Left            =   5760
      TabIndex        =   30
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cambios"
      Height          =   255
      Left            =   4560
      TabIndex        =   29
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Compatibles"
      Height          =   255
      Left            =   3120
      TabIndex        =   28
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remanofacturas"
      Height          =   255
      Left            =   1440
      TabIndex        =   27
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Recargas"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   1200
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9960
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9600
      TabIndex        =   22
      Top             =   5040
      Width           =   975
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EXCEL"
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
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmComiciones.frx":089F
         MousePointer    =   99  'Custom
         Picture         =   "FrmComiciones.frx":0BA9
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      Picture         =   "FrmComiciones.frx":26EB
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rango de Fechas"
      Height          =   975
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   3615
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17104897
         CurrentDate     =   39253
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17104897
         CurrentDate     =   39253
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Al :"
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Del :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9600
      TabIndex        =   4
      Top             =   6240
      Width           =   975
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmComiciones.frx":50BD
         MousePointer    =   99  'Custom
         Picture         =   "FrmComiciones.frx":53C7
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmComiciones.frx":74A9
      Left            =   1080
      List            =   "FrmComiciones.frx":74AB
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Clientes Asignados"
      TabPicture(0)   =   "FrmComiciones.frx":74AD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Totales Ventas"
      TabPicture(1)   =   "FrmComiciones.frx":74C9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text5"
      Tab(1).Control(1)=   "Text4"
      Tab(1).Control(2)=   "Text3"
      Tab(1).Control(3)=   "Text2"
      Tab(1).Control(4)=   "Text1"
      Tab(1).Control(5)=   "Label8"
      Tab(1).Control(6)=   "Label7"
      Tab(1).Control(7)=   "Label6"
      Tab(1).Control(8)=   "Label5"
      Tab(1).Control(9)=   "Label4"
      Tab(1).ControlCount=   10
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   -71040
         TabIndex        =   24
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   -71040
         TabIndex        =   19
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -71040
         TabIndex        =   18
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -71040
         TabIndex        =   17
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -71040
         TabIndex        =   16
         Top             =   1440
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7858
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label8 
         Caption         =   "Originales"
         Height          =   255
         Left            =   -72600
         TabIndex        =   25
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Cambios"
         Height          =   255
         Left            =   -72600
         TabIndex        =   15
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Compatibles"
         Height          =   255
         Left            =   -72600
         TabIndex        =   14
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Remanofacturados"
         Height          =   255
         Left            =   -72600
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Recargas"
         Height          =   255
         Left            =   -72600
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usuario :"
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   1680
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   9600
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sucursal :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "FrmComiciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim ClvAgente As String
Dim Sucursal As String
Dim FechaDel As String
Dim FechaAl As String
Private Sub cmdOk_Click()
    If Combo1.Text <> "" Then
        If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 And Check4.Value = 0 And Check5.Value = 0 And Check6.Value = 0 Then
            MsgBox "Debe seleccionar un tipo de producto a evaluar", vbExclamation, "SACC"
        Else
            Dim sBuscar As String
            Dim tRs As ADODB.Recordset
            Dim tLi As ListItem
            Dim sEncab As String
            Dim iTotal As String
            Dim sTipo As String
            Dim sIdUsuario As String
            ListView1.ListItems.Clear
            ListView1.ColumnHeaders.Clear
            If Option4.Value Then
                With ListView1
                    .View = lvwReport
                    .GridLines = True
                    .LabelEdit = lvwManual
                    .HideSelection = False
                    .HotTracking = False
                    .HoverSelection = False
                    .FullRowSelect = True
                    .ColumnHeaders.Add , , "Venta", 1500
                    .ColumnHeaders.Add , , "Producto", 2500
                    .ColumnHeaders.Add , , "Cantidad", 1500
                    .ColumnHeaders.Add , , "Precio", 1500
                    .ColumnHeaders.Add , , "Usuario", 1700
                End With
            Else
                With ListView1
                    .View = lvwReport
                    .GridLines = True
                    .LabelEdit = lvwManual
                    .HideSelection = False
                    .HotTracking = False
                    .HoverSelection = False
                    .FullRowSelect = True
                    .ColumnHeaders.Add , , "Producto", 2500
                    .ColumnHeaders.Add , , "Cantidad", 1500
                    .ColumnHeaders.Add , , "Precio", 1500
                    .ColumnHeaders.Add , , "Sucursal", 1700
                End With
            End If
            If Combo2.Text <> "" Then
                sBuscar = "SELECT ID_USUARIO FROM USUARIOS WHERE NOMBRE = '" & Combo2.Text & "'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    sIdUsuario = tRs.Fields("ID_USUARIO")
                End If
            End If
            If Option1.Value Then
                sTipo = " AND UNA_EXIBICION  = 'S'"
            Else
                If Option2.Value Then
                    sTipo = " AND UNA_EXIBICION  = 'N'"
                Else
                    sTipo = ""
                End If
            End If
            Sucursal = Combo1.Text
            ListView1.ListItems.Clear
            If Option4.Value Then
                If Combo1.Text = "<TODAS>" Then
                    sEncab = "SELECT VENTAS.ID_VENTA, VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.CANTIDAD, VENTAS_DETALLE.PRECIO_VENTA, USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS USUARIO FROM VENTAS INNER JOIN VENTAS_DETALLE ON VENTAS.ID_VENTA = VENTAS_DETALLE.ID_VENTA INNER JOIN USUARIOS ON VENTAS.ID_USUARIO = USUARIOS.ID_USUARIO INNER JOIN ALMACEN3 ON VENTAS_DETALLE.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO INNER JOIN CLIENTE ON VENTAS.ID_CLIENTE = CLIENTE.ID_CLIENTE WHERE (VENTAS.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " 23:59:59.997') AND VENTAS.FACTURADO IN (0,1) AND"
                Else
                    sEncab = "SELECT VENTAS.ID_VENTA, VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.CANTIDAD, VENTAS_DETALLE.PRECIO_VENTA, USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS USUARIO FROM VENTAS INNER JOIN VENTAS_DETALLE ON VENTAS.ID_VENTA = VENTAS_DETALLE.ID_VENTA INNER JOIN USUARIOS ON VENTAS.ID_USUARIO = USUARIOS.ID_USUARIO INNER JOIN ALMACEN3 ON VENTAS_DETALLE.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO INNER JOIN CLIENTE ON VENTAS.ID_CLIENTE = CLIENTE.ID_CLIENTE WHERE (VENTAS.SUCURSAL = '" & Combo1.Text & "') AND (VENTAS.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " 23:59:59.997') AND VENTAS.FACTURADO IN (0,1) AND"
                End If
            Else
                If Combo1.Text = "<TODAS>" Then
                    sEncab = "SELECT VENTAS_DETALLE.ID_PRODUCTO, SUM(VENTAS_DETALLE.CANTIDAD) AS CANTIDAD, SUM(VENTAS_DETALLE.PRECIO_VENTA * VENTAS_DETALLE.CANTIDAD) AS IMPORTE, VENTAS.SUCURSAL FROM VENTAS INNER JOIN VENTAS_DETALLE ON VENTAS.ID_VENTA = VENTAS_DETALLE.ID_VENTA INNER JOIN ALMACEN3 ON VENTAS_DETALLE.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO INNER JOIN CLIENTE ON VENTAS.ID_CLIENTE = CLIENTE.ID_CLIENTE WHERE (VENTAS.FECHA BETWEEN '01/01/2011' AND '31/12/2011 23:59:59.997') AND "
                Else
                    sEncab = "SELECT VENTAS_DETALLE.ID_PRODUCTO, SUM(VENTAS_DETALLE.CANTIDAD) AS CANTIDAD, SUM(VENTAS_DETALLE.PRECIO_VENTA * VENTAS_DETALLE.CANTIDAD) AS IMPORTE, VENTAS.SUCURSAL FROM VENTAS INNER JOIN VENTAS_DETALLE ON VENTAS.ID_VENTA = VENTAS_DETALLE.ID_VENTA INNER JOIN ALMACEN3 ON VENTAS_DETALLE.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO INNER JOIN CLIENTE ON VENTAS.ID_CLIENTE = CLIENTE.ID_CLIENTE WHERE (VENTAS.FECHA BETWEEN '01/01/2011' AND '31/12/2011 23:59:59.997') AND (VENTAS.SUCURSAL = '" & Combo1.Text & "') AND"
                End If
            End If
            iTotal = "0"
            FechaDel = DTPicker1.Value
            FechaAl = DTPicker2.Value
            If Check7.Value = 1 Then
                sBuscar = sEncab & " VENTAS.FACTURADO = 1 AND"
            Else
                sBuscar = sEncab & "  AND (VENTAS.FACTURADO IN (0, 1))"
            End If
            If sIdUsuario <> "" Then
                If Option7.Value Then
                    sEncab = sEncab & " VENTAS.ID_USUARIO  = '" & sIdUsuario & "' AND "
                Else
                    sEncab = sEncab & " CLIENTE.ID_AGENTE  = '" & sIdUsuario & "' AND "
                End If
            End If
            Text6.Text = sEncab
            If Check1.Value = 1 Then
                If Option4.Value Then
                    sBuscar = sEncab & " (VENTAS_DETALLE.ID_PRODUCTO LIKE '%REC')" & sTipo
                Else
                    sBuscar = sEncab & " (VENTAS_DETALLE.ID_PRODUCTO LIKE '%REC')" & sTipo & "GROUP BY  VENTAS_DETALLE.ID_PRODUCTO, VENTAS.SUCURSAL"
                End If
                cnn.CommandTimeout = 600
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    Set tLi = ListView1.ListItems.Add(, , "RECARGAS")
                    Do While Not tRs.EOF
                        If Option4.Value Then
                            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
                            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
                            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(3) = tRs.Fields("PRECIO_VENTA")
                            If Not IsNull(tRs.Fields("USUARIO")) Then tLi.SubItems(4) = tRs.Fields("USUARIO")
                            iTotal = CDbl(iTotal) + (CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("CANTIDAD")))
                        Else
                            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(1) = tRs.Fields("CANTIDAD")
                            If Not IsNull(tRs.Fields("IMPORTE")) Then tLi.SubItems(2) = tRs.Fields("IMPORTE")
                            If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(3) = tRs.Fields("SUCURSAL")
                            iTotal = CDbl(iTotal) + CDbl(tRs.Fields("IMPORTE"))
                        End If
                        tRs.MoveNext
                    Loop
                    Text1.Text = iTotal
                End If
            End If
            iTotal = "0"
            If Check2.Value = 1 Then
                If Option4.Value Then
                    sBuscar = sEncab & " (VENTAS_DETALLE.ID_PRODUCTO LIKE '%REM')" & sTipo
                Else
                    sBuscar = sEncab & " (VENTAS_DETALLE.ID_PRODUCTO LIKE '%REM')" & sTipo & "GROUP BY  VENTAS_DETALLE.ID_PRODUCTO, VENTAS.SUCURSAL"
                End If
                cnn.CommandTimeout = 600
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    Set tLi = ListView1.ListItems.Add(, , "REMANOFACTURAS")
                    Do While Not tRs.EOF
                        If Option4.Value Then
                            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
                            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
                            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(3) = tRs.Fields("PRECIO_VENTA")
                            If Not IsNull(tRs.Fields("USUARIO")) Then tLi.SubItems(4) = tRs.Fields("USUARIO")
                            iTotal = CDbl(iTotal) + (CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("CANTIDAD")))
                        Else
                            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(1) = tRs.Fields("CANTIDAD")
                            If Not IsNull(tRs.Fields("IMPORTE")) Then tLi.SubItems(2) = tRs.Fields("IMPORTE")
                            If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(3) = tRs.Fields("SUCURSAL")
                            iTotal = CDbl(iTotal) + CDbl(tRs.Fields("IMPORTE"))
                        End If
                        tRs.MoveNext
                    Loop
                    Text2.Text = iTotal
                End If
            End If
            iTotal = "0"
            If Check3.Value = 1 Then
                If Option4.Value Then
                    sBuscar = sEncab & " ((VENTAS_DETALLE.ID_PRODUCTO LIKE '%COMAP') OR (VENTAS_DETALLE.ID_PRODUCTO LIKE '%COMGEN'))" & sTipo
                Else
                    sBuscar = sEncab & " ((VENTAS_DETALLE.ID_PRODUCTO LIKE '%COMAP') OR (VENTAS_DETALLE.ID_PRODUCTO LIKE '%COMGEN'))" & sTipo & "GROUP BY  VENTAS_DETALLE.ID_PRODUCTO, VENTAS.SUCURSAL"
                End If
                cnn.CommandTimeout = 600
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    Set tLi = ListView1.ListItems.Add(, , "COMPATIBLES")
                    Do While Not tRs.EOF
                        If Option4.Value Then
                            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
                            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
                            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(3) = tRs.Fields("PRECIO_VENTA")
                            If Not IsNull(tRs.Fields("USUARIO")) Then tLi.SubItems(4) = tRs.Fields("USUARIO")
                            iTotal = CDbl(iTotal) + (CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("CANTIDAD")))
                        Else
                            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(1) = tRs.Fields("CANTIDAD")
                            If Not IsNull(tRs.Fields("IMPORTE")) Then tLi.SubItems(2) = tRs.Fields("IMPORTE")
                            If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(3) = tRs.Fields("SUCURSAL")
                            iTotal = CDbl(iTotal) + CDbl(tRs.Fields("IMPORTE"))
                        End If
                        tRs.MoveNext
                    Loop
                    Text3.Text = iTotal
                End If
            End If
            iTotal = "0"
            If Check4.Value = 1 Then
                If Option4.Value Then
                    sBuscar = sEncab & " (VENTAS_DETALLE.ID_PRODUCTO LIKE '%CAMAP')" & sTipo
                Else
                    sBuscar = sEncab & " (VENTAS_DETALLE.ID_PRODUCTO LIKE '%CAMAP')" & sTipo & "GROUP BY  VENTAS_DETALLE.ID_PRODUCTO, VENTAS.SUCURSAL"
                End If
                cnn.CommandTimeout = 600
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    Set tLi = ListView1.ListItems.Add(, , "CAMBIOS")
                    Do While Not tRs.EOF
                        If Option4.Value Then
                            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
                            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
                            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(3) = tRs.Fields("PRECIO_VENTA")
                            If Not IsNull(tRs.Fields("USUARIO")) Then tLi.SubItems(4) = tRs.Fields("USUARIO")
                            iTotal = CDbl(iTotal) + (CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("CANTIDAD")))
                        Else
                            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(1) = tRs.Fields("CANTIDAD")
                            If Not IsNull(tRs.Fields("IMPORTE")) Then tLi.SubItems(2) = tRs.Fields("IMPORTE")
                            If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(3) = tRs.Fields("SUCURSAL")
                            iTotal = CDbl(iTotal) + CDbl(tRs.Fields("IMPORTE"))
                        End If
                        tRs.MoveNext
                    Loop
                    Text4.Text = iTotal
                End If
            End If
            iTotal = "0"
            If Check5.Value = 1 Then
                If Option4.Value Then
                    sBuscar = sEncab & " (VENTAS_DETALLE.ID_PRODUCTO NOT LIKE '%REC') AND (VENTAS_DETALLE.ID_PRODUCTO NOT LIKE '%REM') AND (VENTAS_DETALLE.ID_PRODUCTO NOT LIKE '%CAMAP') AND (VENTAS_DETALLE.ID_PRODUCTO NOT LIKE '%COMAP') AND (VENTAS_DETALLE.ID_PRODUCTO NOT LIKE '%COMGEN') AND (CLASIFICACION = 'ORIGINAL')" & sTipo
                Else
                    sBuscar = sEncab & " (VENTAS_DETALLE.ID_PRODUCTO NOT LIKE '%REC') AND (VENTAS_DETALLE.ID_PRODUCTO NOT LIKE '%REM') AND (VENTAS_DETALLE.ID_PRODUCTO NOT LIKE '%CAMAP') AND (VENTAS_DETALLE.ID_PRODUCTO NOT LIKE '%COMAP') AND (VENTAS_DETALLE.ID_PRODUCTO NOT LIKE '%COMGEN') AND (CLASIFICACION = 'ORIGINAL')" & sTipo & " GROUP BY  VENTAS_DETALLE.ID_PRODUCTO, VENTAS.SUCURSAL"
                End If
                cnn.CommandTimeout = 600
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    Set tLi = ListView1.ListItems.Add(, , "ORIGINALES")
                    Do While Not tRs.EOF
                        If Option4.Value Then
                            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
                            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
                            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(3) = tRs.Fields("PRECIO_VENTA")
                            If Not IsNull(tRs.Fields("USUARIO")) Then tLi.SubItems(4) = tRs.Fields("USUARIO")
                            iTotal = CDbl(iTotal) + (CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("CANTIDAD")))
                        Else
                            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(1) = tRs.Fields("CANTIDAD")
                            If Not IsNull(tRs.Fields("IMPORTE")) Then tLi.SubItems(2) = tRs.Fields("IMPORTE")
                            If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(3) = tRs.Fields("SUCURSAL")
                            iTotal = CDbl(iTotal) + CDbl(tRs.Fields("IMPORTE"))
                        End If
                        tRs.MoveNext
                    Loop
                    Text5.Text = iTotal
                End If
            End If
            iTotal = "0"
            If Check6.Value = 1 Then
                If Option4.Value Then
                    sBuscar = sEncab & " ALMACEN3.CLASIFICACION = 'SERVICIO' " & sTipo
                Else
                    sBuscar = sEncab & " ALMACEN3.CLASIFICACION = 'SERVICIO' " & sTipo & "GROUP BY  VENTAS_DETALLE.ID_PRODUCTO, VENTAS.SUCURSAL"
                End If
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    Set tLi = ListView1.ListItems.Add(, , "SERVICIOS")
                    Do While Not tRs.EOF
                        If Option4.Value Then
                            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
                            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
                            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(3) = tRs.Fields("PRECIO_VENTA")
                            If Not IsNull(tRs.Fields("USUARIO")) Then tLi.SubItems(4) = tRs.Fields("USUARIO")
                            iTotal = CDbl(iTotal) + (CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("CANTIDAD")))
                        Else
                            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(1) = tRs.Fields("CANTIDAD")
                            If Not IsNull(tRs.Fields("IMPORTE")) Then tLi.SubItems(2) = tRs.Fields("IMPORTE")
                            If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(3) = tRs.Fields("SUCURSAL")
                            iTotal = CDbl(iTotal) + CDbl(tRs.Fields("IMPORTE"))
                        End If
                        tRs.MoveNext
                    Loop
                    Text5.Text = iTotal
                End If
            End If
            
        End If
    Else
        MsgBox "Debe seleccionar una sucursal", vbExclamation, "SACC"
    End If
End Sub
Private Sub Combo1_DropDown()
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Combo1.Clear
    Combo1.AddItem "<TODAS>"
    sBuscar = "SELECT NOMBRE FROM SUCURSALES GROUP BY NOMBRE ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Combo1_LostFocus()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    sBuscar = "SELECT ID_USUARIO FROM VsAgentesDeVentas WHERE NOMBRE = '" & Combo1.Text & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        ClvAgente = tRs.Fields("ID_USUARIO")
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    DTPicker1.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
           "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    sBuscar = "SELECT NOMBRE FROM USUARIOS WHERE PUESTO LIKE '%VENTA%' AND ESTADO = 'A'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo2.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image10_Click()
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Dim foo As Integer
    Me.CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
    Me.CommonDialog1.ShowSave
    Ruta = Me.CommonDialog1.FileName
    If SSTab1.Tab = 0 Then
        If ListView1.ListItems.Count > 0 Then
            If Ruta <> "" Then
                NumColum = ListView1.ColumnHeaders.Count
                For Con = 1 To ListView1.ColumnHeaders.Count
                    StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
                For Con = 1 To ListView1.ListItems.Count
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                    For Con2 = 1 To NumColum - 1
                        StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                Next
                'archivo TXT
                foo = FreeFile
                Open Ruta For Output As #foo
                    Print #foo, StrCopi
                Close #foo
            End If
        Else
            If ListView2.ListItems.Count > 0 Then
                If Ruta <> "" Then
                    NumColum = ListView2.ColumnHeaders.Count
                    For Con = 1 To ListView2.ColumnHeaders.Count
                        StrCopi = StrCopi & ListView2.ColumnHeaders(Con).Text & Chr(9)
                    Next
                    StrCopi = StrCopi & Chr(13)
                    For Con = 1 To ListView2.ListItems.Count
                        StrCopi = StrCopi & ListView2.ListItems.Item(Con) & Chr(9)
                        For Con2 = 1 To NumColum - 1
                            StrCopi = StrCopi & ListView2.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                        Next
                        StrCopi = StrCopi & Chr(13)
                    Next
                    'archivo TXT
                    foo = FreeFile
                    Open Ruta For Output As #foo
                        Print #foo, StrCopi
                    Close #foo
                End If
            End If
        End If
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
Private Sub Image26_Click()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    Dim Suma As String
    Dim Total As String
    ConPag = 1
    Total = "0"
    Suma = "0"
    'text1.text = no_orden
    'nvomen.Text1(0).Text = id_usuario
    If Not (ListView1.ListItems.Count = 0) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\VentasSucursalEspecie.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        ' asi se agregan los logos... solo te falto poner un control IMAGE1 para cargar la imagen en el
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs1 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 175, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 175, tRs1.Fields("DIRECCION"), "F3", 8, hLeft
        oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs1.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 70, 205, 100, 175, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 80, 205, 100, 175, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 30, 380, 20, 250, "Del: " & FechaDel & " Al: " & FechaAl, "F3", 8, hCenter
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE VENTAS DE SUCURSAL " & Sucursal, "F3", 8, hCenter
        Posi = 120
        If Option4.Value Then
            oDoc.WTextBox Posi, 5, 20, 200, "Producto", "F2", 8, hCenter
            oDoc.WTextBox Posi, 205, 20, 70, "Cantidad", "F2", 8, hCenter
            oDoc.WTextBox Posi, 280, 20, 80, "Precio", "F2", 8, hCenter
            oDoc.WTextBox Posi, 365, 20, 80, "Importe", "F2", 8, hCenter
        Else
            oDoc.WTextBox Posi, 5, 20, 200, "Producto", "F2", 8, hCenter
            oDoc.WTextBox Posi, 205, 20, 70, "Cantidad", "F2", 8, hCenter
            oDoc.WTextBox Posi, 280, 20, 80, "Importe", "F2", 8, hCenter
            oDoc.WTextBox Posi, 365, 20, 80, "Sucursal", "F2", 8, hCenter
        End If
        Posi = Posi + 12
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        For Cont = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(Cont) = "RECARGAS" Or ListView1.ListItems(Cont) = "REMANOFACTURAS" Or ListView1.ListItems(Cont) = "COMPATIBLES" Or ListView1.ListItems(Cont) = "CAMBIOS" Or ListView1.ListItems(Cont) = "ORIGINALES" Or ListView1.ListItems(Cont) = "SERVICIOS" Then
                If ListView1.ListItems(Cont) = "REMANOFACTURAS" Or ListView1.ListItems(Cont) = "COMPATIBLES" Or ListView1.ListItems(Cont) = "CAMBIOS" Or ListView1.ListItems(Cont) = "ORIGINALES" Or ListView1.ListItems(Cont) = "SERVICIOS" Then
                    oDoc.WTextBox Posi, 300, 20, 150, "Sumatoria: " & Format(Suma, "###,###,##0.00"), "F2", 8, hRight
                    Posi = Posi + 16
                End If
                Total = CDbl(Total) + CDbl(Suma)
                Suma = "0"
                oDoc.WTextBox Posi, 5, 20, 200, ListView1.ListItems(Cont), "F2", 7, hLeft
            Else
                If Option4.Value Then
                    oDoc.WTextBox Posi, 45, 20, 160, ListView1.ListItems(Cont).SubItems(1), "F3", 7, hLeft
                    oDoc.WTextBox Posi, 205, 20, 70, Format(ListView1.ListItems(Cont).SubItems(2), "###,###,##0.00"), "F3", 7, hRight
                    oDoc.WTextBox Posi, 280, 20, 80, Format(ListView1.ListItems(Cont).SubItems(3), "###,###,##0.00"), "F3", 7, hRight
                    oDoc.WTextBox Posi, 365, 20, 80, Format(ListView1.ListItems(Cont).SubItems(4), "###,###,##0.00"), "F3", 7, hRight
                    If ListView1.ListItems(Cont).SubItems(3) <> "" Then
                        Suma = CDbl(Suma) + (CDbl(ListView1.ListItems(Cont).SubItems(3)) * CDbl(ListView1.ListItems(Cont).SubItems(2)))
                    End If
                Else
                    oDoc.WTextBox Posi, 45, 20, 160, ListView1.ListItems(Cont), "F3", 7, hLeft
                    oDoc.WTextBox Posi, 205, 20, 70, Format(ListView1.ListItems(Cont).SubItems(1), "###,###,##0.00"), "F3", 7, hRight
                    oDoc.WTextBox Posi, 280, 20, 80, Format(ListView1.ListItems(Cont).SubItems(2), "###,###,##0.00"), "F3", 7, hRight
                    oDoc.WTextBox Posi, 365, 20, 80, Format(ListView1.ListItems(Cont).SubItems(3), "###,###,##0.00"), "F3", 7, hRight
                    If ListView1.ListItems(Cont).SubItems(2) <> "" Then
                        Suma = CDbl(Suma) + (CDbl(ListView1.ListItems(Cont).SubItems(2)))
                    End If
                End If
            End If
            Posi = Posi + 12
            If Posi >= 650 Then
                oDoc.NewPage A4_Vertical
                oDoc.WImage 70, 40, 43, 161, "Logo"
                sBuscar = "SELECT * FROM EMPRESA"
                Set tRs4 = cnn.Execute(sBuscar)
                oDoc.WTextBox 40, 205, 100, 175, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
                oDoc.WTextBox 60, 224, 100, 175, tRs1.Fields("DIRECCION"), "F3", 8, hLeft
                oDoc.WTextBox 60, 328, 100, 175, "Col." & tRs1.Fields("COLONIA"), "F3", 8, hLeft
                oDoc.WTextBox 70, 205, 100, 175, tRs1.Fields("ESTADO") & "," & tRs1.Fields("CD"), "F3", 8, hCenter
                oDoc.WTextBox 80, 205, 100, 175, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
                oDoc.WTextBox 30, 380, 20, 250, "Del: " & FechaDel & " Al: " & FechaAl, "F3", 8, hCenter
                ' ENCABEZADO DEL DETALLE
                oDoc.WTextBox 100, 205, 100, 175, "REPORTE DE VENTAS DE SUCURSAL " & Sucursal, "F3", 8, hCenter
                Posi = 120
                If Option4.Value Then
                    oDoc.WTextBox Posi, 5, 20, 200, "Producto", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 205, 20, 70, "Cantidad", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 280, 20, 80, "Precio", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 365, 20, 80, "Importe", "F2", 8, hCenter
                Else
                    oDoc.WTextBox Posi, 5, 20, 200, "Producto", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 205, 20, 70, "Cantidad", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 280, 20, 80, "Importe", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 365, 20, 80, "Sucursal", "F2", 8, hCenter
                End If
                Posi = Posi + 12
                ' Linea
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, Posi
                oDoc.WLineTo 580, Posi
                oDoc.LineStroke
                Posi = Posi + 6
            End If
        Next Cont
        oDoc.WTextBox Posi, 300, 20, 150, "Sumatoria: " & Format(Suma, "###,###,##0.00"), "F2", 8, hRight
        Posi = Posi + 16
        ' Linea
        Posi = Posi + 6
        Total = CDbl(Total) + CDbl(Suma)
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
         Posi = Posi + 16
        ' TEXTO ABAJO
        oDoc.WTextBox Posi, 300, 20, 150, "Total:" & Format(Total, "###,###,##0.00"), "F2", 8, hRight
        Posi = Posi + 16
        oDoc.WTextBox Posi, 205, 100, 175, "COMENTARIOS", "F3", 8, hCenter
        Posi = Posi + 20
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 16
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 16
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
