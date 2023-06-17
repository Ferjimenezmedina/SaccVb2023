VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2B26B39A-53D1-4401-B64E-1B727C1D2B68}#9.0#0"; "ADMGráficos.ocx"
Begin VB.Form FrmRendimiento 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rendimiento por Depatamentos"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8280
      TabIndex        =   7
      Top             =   5280
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
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRendimiento.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRendimiento.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "FrmRendimiento.frx":23EC
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
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListView5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ListView4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ListView3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ListView2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DTPicker1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ListView1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Combo1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Gráficas De Ventas"
      TabPicture(1)   =   "FrmRendimiento.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ADMGraf1"
      Tab(1).Control(1)=   "ADMGraf2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Gráficas de Compras"
      TabPicture(2)   =   "FrmRendimiento.frx":2424
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ADMGraf4"
      Tab(2).Control(1)=   "ADMGraf3"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Gráficas de Almacén"
      TabPicture(3)   =   "FrmRendimiento.frx":2440
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ADMGraf5"
      Tab(3).ControlCount=   1
      Begin ADMGráficos.ADMGraf ADMGraf1 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4895
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mostrar_Leyenda =   0   'False
         Color_Fondo     =   -2147483643
         Color_Barra1    =   8388608
         Color_Barra2    =   0
         Gráfico_Barras  =   -1  'True
         Color_Texto     =   -2147483640
         Mostrar_Media   =   0   'False
         Color_Media     =   0
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3000
         TabIndex        =   1
         Text            =   "<TODAS>"
         Top             =   600
         Width           =   2415
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2566
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command5 
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
         Left            =   5520
         Picture         =   "FrmRendimiento.frx":245C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   840
         TabIndex        =   0
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17104897
         CurrentDate     =   40631
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         Top             =   3000
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2566
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1455
         Left            =   4080
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2566
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   1455
         Left            =   4080
         TabIndex        =   5
         Top             =   3000
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2566
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ADMGráficos.ADMGraf ADMGraf2 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   17
         Top             =   3360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5106
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mostrar_Leyenda =   0   'False
         Color_Fondo     =   -2147483643
         Color_Barra1    =   8388608
         Color_Barra2    =   0
         Gráfico_Barras  =   -1  'True
         Color_Texto     =   -2147483640
         Mostrar_Media   =   0   'False
         Color_Media     =   0
      End
      Begin ADMGráficos.ADMGraf ADMGraf3 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4895
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mostrar_Leyenda =   0   'False
         Color_Fondo     =   -2147483643
         Color_Barra1    =   8388608
         Color_Barra2    =   0
         Gráfico_Barras  =   -1  'True
         Color_Texto     =   -2147483640
         Mostrar_Media   =   0   'False
         Color_Media     =   0
      End
      Begin ADMGráficos.ADMGraf ADMGraf4 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   19
         Top             =   3360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5106
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mostrar_Leyenda =   0   'False
         Color_Fondo     =   -2147483643
         Color_Barra1    =   8388608
         Color_Barra2    =   0
         Gráfico_Barras  =   -1  'True
         Color_Texto     =   -2147483640
         Mostrar_Media   =   0   'False
         Color_Media     =   0
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   1455
         Left            =   0
         TabIndex        =   20
         Top             =   4680
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2566
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ADMGráficos.ADMGraf ADMGraf5 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4895
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mostrar_Leyenda =   0   'False
         Color_Fondo     =   -2147483643
         Color_Barra1    =   8388608
         Color_Barra2    =   0
         Gráfico_Barras  =   -1  'True
         Color_Texto     =   -2147483640
         Mostrar_Media   =   0   'False
         Color_Media     =   0
      End
      Begin VB.Label Label7 
         Caption         =   "Entradas a Almacén Por Usuario"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   4440
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "Sucursal :"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Ordenes De Compra Rápidas Por Usuario"
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "Ordenes De Compra Por Usuario"
         Height          =   255
         Left            =   4200
         TabIndex        =   12
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Ventas Programadas Por Usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Ventas Por Usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmRendimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Command5_Click()
    Actualiza
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    DTPicker1.Value = Date
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' GROUP BY NOMBRE ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Combo1.AddItem "<TODAS>"
        Do While Not (tRs.EOF)
            If Trim(tRs.Fields("NOMBRE")) <> "" Then Combo1.AddItem (tRs.Fields("NOMBRE"))
            tRs.MoveNext
        Loop
    End If
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Vendedor", 2500
        .ColumnHeaders.Add , , "Cant. Notas", 1500
        .ColumnHeaders.Add , , "Importe", 2000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Vendedor", 2500
        .ColumnHeaders.Add , , "Cant. Notas", 1500
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Comprador", 2500
        .ColumnHeaders.Add , , "Cant. Ordenes", 1500
        .ColumnHeaders.Add , , "Importe", 2000
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Comprador", 2500
        .ColumnHeaders.Add , , "Cant. Ordenes", 1500
        .ColumnHeaders.Add , , "Importe", 2000
    End With
    With ListView5
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Almacenista", 2500
        .ColumnHeaders.Add , , "Cant. Entradas", 1500
        .ColumnHeaders.Add , , "Importe", 2000
    End With
    Actualiza
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Actualiza()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    If Combo1.Text <> "<TODAS>" And Combo1.Text <> "" Then
        sBuscar = "SELECT COUNT(VENTAS.ID_VENTA) AS TotVentas, SUM(VENTAS.TOTAL) AS TotImporte, USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS VENDEDOR FROM VENTAS INNER JOIN USUARIOS ON VENTAS.ID_USUARIO = USUARIOS.ID_USUARIO WHERE (VENTAS.FECHA = '" & DTPicker1.Value & "') AND VENTAS.SUCURSAL = '" & Combo1.Text & "' GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    Else
        sBuscar = "SELECT COUNT(VENTAS.ID_VENTA) AS TotVentas, SUM(VENTAS.TOTAL) AS TotImporte, USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS VENDEDOR FROM VENTAS INNER JOIN USUARIOS ON VENTAS.ID_USUARIO = USUARIOS.ID_USUARIO WHERE (VENTAS.FECHA = '" & DTPicker1.Value & "') GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("VENDEDOR"))
            tLi.SubItems(1) = tRs.Fields("TotVentas")
            tLi.SubItems(2) = tRs.Fields("TotImporte")
            tRs.MoveNext
        Loop
    End If
    ListView2.ListItems.Clear
    If Combo1.Text <> "<TODAS>" And Combo1.Text <> "" Then
        sBuscar = "SELECT USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS VENDEDOR, COUNT(PED_CLIEN.NO_PEDIDO) AS TotVentas FROM PED_CLIEN INNER JOIN USUARIOS ON PED_CLIEN.USUARIO = USUARIOS.ID_USUARIO INNER JOIN SUCURSALES ON USUARIOS.ID_SUCURSAL = SUCURSALES.ID_SUCURSAL WHERE (PED_CLIEN.FECHA_CAPTURA = '" & DTPicker1.Value & "') AND (SUCURSALES.NOMBRE = '" & Combo1.Text & "') GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    Else
        sBuscar = "SELECT USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS VENDEDOR, COUNT(PED_CLIEN.NO_PEDIDO) AS TotVentas FROM PED_CLIEN INNER JOIN USUARIOS ON PED_CLIEN.USUARIO = USUARIOS.ID_USUARIO INNER JOIN SUCURSALES ON USUARIOS.ID_SUCURSAL = SUCURSALES.ID_SUCURSAL WHERE (PED_CLIEN.FECHA_CAPTURA = '" & DTPicker1.Value & "') GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("VENDEDOR"))
            tLi.SubItems(1) = tRs.Fields("TotVentas")
            tRs.MoveNext
        Loop
    End If
    ListView3.ListItems.Clear
    sBuscar = "SELECT USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS Comprador, COUNT(ORDEN_COMPRA.ID_ORDEN_COMPRA) AS TotCompras, SUM(dbo.ORDEN_COMPRA.TOTAL) AS TotImporte FROM ORDEN_COMPRA INNER JOIN USUARIOS ON ORDEN_COMPRA.ID_USUARIO = USUARIOS.ID_USUARIO WHERE (ORDEN_COMPRA.FECHA = '" & DTPicker1.Value & "') GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView3.ListItems.Add(, , tRs.Fields("Comprador"))
            tLi.SubItems(1) = tRs.Fields("TotCompras")
            tLi.SubItems(2) = tRs.Fields("TotImporte")
            tRs.MoveNext
        Loop
    End If
    ListView4.ListItems.Clear
    sBuscar = "SELECT USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS Comprador, SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TotImporte, COUNT(ORDEN_RAPIDA.ID_ORDEN_RAPIDA) As TotCompras FROM ORDEN_RAPIDA INNER JOIN USUARIOS ON ORDEN_RAPIDA.ID_USUARIO = USUARIOS.ID_USUARIO INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA WHERE (ORDEN_RAPIDA.FECHA = '" & DTPicker1.Value & "') GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView4.ListItems.Add(, , tRs.Fields("Comprador"))
            tLi.SubItems(1) = tRs.Fields("TotCompras")
            tLi.SubItems(2) = tRs.Fields("TotImporte")
            tRs.MoveNext
        Loop
    End If
    ListView5.ListItems.Clear
    sBuscar = "SELECT USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS NOMBRE, COUNT(ENTRADAS.ID_ENTRADA) AS ENTRADAS, SUM(ENTRADAS.Total) As Total FROM ENTRADAS INNER JOIN USUARIOS ON ENTRADAS.ID_USUARIO = USUARIOS.ID_USUARIO WHERE (ENTRADAS.FECHA = '" & DTPicker1.Value & "') GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView5.ListItems.Add(, , tRs.Fields("NOMBRE"))
            tLi.SubItems(1) = tRs.Fields("ENTRADAS")
            tLi.SubItems(2) = tRs.Fields("Total")
            tRs.MoveNext
        Loop
    End If
    GRAFICA
End Sub
Private Sub GRAFICA()
    Dim n As Integer
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    n = n + 1
    ADMGraf1.Gráfico_Barras = True
    ADMGraf1.Limpiar
    If Combo1.Text <> "<TODAS>" And Combo1.Text <> "" Then
        sBuscar = "SELECT COUNT(VENTAS.ID_VENTA) AS TotVentas, SUM(VENTAS.TOTAL) AS TotImporte, USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS VENDEDOR FROM VENTAS INNER JOIN USUARIOS ON VENTAS.ID_USUARIO = USUARIOS.ID_USUARIO WHERE (VENTAS.FECHA = '" & DTPicker1.Value & "') AND VENTAS.SUCURSAL = '" & Combo1.Text & "' GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    Else
        sBuscar = "SELECT COUNT(VENTAS.ID_VENTA) AS TotVentas, SUM(VENTAS.TOTAL) AS TotImporte, USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS VENDEDOR FROM VENTAS INNER JOIN USUARIOS ON VENTAS.ID_USUARIO = USUARIOS.ID_USUARIO WHERE (VENTAS.FECHA = '" & DTPicker1.Value & "') GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    End If
    Set tRs = cnn.Execute(sBuscar)
    ADMGraf1.Título = "Ventas Por Usuario"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            ADMGraf1.Introducir Mid(tRs.Fields("VENDEDOR"), 1, 10), CSng(CDbl(tRs.Fields("TotVentas"))), CLng(CDbl(tRs.Fields("TotVentas")) * 16000000), QBColor(15)
            tRs.MoveNext
        Loop
        ADMGraf1.Dibujar
    End If
    ADMGraf2.Gráfico_Barras = True
    ADMGraf2.Limpiar
    If Combo1.Text <> "<TODAS>" And Combo1.Text <> "" Then
        sBuscar = "SELECT USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS VENDEDOR, COUNT(PED_CLIEN.NO_PEDIDO) AS TotVentas FROM PED_CLIEN INNER JOIN USUARIOS ON PED_CLIEN.USUARIO = USUARIOS.ID_USUARIO INNER JOIN SUCURSALES ON USUARIOS.ID_SUCURSAL = SUCURSALES.ID_SUCURSAL WHERE (PED_CLIEN.FECHA_CAPTURA = '" & DTPicker1.Value & "') AND (SUCURSALES.NOMBRE = '" & Combo1.Text & "') GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    Else
        sBuscar = "SELECT USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS VENDEDOR, COUNT(PED_CLIEN.NO_PEDIDO) AS TotVentas FROM PED_CLIEN INNER JOIN USUARIOS ON PED_CLIEN.USUARIO = USUARIOS.ID_USUARIO INNER JOIN SUCURSALES ON USUARIOS.ID_SUCURSAL = SUCURSALES.ID_SUCURSAL WHERE (PED_CLIEN.FECHA_CAPTURA = '" & DTPicker1.Value & "') GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    End If
    Set tRs = cnn.Execute(sBuscar)
    ADMGraf2.Título = "Ventas Programadas Por Usuario"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            ADMGraf2.Introducir Mid(tRs.Fields("VENDEDOR"), 1, 10), CSng(CDbl(tRs.Fields("TotVentas"))), CLng(CDbl(tRs.Fields("TotVentas")) * 16000000), QBColor(15)
            tRs.MoveNext
        Loop
        ADMGraf2.Dibujar
    End If
    ADMGraf3.Gráfico_Barras = True
    ADMGraf3.Limpiar
    sBuscar = "SELECT USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS Comprador, COUNT(ORDEN_COMPRA.ID_ORDEN_COMPRA) AS TotCompras, SUM(dbo.ORDEN_COMPRA.TOTAL) AS TotImporte FROM ORDEN_COMPRA INNER JOIN USUARIOS ON ORDEN_COMPRA.ID_USUARIO = USUARIOS.ID_USUARIO WHERE (ORDEN_COMPRA.FECHA = '" & DTPicker1.Value & "') GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    Set tRs = cnn.Execute(sBuscar)
    ADMGraf3.Título = "Ordenes De Compra Por Usuario"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            ADMGraf3.Introducir Mid(tRs.Fields("Comprador"), 1, 10), CSng(CDbl(tRs.Fields("TotCompras"))), CLng(CDbl(tRs.Fields("TotCompras")) * 16000000), QBColor(15)
            tRs.MoveNext
        Loop
        ADMGraf3.Dibujar
    End If
    ADMGraf4.Gráfico_Barras = True
    ADMGraf4.Limpiar
    sBuscar = "SELECT USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS Comprador, SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TotImporte, COUNT(ORDEN_RAPIDA.ID_ORDEN_RAPIDA) As TotCompras FROM ORDEN_RAPIDA INNER JOIN USUARIOS ON ORDEN_RAPIDA.ID_USUARIO = USUARIOS.ID_USUARIO INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA WHERE (ORDEN_RAPIDA.FECHA = '" & DTPicker1.Value & "') GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    Set tRs = cnn.Execute(sBuscar)
    ADMGraf4.Título = "Ordenes De Compra Rápidas Por Usuario"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            ADMGraf4.Introducir Mid(tRs.Fields("Comprador"), 1, 10), CSng(CDbl(tRs.Fields("TotCompras"))), CLng(CDbl(tRs.Fields("TotCompras")) * 16000000), QBColor(15)
            tRs.MoveNext
        Loop
        ADMGraf4.Dibujar
    End If
    ADMGraf5.Gráfico_Barras = True
    ADMGraf5.Limpiar
    sBuscar = "SELECT USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS NOMBRE, COUNT(ENTRADAS.ID_ENTRADA) AS ENTRADAS, SUM(ENTRADAS.Total) As Total FROM ENTRADAS INNER JOIN USUARIOS ON ENTRADAS.ID_USUARIO = USUARIOS.ID_USUARIO WHERE (ENTRADAS.FECHA = '" & DTPicker1.Value & "') GROUP BY USUARIOS.NOMBRE, USUARIOS.APELLIDOS"
    Set tRs = cnn.Execute(sBuscar)
    ADMGraf5.Título = "Entradas a Almacén"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            ADMGraf5.Introducir Mid(tRs.Fields("NOMBRE"), 1, 10), CSng(CDbl(tRs.Fields("ENTRADAS"))), CLng(CDbl(tRs.Fields("ENTRADAS")) * 16000000), QBColor(15)
            tRs.MoveNext
        Loop
        ADMGraf5.Dibujar
    End If
End Sub
