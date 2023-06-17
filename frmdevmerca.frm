VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmdevmerca 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devolucion de Mercancia a Proveedores"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleMode       =   0  'User
   ScaleWidth      =   10791.67
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10560
      TabIndex        =   27
      Top             =   7080
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmdevmerca.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmdevmerca.frx":030A
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
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10320
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmdevmerca.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label13"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ListView1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Option2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text5"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Option1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text4"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text3"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Command1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Command2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Command3"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text6"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text7"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Combo1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text8"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text9"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text10"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "ListView2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text11"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text12"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "ListView3"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "ListView4"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).ControlCount=   36
      Begin MSComctlLib.ListView ListView4 
         Height          =   2055
         Left            =   120
         TabIndex        =   40
         Top             =   4680
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1935
         Left            =   120
         TabIndex        =   38
         Top             =   2520
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3413
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
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   960
         TabIndex        =   37
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1200
         TabIndex        =   35
         Top             =   600
         Width           =   3135
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1095
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   1931
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
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   9000
         TabIndex        =   32
         Top             =   6720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   8880
         TabIndex        =   12
         Top             =   7200
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   9000
         TabIndex        =   11
         Top             =   7680
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4800
         TabIndex        =   15
         Top             =   7200
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         TabIndex        =   7
         Top             =   6600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   1560
         TabIndex        =   13
         Top             =   7680
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
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
         Left            =   1680
         Picture         =   "frmdevmerca.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   7680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
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
         Height          =   375
         Left            =   3840
         Picture         =   "frmdevmerca.frx":4DDA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7560
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "frmdevmerca.frx":77AC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6960
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   8520
         TabIndex        =   6
         Top             =   5880
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   7440
         TabIndex        =   8
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   7440
         TabIndex        =   9
         Top             =   7200
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   7920
         TabIndex        =   10
         Top             =   7680
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Consultar Entrada"
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   6840
         Width           =   135
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   6960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Consultar Devoluciones."
         Height          =   375
         Left            =   4800
         TabIndex        =   20
         Top             =   6840
         Width           =   255
      End
      Begin VB.Frame Frame1 
         Caption         =   "RANGO"
         Height          =   495
         Left            =   6840
         TabIndex        =   16
         Top             =   6240
         Width           =   135
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   255
            Left            =   600
            TabIndex        =   5
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   393216
            Format          =   69337089
            CurrentDate     =   39651
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   255
            Left            =   600
            TabIndex        =   4
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   393216
            Format          =   69337089
            CurrentDate     =   39600
         End
         Begin VB.Label Label7 
            Caption         =   "Al"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "De"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   615
         Left            =   720
         TabIndex        =   19
         Top             =   7080
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         BorderStyle     =   4  'Dash-Dot
         BorderWidth     =   3
         X1              =   10200
         X2              =   5640
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label13 
         Caption         =   "  Ordenes :"
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
         TabIndex        =   39
         Top             =   2280
         Width           =   4095
      End
      Begin VB.Label Label12 
         Caption         =   "Proveedor :"
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
         TabIndex        =   36
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "                      DATOS DEL PROVEEDOR"
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
         TabIndex        =   34
         Top             =   120
         Width           =   4095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         BorderStyle     =   4  'Dash-Dot
         BorderWidth     =   3
         X1              =   5640
         X2              =   5640
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Label Label10 
         Caption         =   "TOTAL_DEVUELTO :"
         Height          =   255
         Left            =   8640
         TabIndex        =   31
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "PRECIO:"
         Height          =   375
         Left            =   8880
         TabIndex        =   30
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "MOTIVO ::"
         Height          =   255
         Left            =   8040
         TabIndex        =   29
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "NUM_ORDEN"
         Height          =   255
         Left            =   7440
         TabIndex        =   25
         Top             =   6720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "NUM_ENTRADA"
         Height          =   255
         Left            =   7920
         TabIndex        =   24
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "PRODUCTO DEV"
         Height          =   255
         Left            =   6480
         TabIndex        =   23
         Top             =   7680
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "CANTIDAD A DEV."
         Height          =   255
         Left            =   7080
         TabIndex        =   22
         Top             =   7320
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   6000
         TabIndex        =   21
         Top             =   7440
         Visible         =   0   'False
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmdevmerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
 Dim StrRep As String
Private Sub cmdBuscar_Click()
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
   
    ListView1.ListItems.Clear
    
     'temporal_abonos
     sBuscar = "SELECT ID_VENTA,FOLIO,FECHA,NOMBRE,BANCO,FECHAABONO,TOTAL,CANT_ABONO,NO_CHEQUE,DEUDA,PAGADA,LIMITE_CREDITO,TOTAL_COMPRA FROM temporal_abonos WHERE FOlIO NOT IN ('CANCELADO') AND NOMBRE LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " ' ORDER BY NOMBRE,FECHA ASC "
     ' sBuscar = "SELECT ID_VENTA,FOLIO,FECHA,NOMBRE,NO_CHEQUE,BANCO,SUM(CANT_ABONO)AS CREDITO_DISPO,LIMITE_CREDITO, SUM(DEUDA)AS DEUDA_ACTUAL FROM VsRepAbonos WHERE ID_CLIENTE = " & Text1.Text & "%' AND FECHA BETWEEN'" & DTPicker1.Value & " ' AND '" & DTPicker2.Value & "' ORDER BY FECHA"
 StrRep = sBuscar
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
            tLi.SubItems(1) = tRs.Fields("FOLIO")
            tLi.SubItems(2) = tRs.Fields("FECHA")
             tLi.SubItems(3) = tRs.Fields("NOMBRE")
            'tLi.SubItems(4) = tRs.Fields("BANCO")
            If Not IsNull(tRs.Fields("BANCO")) Then tLi.SubItems(4) = tRs.Fields("BANCO")
            If Not IsNull(tRs.Fields("CANT_ABONO")) Then tLi.SubItems(5) = tRs.Fields("CANT_ABONO")
            'tLi.SubItems(4) = tRs.Fields(Not ("CANT_ABONO"))
            tLi.SubItems(6) = tRs.Fields("DEUDA")
             tLi.SubItems(7) = tRs.Fields("PAGADA")
           tLi.SubItems(8) = tRs.Fields("LIMITE_CREDITO")
            tLi.SubItems(9) = tRs.Fields("TOTAL")
            tRs.MoveNext
        Loop
      
    End If
  
End Sub

Private Sub Command1_Click()
Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
   
    ListView1.ListItems.Clear
    
     'temporal_abonos
     sBuscar = "SELECT * FROM VSDEVOLUCION WHERE NOMBRE LIKE '%" & Text5.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " ' ORDER BY FECHA DESC "
     ' sBuscar = "SELECT ID_VENTA,FOLIO,FECHA,NOMBRE,NO_CHEQUE,BANCO,SUM(CANT_ABONO)AS CREDITO_DISPO,LIMITE_CREDITO, SUM(DEUDA)AS DEUDA_ACTUAL FROM VsRepAbonos WHERE ID_CLIENTE = " & Text1.Text & "%' AND FECHA BETWEEN'" & DTPicker1.Value & " ' AND '" & DTPicker2.Value & "' ORDER BY FECHA"
 StrRep = sBuscar
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR"))
            tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tLi.SubItems(2) = tRs.Fields("ID_ENTRADA")
            If Not IsNull(tRs.Fields("NUM_ORDEN")) Then tLi.SubItems(3) = tRs.Fields("NUM_ORDEN")
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(4) = tRs.Fields("ID_PRODUCTO")
            tLi.SubItems(5) = tRs.Fields("CANTIDAD")
            tLi.SubItems(6) = tRs.Fields("FECHA")
            tLi.SubItems(7) = tRs.Fields("PRECIO")
            tLi.SubItems(8) = tRs.Fields("FACTURA")
            
             tLi.SubItems(9) = (CDbl(tLi.SubItems(5)) * CDbl(tLi.SubItems(7)))
             
            tRs.MoveNext
        Loop
      
    End If
  
End Sub

Private Sub Command2_Click()
Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    sBuscar = "INSERT INTO DEVOLUCION (NUM_ORDEN,ID_ENTRADA,ID_PRODUCTO,CANTIDAD,PROVEEDOR,MOTIVO,SUCURSAL,FECHA,PRECIO,FACT_PRO,TOT_DEV) VALUES ('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "', '" & Text7.Text & "', '" & Text6.Text & "','" & Combo1.Text & "','" & DTPicker2.Value & " ', '" & Text8.Text & "', '" & Text10.Text & "', '" & Text9.Text & "');"
    cnn.Execute (sBuscar)
    StrRep = sBuscar
    sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE SUCURSAL = '" & Combo1.Text & "'  AND ID_PRODUCTO = '" & Text3.Text & "'"
    Set tRs = cnn.Execute(sBuscar)
    sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs.Fields("CANTIDAD")) - CDbl(Text4.Text) & " WHERE ID_PRODUCTO = '" & Text3.Text & "' AND SUCURSAL ='" & Combo1.Text & "'"
    cnn.Execute (sBuscar)
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    MsgBox "SU INFORMACION YA FUE  ALMACENADA"
    Imprimir
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id_Proveedor", 1000
        .ColumnHeaders.Add , , "Nombre", 1200
        .ColumnHeaders.Add , , "Id__entrada", 1200
        .ColumnHeaders.Add , , "Num_Orden", 1200
         .ColumnHeaders.Add , , "Id_Producto", 1500
          .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Fecha", 1000
         .ColumnHeaders.Add , , "Precio", 1000
         .ColumnHeaders.Add , , "Factura", 1000
         .ColumnHeaders.Add , , "Total_nota", 1000
              
        End With
        With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id_Proveedor", 1000
        .ColumnHeaders.Add , , "Nombre", 3500
       
              
        End With
        With ListView3
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id_Orden", 500
        .ColumnHeaders.Add , , "Orden de Compra", 2000
         .ColumnHeaders.Add , , "Tipo", 3000
        .ColumnHeaders.Add , , "Fecha", 3000
       
              
        End With
        With ListView4
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "O.C", 1000
        .ColumnHeaders.Add , , "ID_PRODUCTO", 1500
         .ColumnHeaders.Add , , "DESCRIPCION", 2500
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "PRECIO", 1000
         
            .ColumnHeaders.Add , , "SURTIDO", 1000
            .ColumnHeaders.Add , , "STATUS", 3000
              
        End With
        
        
        sBuscar = "SELECT NOMBRE FROM SUCURSALES ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    Combo1.AddItem "<TODAS>"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
     
End Sub

Private Sub Imprimir()

    Dim Path As String
    Dim SelectionFormula As Date
     Path = App.Path
     'sBuscar = "SELECT NOMBRE,ID_VENTA,FOLIO,FECHA,NOMBRE,BANCO,CANT_ABONO,NO_CHEQUE,DEUDA,LIMITE_CREDITO FROM VsRepAbonos WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY FECHA"
    'StrRep = sBuscar
        Set crReport = crApplication.OpenReport(Path & "\REPORTES\repdev.rpt")
        crReport.Database.LogOnServer "p2ssql.dll", GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "1"), GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER"), GetSetting("APTONER", "ConfigSACC", "USUARIO", "1"), GetSetting("APTONER", "ConfigSACC", "PASSWORD", "1")
        'crReport.SelectionFormula = "{VSABONOS.FECHA}>=Date (Year (" & DTPicker1.Value & "),Month (" & DTPicker1.Value & ") , Day (" & DTPicker1.Value & ")) and {VSABONOS.FECHA}<=Date (Year (" & DTPicker2.Value & "),Month (" & DTPicker2.Value & ") , Day ( " & DTPicker2.Value & "))"
        'crReport.Action = 0
       'crReport.Destination = crptToPrinter
      'crReport.Destination = crptToWindow
        crReport.SQLQueryString = StrRep
        'crReport.SQLQueryString = "SELECT NOMBRE,ID_VENTA,FOLIO,FECHA,NOMBRE,BANCO,CANT_ABONO,NO_CHEQUE,DEUDA,LIMITE_CREDITO FROM VsRepAbonos WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY FECHA"
        crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        frmRep.Show vbModal, Me
        
     End Sub



Private Sub Image10_Click()
    If ListView1.ListItems.Count > 0 Then
        Dim StrCopi As String
        Dim Con As Integer
        Dim Con2 As Integer
        Dim NumColum As Integer
        Dim Ruta As String
        
        'FileName
        Me.CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "Guardar como"
        CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
        Me.CommonDialog1.ShowSave
        Ruta = Me.CommonDialog1.FileName
        StrCopi = "Nota" & Chr(9) & "Factura" & Chr(9) & "Fecha" & Chr(9) & "Nombre" & Chr(9) & "Banco" & Chr(9) & "Monto" & Chr(9) & "DEUDA" & Chr(9) & Chr(9) & "LIMITE_CREDITO" & Chr(13)
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.Count
            For Con = 1 To ListView1.ListItems.Count
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            Text2.Text = StrCopi
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, Text2.Text
            Close #foo
        End If
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

Text7.Text = Item.SubItems(1)
Text1.Text = Item.SubItems(3)
Text2.Text = Item.SubItems(2)
Text3.Text = Item.SubItems(4)
Text4.Text = Item.SubItems(5)
Text8.Text = Item.SubItems(7)
Text9.Text = Item.SubItems(9)
Text10.Text = Item.SubItems(8)



        
    
    
    'DesProd = Item.SubItems(1)
    'PreProd = Item.SubItems(2)
    'ClasProd = Item.SubItems(4)
    End Sub
    



Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs3 As Recordset
    Dim orde As Integer
    Dim tip As String
      Dim pro As String
ListView3.ListItems.Clear
Dim ordeee As Integer



            sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE   ID_PROVEEDOR= '" & Item & "'   ORDER BY FECHA DESC   "
            'AND TIPO= '" & txttipo & "'"
             Set tRs = cnn.Execute(sBuscar)
     
     
     
         
     
           
        If Not (tRs.EOF And tRs.BOF) Then
          Do While Not (tRs.EOF)
            Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_ORDEN_COMPRA"))
            If Not IsNull(tRs.Fields("NUM_ORDEN")) Then tLi.SubItems(1) = tRs.Fields("NUM_ORDEN")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(3) = tRs.Fields("FECHA")
             If tRs.Fields("TIPO") = "I" Then
             tLi.SubItems(2) = "INTERNACIONAL"
             End If
              If tRs.Fields("TIPO") = "N" Then
             tLi.SubItems(2) = "NACIONAL"
             End If
            tRs.MoveNext
        Loop
End If
    
End Sub

Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
 Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs3 As Recordset
    Dim orde As Integer
    Dim tip As String
      Dim pro As String
ListView4.ListItems.Clear
Dim ordeee As Integer



            sBuscar = "SELECT * FROM vsordencom WHERE   ID_ORDEN_COMPRA= '" & Item & "'     "
            'AND TIPO= '" & txttipo & "'"
             Set tRs = cnn.Execute(sBuscar)
     
     
     
         
     
           
        If Not (tRs.EOF And tRs.BOF) Then
          Do While Not (tRs.EOF)
            Set tLi = ListView4.ListItems.Add(, , tRs.Fields("NUM_ORDEN"))
            'If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
           ' If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(2) = tRs.Fields("FECHA")
            'If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(3) = tRs.Fields("TOTAL")
           ' If Not IsNull(tRs.Fields("ID_PROVEEDOR")) Then tLi.SubItems(4) = tRs.Fields("ID_PROVEEDOR")
            'If Not IsNull(tRs.Fields("TIPO")) Then tLi.SubItems(5) = tRs.Fields("TIPO")
            orde = tRs.Fields("NUM_ORDEN")
                      
             If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
             If Not IsNull(tRs.Fields("DESCRIPCION")) Then tLi.SubItems(2) = tRs.Fields("DESCRIPCION")
             If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(3) = tRs.Fields("CANTIDAD")
             If Not IsNull(tRs.Fields("PRECIO")) Then tLi.SubItems(4) = tRs.Fields("PRECIO")
             If Not IsNull(tRs.Fields("SURTIDO")) Then tLi.SubItems(5) = tRs.Fields("SURTIDO")
                                    
                
               'pro = tRs.Fields("ID_PRODUCTO")
               
            'tip = tRs.Fields("TIPO")
            If tRs.Fields("CONFIRMADA") = "N" Then
                tLi.SubItems(6) = "PRE-ORDEN"
            End If
            If tRs.Fields("CONFIRMADA") = "P" Then
                tLi.SubItems(6) = "PENDIENTE DE AUTORIZAR"
            End If
            If tRs.Fields("CONFIRMADA") = "S" Then
                tLi.SubItems(6) = "PENDIENTE DE IMPRIMIR"
            End If
            
            sBuscar = "SELECT * FROM vsordpende WHERE NUM_ORDEN= '" & tRs.Fields("NUM_ORDEN") & "' AND  TIPO= '" & tRs.Fields("TIPO") & "' AND  ID_PRODUCTO= '" & tRs.Fields("ID_PRODUCTO") & "'"
            Set tRs3 = cnn.Execute(sBuscar)
                Dim catpe  As Double
                catpe = CDbl(tRs3.Fields("CANTIDAD")) - CDbl(tRs3.Fields("SURTIDO"))
            If tRs.Fields("CONFIRMADA") = "X" And tRs3.Fields("SURTIDO") = 0 Then
                tLi.SubItems(6) = "PENDIENTE DE  LLEGAR "
            End If
            If tRs.Fields("CONFIRMADA") = "X" And catpe = 0 Then
                tLi.SubItems(6) = "PENDIENTE DE PAGO/EN ALMACEN"
            End If
            
            If tRs.Fields("CONFIRMADA") = "X" And catpe < tRs3.Fields("CANTIDAD") And tRs3.Fields("SURTIDO") < 0 Then
          
                tLi.SubItems(6) = "PENDIENTE DE PAGO/LLEGADA PARCIAL"
            End If
            
            If tRs.Fields("CONFIRMADA") = "Y" Then
                tLi.SubItems(6) = "PAGADA"
            End If
            catpe = 0
             
            tRs.MoveNext
        Loop
End If
End Sub

Private Sub option1_Click()
Text5.Visible = True
Command1.Visible = True
Label5.Visible = True
End Sub

Private Sub Option2_Click()
Text5.Visible = False
Command1.Visible = False
Label5.Visible = False
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)

    Dim sqlQuery As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
    Dim tRs3 As Recordset
    Dim orde As Integer
    Dim tip As String
    Dim pro As String

    ListView2.ListItems.Clear
    If KeyAscii = 13 Then
        
        
                sqlQuery = "SELECT * FROM PROVEEDOR WHERE NOMBRE LIKE '%" & Text11.Text & "%'"
                Set tRs = cnn.Execute(sqlQuery)
                With tRs
                    If Not (.BOF And .EOF) Then
                        Do While Not .EOF
                            Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PROVEEDOR"))
                            If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = Trim(.Fields("NOMBRE"))
                            .MoveNext
                        Loop
                    End If
                End With
                '
   End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command1.Value = True
    End If
End Sub



