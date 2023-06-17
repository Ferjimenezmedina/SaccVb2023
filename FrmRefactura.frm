VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRefactura 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cancelar Ventas Refacturando"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " Ventas"
      TabPicture(0)   =   "FrmRefactura.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Facturar"
      TabPicture(1)   =   "FrmRefactura.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Option5"
      Tab(1).Control(1)=   "Command4"
      Tab(1).Control(2)=   "Option4"
      Tab(1).Control(3)=   "Option3"
      Tab(1).Control(4)=   "Option2"
      Tab(1).Control(5)=   "Option1"
      Tab(1).Control(6)=   "Check1"
      Tab(1).Control(7)=   "Command3"
      Tab(1).Control(8)=   "Text4"
      Tab(1).Control(9)=   "Text3"
      Tab(1).Control(10)=   "Command2"
      Tab(1).Control(11)=   "Text2"
      Tab(1).Control(12)=   "ListView3"
      Tab(1).Control(13)=   "ListView4"
      Tab(1).Control(14)=   "Label6"
      Tab(1).Control(15)=   "Label5"
      Tab(1).Control(16)=   "Label4"
      Tab(1).Control(17)=   "Label3"
      Tab(1).ControlCount=   18
      Begin VB.CommandButton Command6 
         Caption         =   "-"
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
         Left            =   7320
         Picture         =   "FrmRefactura.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5760
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "+"
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
         Left            =   6960
         Picture         =   "FrmRefactura.frx":2A0A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5760
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         Caption         =   "T.Debito"
         Height          =   255
         Left            =   -70920
         TabIndex        =   16
         Top             =   5640
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Eliminar"
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
         Left            =   -68400
         Picture         =   "FrmRefactura.frx":53DC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5640
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Cheque"
         Height          =   255
         Left            =   -72960
         TabIndex        =   14
         Top             =   5640
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "T. Credito"
         Height          =   255
         Left            =   -72000
         TabIndex        =   15
         Top             =   5640
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Transferencia"
         Height          =   255
         Left            =   -69960
         TabIndex        =   17
         Top             =   5640
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Efectivo"
         Height          =   255
         Left            =   -73920
         TabIndex        =   13
         Top             =   5640
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Credito"
         Height          =   195
         Left            =   -74760
         TabIndex        =   12
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Agregar"
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
         Left            =   -68400
         Picture         =   "FrmRefactura.frx":7DAE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   -69960
         TabIndex        =   9
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -72000
         TabIndex        =   8
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
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
         Left            =   -68400
         Picture         =   "FrmRefactura.frx":A780
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -73800
         TabIndex        =   6
         Top             =   720
         Width           =   5175
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
         Left            =   6600
         Picture         =   "FrmRefactura.frx":D152
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   720
         Width           =   5295
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3836
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2175
         Left            =   120
         TabIndex        =   3
         Top             =   3600
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3836
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
         Height          =   1815
         Left            =   -74880
         TabIndex        =   7
         Top             =   1080
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3201
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
         Height          =   2055
         Left            =   -74880
         TabIndex        =   11
         Top             =   3480
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Precio"
         Height          =   255
         Left            =   -70560
         TabIndex        =   29
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   -72840
         TabIndex        =   28
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Productos :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Ventas"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8160
      TabIndex        =   21
      Top             =   3840
      Width           =   975
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRefactura.frx":FB24
         MousePointer    =   99  'Custom
         Picture         =   "FrmRefactura.frx":FE2E
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8160
      TabIndex        =   19
      Top             =   5040
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRefactura.frx":117F0
         MousePointer    =   99  'Custom
         Picture         =   "FrmRefactura.frx":11AFA
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label9 
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
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmRefactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim IdCliente As String
Dim Nombre As String
Dim IdProducto As String
Dim Descripcion As String
Dim Elim As Integer
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%' ORDER BY ID_CLIENTE"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
    Exit Sub
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%' ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    ListView3.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("PRECIO_COSTO")) Then tLi.SubItems(2) = tRs.Fields("PRECIO_COSTO") * (1 + tRs.Fields("GANANCIA"))
            tRs.MoveNext
        Loop
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
    Exit Sub
End Sub
Private Sub Command3_Click()
    If IdProducto <> "" Then
        Dim tLi As ListItem
        Set tLi = ListView4.ListItems.Add(, , IdProducto)
        If Not IsNull(Descripcion) Then tLi.SubItems(1) = Descripcion
        If Not IsNull(Text3.Text) Then tLi.SubItems(2) = Text3.Text
        If Not IsNull(Text4.Text) Then tLi.SubItems(3) = Text4.Text
        Text3.Text = ""
        Text4.Text = ""
        IdProducto = ""
        Descripcion = ""
        Label6.Caption = ""
    End If
End Sub
Private Sub Command4_Click()
    If Elim <> 0 Then
        ListView4.ListItems.Remove (Elim)
        Elim = 0
    End If
End Sub
Private Sub Command5_Click()
    Dim Cont As Double
    For Cont = 1 To ListView2.ListItems.COUNT
        ListView2.ListItems(Cont).Checked = True
    Next Cont
End Sub
Private Sub Command6_Click()
    Dim Cont As Double
    For Cont = 1 To ListView2.ListItems.COUNT
        ListView2.ListItems(Cont).Checked = False
    Next Cont
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .FullRowSelect = True
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID CLIENTE", 2000
        .ColumnHeaders.Add , , "NOMBRE", 6000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .FullRowSelect = True
        .Checkboxes = True
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID VENTA", 2000
        .ColumnHeaders.Add , , "FECHA", 2000
        .ColumnHeaders.Add , , "TOTAL", 2000
        .ColumnHeaders.Add , , "FACTURA", 2000
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .FullRowSelect = True
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID PRODUCTO", 2000
        .ColumnHeaders.Add , , "Descripcion", 4000
        .ColumnHeaders.Add , , "PRECIO", 2000
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .FullRowSelect = True
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID PRODUCTO", 2000
        .ColumnHeaders.Add , , "Descripcion", 4000
        .ColumnHeaders.Add , , "CANTIDAD", 2000
        .ColumnHeaders.Add , , "PRECIO", 2000
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Image8_Click()
    On Error GoTo ManejaError
    Dim tRs8 As ADODB.Recordset
    Dim tRs7 As ADODB.Recordset
    Dim Item As MSComctlLib.ListItem
    Dim tLi As ListItem
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim NumeroRegistros As Integer
    Dim Conta As Integer
    Dim NRegistros As Integer
    Dim Con As Integer
    Dim POSY As Integer
    Dim FechaVence As String
    Dim sBuscar As String
    Dim P_COSTO As String
    Dim cant As String
    Dim CanProd As String
    Dim P_ven As String
    Dim IdVentAut As String
    Dim Ganan As String
    Dim IdCta As String
    Dim TPago As String
    Dim TotDeuda As Double
    Dim Vale As Double
    Dim MosLey As String
    Dim ID_VALE As Integer
    Dim continuar As Boolean
    Dim AbonoClien As String
    Dim sExibicion As String
    Dim TotPago As Double
    If Option1.value = True Then
        TPago = "C"
    End If
    If Option4.value = True Then
        TPago = "H"
    End If
    If Option3.value = True Then
        TPago = "T"
    End If
    If Option2.value = True Then
        TPago = "E"
    End If
    If Option5.value = True Then
        TPago = "D"
    End If
    If Option1.value = False And Option2.value = False And Option3.value = False And Option4.value = False Then
        MsgBox "DEBE MARCAR UNA FORMA DE PAGO!", vbExclamation, "SACC"
        Exit Sub
    End If
    For Con = 1 To ListView4.ListItems.COUNT
        TotPago = CDbl(ListView4.ListItems(Con).SubItems(3) * ListView4.ListItems(Con).SubItems(2)) + TotPago
    Next Con
    If Check1.value = 0 Then
        sBuscar = "INSERT INTO VENTAS (ID_CLIENTE, NOMBRE, DESCUENTO, FECHA, TOTAL, SUCURSAL, ID_USUARIO, IVA, SUBTOTAL, TIPO_PAGO, UNA_EXIBICION, NOOC, COMENTARIO, FORMA_PAGO) VALUES (" & IdCliente & ", '" & Nombre & "', 0,  DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), " & TotPago + ((VarMen.Text4(7).Text / 100) * TotPago) & ", '" & VarMen.Text4(0).Text & "', '" & VarMen.Text1(0).Text & "', " & (VarMen.Text4(7).Text / 100) * TotPago & ", " & TotPago & ", '" & TPago & "', 'S', '0', '0', 'PAGO EN UNA SOLA EXHIBICION');"
        cnn.Execute (sBuscar)
        sBuscar = "SELECT ID_VENTA, UNA_EXIBICION FROM VENTAS WHERE SUCURSAL = '" & VarMen.Text4(0).Text & "' ORDER BY ID_VENTA DESC"
        Set tRs = cnn.Execute(sBuscar)
        sExibicion = tRs.Fields("UNA_EXIBICION")
        IdVentAut = tRs.Fields("ID_VENTA")
        NumeroRegistros = ListView4.ListItems.COUNT
        For Conta = 1 To NumeroRegistros
            cant = Replace(ListView4.ListItems.Item(Conta).SubItems(2), ",", "")
            P_ven = Replace(ListView4.ListItems.Item(Conta).SubItems(3), ",", "")
            If Val(cant) > 0 Then
                P_ven = Format(Val(P_ven) / Val(cant), "###,###,##0.00")
                P_ven = Replace(P_ven, ",", "")
            Else
                P_ven = "0.00"
            End If
            sBuscar = "INSERT INTO VENTAS_DETALLE (ID_PRODUCTO, DESCRIPCION, PRECIO_VENTA, CANTIDAD, ID_VENTA, NO_COM_AT, IMPORTE) VALUES ('" & ListView4.ListItems.Item(Conta) & "', '" & ListView4.ListItems.Item(Conta).SubItems(1) & "', " & P_ven & ", " & cant & ", " & IdVentAut & ", '0', " & Format(Val(P_ven) * Val(cant), "###,###,##0.00") & ");"
            cnn.Execute (sBuscar)
        Next Conta
        
        sBuscar = "SELECT * FROM SUCURSALES WHERE NOMBRE = '" & VarMen.Text4(0).Text & "' AND ELIMINADO = 'N'"
        Set tRs8 = cnn.Execute(sBuscar)
        '********************************IMPRIMIR TICKET********************************************
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(tRs8.Fields("CALLE") & " COL. " & tRs8.Fields("COLONIA"))) / 2
        Printer.Print tRs8.Fields("CALLE") & " COL. " & tRs8.Fields("COLONIA")
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & tRs8.Fields("CP"))) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & tRs8.Fields("CP")
        Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
        Printer.Print "SUCURSAL : " & VarMen.Text4(0).Text
        Printer.Print "TELEFONO SUCURSAL : " & tRs8.Fields("TELEFONO")
        Printer.Print "No. DE VENTA : " & IdVentAut
        If Option1.value = True Then
            Printer.Print "FORMA DE PAGO : EFECTIVO"
        Else
            If Option4.value = True Then
                Printer.Print "FORMA DE PAGO : CHEQUE"
            Else
                If Option3.value = True Then
                    Printer.Print "FORMA DE PAGO : TARJETA DE CREDITO"
                Else
                    If Option2.value = True Then
                        Printer.Print "FORMA DE PAGO : TRANSFERENCIA ELECTRONICA"
                    Else
                        Printer.Print "FORMA DE PAGO : NO INDICADO"
                    End If
                End If
            End If
        End If
        Printer.Print "ATENDIDO POR : " & VarMen.Text1(1).Text & " "; VarMen.Text1(2).Text
        Printer.Print "CLIENTE : " & Nombre
        If sExibicion = "N" Then
            Printer.Print "VENTA A CREDITO"
        Else
            Printer.Print "VENTA DE CONTADO"
        End If
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                          NOTA DE FACTURA"
        Printer.Print "--------------------------------------------------------------------------------"
        NRegistros = ListView4.ListItems.COUNT
        POSY = 2900
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 1900
        Printer.Print "Cant."
        Printer.CurrentY = POSY
        Printer.CurrentX = 3000
        Printer.Print "Precio unitario"
        For Con = 1 To NRegistros
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print ListView4.ListItems(Con).Text
            Printer.CurrentY = POSY
            Printer.CurrentX = 1900
            Printer.Print ListView4.ListItems(Con).SubItems(2)
            Printer.CurrentY = POSY
            Printer.CurrentX = 2900
            Printer.Print CDbl(ListView4.ListItems(Con).SubItems(3))
        Next Con
        Printer.Print ""
        Printer.Print "SUBTOTAL : " & TotPago
        Printer.Print "IVA              : " & (VarMen.Text4(7).Text / 100) * TotPago
        Printer.Print "TOTAL        : " & TotPago + ((VarMen.Text4(7).Text / 100) * TotPago)
        Printer.Print ""
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "               GRACIAS POR SU COMPRA"
        Printer.Print "LA GARANTÍA SERÁ VÁLIDA HASTA 15 DIAS"
        Printer.Print "     DESPUES DE HABER EFECTUADO SU "
        Printer.Print "                                COMPRA"
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.EndDoc
        For Con = 1 To ListView2.ListItems.COUNT
            If ListView2.ListItems(Con).Checked Then
                sBuscar = "UPDATE VENTAS SET FOLIO = 'CANCELADO', FACTURADO = 2, FLAG_CANCELADO = 'S' WHERE ID_VENTA = " & ListView2.ListItems(Con)
                cnn.Execute (sBuscar)
            End If
        Next Con
        sBuscar = "DELETE FROM CUENTAS WHERE (ID_CUENTA IN (SELECT ID_CUENTA From CUENTA_VENTA WHERE (ID_VENTA IN (SELECT ID_VENTA From Ventas WHERE (FACTURADO = 2)))))"
        cnn.Execute (sBuscar)
        ListView1.ListItems.Clear
        ListView2.ListItems.Clear
        ListView3.ListItems.Clear
        ListView4.ListItems.Clear
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
    Else
        sBuscar = "SELECT SUM(TOTAL_COMPRA) AS TOTAL FROM CUENTAS WHERE ID_CLIENTE = " & IdCliente
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            If IsNull(tRs.Fields("TOTAL")) Then
                TotDeuda = 0
            Else
                TotDeuda = tRs.Fields("TOTAL")
            End If
        Else
            TotDeuda = 0
        End If
        TotDeuda = TotDeuda + Val(Replace(TotPago, ",", ""))
        sBuscar = "SELECT SUM(CANT_ABONO) AS TOTAL FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & IdCliente
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            If IsNull(tRs.Fields("TOTAL")) Then
                AbonoClien = 0
            Else
                AbonoClien = tRs.Fields("TOTAL")
            End If
        Else
            AbonoClien = 0
        End If
        TotDeuda = TotDeuda - CDbl(Replace(AbonoClien, ",", ""))
        '********************************* PARA CREDITO ********************************
        'CAMBIAR TODO EL PROCEDIMIENTO, ASEGURARSE QUE INSERTE LA DEUDA
        NRegistros = ListView3.ListItems.COUNT
        sBuscar = "SELECT LEYENDAS, DIAS_CREDITO FROM CLIENTE WHERE ID_CLIENTE = " & IdCliente
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            FechaVence = Format(Date + CDbl(tRs2.Fields("DIAS_CREDITO")), "dd/mm/yyyy")
            If tRs2.Fields("LEYENDAS") = "N" Then
                MosLey = ""
            Else
                MosLey = "PAGO EN PARCIALIDADES"
            End If
        End If
        sBuscar = "INSERT INTO VENTAS (ID_CLIENTE, NOMBRE, DESCUENTO, FECHA, TOTAL, SUCURSAL, ID_USUARIO, IVA, SUBTOTAL, TIPO_PAGO, UNA_EXIBICION, NOOC, COMENTARIO, FORMA_PAGO) VALUES (" & IdCliente & ", '" & Nombre & "', 0,  DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), " & TotPago + ((VarMen.Text4(7).Text / 100) * TotPago) & ", '" & VarMen.Text4(0).Text & "', '" & VarMen.Text1(0).Text & "', " & (VarMen.Text4(7).Text / 100) * TotPago & ", " & TotPago & ", '" & TPago & "', 'N', '0', '0', 'PAGO EN PARCIALIDADES');"
        cnn.Execute (sBuscar)
        sBuscar = "SELECT ID_VENTA FROM VENTAS WHERE SUCURSAL = '" & VarMen.Text4(0).Text & "' ORDER BY ID_VENTA DESC"
        Set tRs = cnn.Execute(sBuscar)
        IdVentAut = tRs.Fields("ID_VENTA")
        sBuscar = "INSERT INTO CUENTAS (PAGADA, ID_CLIENTE, ID_USUARIO, FECHA, DIAS_CREDITO, FECHA_VENCE, DESCUENTO, SUCURSAL, TOTAL_COMPRA, DEUDA, ID_VENTA) VALUES ( 'N', " & IdCliente & ", '" & VarMen.Text1(0).Text & "',  DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), " & tRs2.Fields("DIAS_CREDITO") & ", '" & FechaVence & "', 0, '" & VarMen.Text4(0).Text & "', " & TotPago + ((VarMen.Text4(7).Text / 100) * TotPago) & ", " & TotPago + ((VarMen.Text4(7).Text / 100) * TotPago) & ", " & IdVentAut & ");"
        cnn.Execute (sBuscar)
        sBuscar = "SELECT TOP 1 ID_CUENTA FROM CUENTAS ORDER BY ID_CUENTA DESC"
        Set tRs = cnn.Execute(sBuscar)
        IdCta = tRs.Fields("ID_CUENTA")
        sBuscar = "INSERT INTO CUENTA_VENTA (ID_VENTA, ID_CUENTA) VALUES (" & IdVentAut & ", " & IdCta & ");"
        cnn.Execute (sBuscar)
        NRegistros = ListView4.ListItems.COUNT
        For Conta = 1 To NRegistros
            CanProd = Replace(ListView4.ListItems(Conta).SubItems(3), ",", "")
            P_ven = Format(CDbl(ListView4.ListItems(Conta).SubItems(2)), "###,###,##0.00")
            P_ven = Replace(P_ven, ",", "")
            sBuscar = "INSERT INTO CUENTA_DETALLE (ID_CUENTA, CANTIDAD, ID_PRODUCTO, PRECIO_VENTA) VALUES (" & IdCta & ", " & ListView4.ListItems(Conta).SubItems(2) & ", '" & ListView4.ListItems(Conta) & "', " & ListView4.ListItems(Conta).SubItems(3) & ");"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView4.ListItems(Conta).Text & "'"
            Set tRs = cnn.Execute(sBuscar)
            P_COSTO = tRs.Fields("PRECIO_COSTO")
            Ganan = tRs.Fields("GANANCIA")
            P_COSTO = Replace(P_COSTO, ",", "")
            Ganan = Replace(Ganan, ",", "")
            sBuscar = "INSERT INTO VENTAS_DETALLE (ID_VENTA, CANTIDAD, ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO, GANANCIA, PRECIO_VENTA, IMPORTE) VALUES (" & IdVentAut & ", " & ListView4.ListItems(Conta).SubItems(2) & ", '" & ListView4.ListItems(Conta).Text & "', '" & ListView4.ListItems(Conta).SubItems(1) & "', " & P_COSTO & ", " & Ganan & ", " & ListView4.ListItems(Conta).SubItems(3) & ", " & CDbl(P_ven) * CDbl(CanProd) & ");"
            cnn.Execute (sBuscar)
        Next Conta
        '********************************IMPRIMIR TICKET********************************************
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
        Printer.Print "SUCURSAL : " & VarMen.Text4(0).Text
        Printer.Print "No. DE VENTA : " & IdVentAut
        Printer.Print "ATENDIDO POR : " & VarMen.Text1(1).Text & " "; VarMen.Text1(2).Text
        If Option1.value = True Then
            Printer.Print "FORMA DE PAGO : EFECTIVO"
        Else
            If Option4.value = True Then
                Printer.Print "FORMA DE PAGO : CHEQUE"
            Else
                If Option3.value = True Then
                    Printer.Print "FORMA DE PAGO : TARJETA DE CREDITO"
                Else
                    If Option2.value = True Then
                        Printer.Print "FORMA DE PAGO : TRANSFERENCIA ELECTRONICA"
                    Else
                        Printer.Print "FORMA DE PAGO : NO INDICADO"
                    End If
                End If
            End If
        End If
        Printer.Print "CLIENTE : " & Nombre
        Printer.Print "VENTA A CREDITO"
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                          NOTA DE FACTURA"
        Printer.Print "--------------------------------------------------------------------------------"
        NRegistros = ListView4.ListItems.COUNT
        POSY = 2600
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 1300
        Printer.Print "Cant."
        Printer.CurrentY = POSY
        Printer.CurrentX = 3000
        Printer.Print "Precio unitario"
        For Con = 1 To NRegistros
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print ListView4.ListItems(Con).Text
            Printer.CurrentY = POSY
            Printer.CurrentX = 1900
            Printer.Print ListView4.ListItems(Con).SubItems(2)
            Printer.CurrentY = POSY
            Printer.CurrentX = 2900
            Printer.Print ListView4.ListItems(Con).SubItems(3)
        Next Con
        Printer.Print ""
        Printer.Print "SUBTOTAL : " & TotPago
        Printer.Print "IVA              : " & (VarMen.Text4(7).Text / 100) * TotPago
        Printer.Print "TOTAL        : " & TotPago + ((VarMen.Text4(7).Text / 100) * TotPago)
        Printer.Print ""
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "               GRACIAS POR SU COMPRA"
        Printer.Print "           PRODUCTO 100% GARANTIZADO"
        Printer.Print "LA GARANTÍA SERÁ VÁLIDA HASTA 15 DIAS"
        Printer.Print "     DESPUES DE HABER EFECTUADO SU "
        Printer.Print "                                COMPRA"
        Printer.Print "SIN SU TICKET NO SERA VALIDA LA GARANTIA."
        Printer.Print "                APLICA RESTRICCIONES"
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.EndDoc
        For Con = 1 To ListView2.ListItems.COUNT
            If ListView2.ListItems(Con).Checked Then
                sBuscar = "UPDATE VENTAS SET FOLIO = 'CANCELADO', FACTURADO = 2, FLAG_CANCELADO = 'S' WHERE ID_VENTA = " & ListView2.ListItems(Con)
                cnn.Execute (sBuscar)
            End If
        Next Con
        sBuscar = "DELETE FROM CUENTAS WHERE (ID_CUENTA IN (SELECT ID_CUENTA From CUENTA_VENTA WHERE (ID_VENTA IN (SELECT ID_VENTA From Ventas WHERE (FACTURADO = 2)))))"
        cnn.Execute (sBuscar)
        ListView1.ListItems.Clear
        ListView2.ListItems.Clear
        ListView3.ListItems.Clear
        ListView4.ListItems.Clear
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdCliente = Item
    Nombre = Item.SubItems(1)
    On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_VENTA, FECHA, TOTAL, FOLIO FROM VENTAS WHERE ID_CLIENTE = " & Item & " AND FACTURADO = 0 ORDER BY ID_VENTA"
    Set tRs = cnn.Execute(sBuscar)
    ListView2.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_VENTA"))
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(1) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(2) = tRs.Fields("TOTAL")
            If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(3) = tRs.Fields("FOLIO")
            tRs.MoveNext
        Loop
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
    Exit Sub
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If IdCliente <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        sBuscar = "SELECT PRECIO_VENTA FROM LICITACIONES WHERE ID_PRODUCTO ='" & Item & "' AND ID_CLIENTE = " & IdCliente
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Text4.Text = tRs.Fields("PRECIO_VENTA")
        Else
            Text4.Text = Item.SubItems(2)
        End If
        Label6.Caption = Item
        IdProducto = Item
        Descripcion = Item.SubItems(1)
    End If
End Sub
Private Sub ListView4_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Elim = ListView4.SelectedItem.Index
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command1.value = True
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command2.value = True
    End If
End Sub
