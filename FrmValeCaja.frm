VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmValeCaja 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hacer Vale de Caja"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   16
      Top             =   3240
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
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmValeCaja.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmValeCaja.frx":030A
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   9
      Top             =   4440
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
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmValeCaja.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "FrmValeCaja.frx":1FD6
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Venta"
      TabPicture(0)   =   "FrmValeCaja.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Vale"
      TabPicture(1)   =   "FrmValeCaja.frx":40D4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton Command3 
         Caption         =   "Quitar"
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
         Left            =   -68040
         Picture         =   "FrmValeCaja.frx":40F0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "General"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   19
         Top             =   4080
         Width           =   6735
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   960
            TabIndex        =   6
            Top             =   960
            Width           =   5655
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   960
            TabIndex        =   23
            Top             =   240
            Width           =   5655
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "0"
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Motivo :"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Cliente :"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Total :"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.CommandButton Command2 
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
         Left            =   3840
         Picture         =   "FrmValeCaja.frx":6AC2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5040
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   5
         Top             =   1080
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5106
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame1 
         Caption         =   "Selección"
         Height          =   1335
         Left            =   120
         TabIndex        =   12
         Top             =   4080
         Width           =   3615
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1080
            TabIndex        =   3
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1080
            TabIndex        =   14
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "Cantidad :"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Producto :"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   735
         End
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
         Left            =   3240
         Picture         =   "FrmValeCaja.frx":9494
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   720
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5106
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   5160
         X2              =   5160
         Y1              =   4200
         Y2              =   5400
      End
      Begin VB.Label Label4 
         Caption         =   $"FrmValeCaja.frx":BE66
         Height          =   1215
         Left            =   5280
         TabIndex        =   18
         Top             =   4200
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Numero de Venta :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmValeCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim EliItm As Integer
Dim IndAgr As Integer
Dim Itm As String
Dim SubItm1 As String
Dim SubItm2 As String
Dim SubItm3 As String
Dim SubItm4 As String
Dim SubItm5 As String
Private Sub Command1_Click()
    If ListView2.ListItems.Count = 0 Then
        Dim tRs As ADODB.Recordset
        Dim sBuscar As String
        Dim tLi As ListItem
        ListView1.ListItems.Clear
        sBuscar = "SELECT  VENTAS.ID_VENTA, VENTAS.NOMBRE, VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.DESCRIPCION, VENTAS_DETALLE.CANTIDAD, VENTAS_DETALLE.Precio_Venta FROM VENTAS INNER JOIN VENTAS_DETALLE ON VENTAS.ID_VENTA = VENTAS_DETALLE.ID_VENTA WHERE VENTAS.ID_VENTA = " & Text1.Text & " AND VENTAS.FACTURADO = 0 AND VENTAS.FECHA <> '" & Format(Date, "dd/mm/yyyy") & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = tRs.Fields("ID_PRODUCTO")
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(3) = tRs.Fields("Descripcion")
                If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(4) = tRs.Fields("CANTIDAD")
                If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(5) = tRs.Fields("PRECIO_VENTA")
                tRs.MoveNext
            Loop
        Else
            MsgBox "LA VENTA NO EXISTE, FUE CANCELADA, FACTURADA O LA FECHA NO ES VALIDA PARA REALIZAR EL VALE", vbInformation, "SACC"
        End If
    Else
        MsgBox "NO PUEDE REALIZAR UN VALE DE CAJA POR VARIAS VENTAS, TIENE QUE TERMINAR O ELIMINAR LOS ARTICULOS SELECCIONADOS", vbInformation, "SACC"
    End If
End Sub
Private Sub Command2_Click()
    If Itm <> "" And IndAgr <> 0 Then
        If Text3.Text = "" Then
            Text3.Text = 1
        Else
            If CDbl(Text3.Text) = 0 Then
                Text3.Text = 1
            End If
        End If
        If CDbl(Text3.Text) <= CDbl(SubItm4) Then
            Dim tLi As ListItem
            Set tLi = ListView2.ListItems.Add(, , Itm)
            tLi.SubItems(1) = SubItm1
            tLi.SubItems(2) = SubItm2
            tLi.SubItems(3) = SubItm3
            tLi.SubItems(4) = Text3.Text
            tLi.SubItems(5) = SubItm5
            Text4.Text = Format(CDbl(Text4.Text) + ((CDbl(Text3.Text) * CDbl(SubItm5)) * (CDbl(CDbl(VarMen.Text4(7).Text) / 100) + 1)), "###,###,##0.00")
            ListView1.ListItems(IndAgr).SubItems(4) = CDbl(ListView1.ListItems(IndAgr).SubItems(4)) - CDbl(Text3.Text)
            IndAgr = 0
            Itm = ""
            SubItm1 = ""
            SubItm2 = ""
            SubItm3 = ""
            SubItm4 = ""
            SubItm5 = ""
            Text2.Text = ""
            Text3.Text = ""
        End If
    Else
        MsgBox "DEBE SELECCIONAR UN ARTICULO PARA AGREGARLO!", vbInformation, "SACC"
    End If
End Sub
Private Sub Command3_Click()
    If EliItm <> 0 Then
        Dim Con As Integer
        For Con = 1 To ListView1.ListItems.Count
            If ListView2.ListItems(EliItm).SubItems(2) = ListView1.ListItems(Con).SubItems(2) Then
                ListView1.ListItems(Con).SubItems(4) = CDbl(ListView1.ListItems(Con).SubItems(4)) + CDbl(ListView2.ListItems(Con).SubItems(4))
            End If
        Next
        Text4.Text = Format(CDbl(Text4.Text) - ((CDbl(Text3.Text) * CDbl(ListView2.ListItems(EliItm).SubItems(5)) * (CDbl(CDbl(VarMen.Text4(7).Text) / 100) + 1))), "###,###,##0.00")
        ListView2.ListItems.Remove (EliItm)
        EliItm = 0
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
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
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. VENTA", 800
        .ColumnHeaders.Add , , "CLIENTE", 4200
        .ColumnHeaders.Add , , "PRODUCTO", 2000
        .ColumnHeaders.Add , , "Descripcion", 4000
        .ColumnHeaders.Add , , "CANTIDAD", 1500
        .ColumnHeaders.Add , , "PRECIO DE VENTA", 2000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. VENTA", 800
        .ColumnHeaders.Add , , "CLIENTE", 4200
        .ColumnHeaders.Add , , "PRODUCTO", 2000
        .ColumnHeaders.Add , , "Descripcion", 4000
        .ColumnHeaders.Add , , "CANTIDAD", 1500
        .ColumnHeaders.Add , , "PRECIO DE VENTA", 2000
    End With
End Sub
Private Sub Image8_Click()
    If ListView2.ListItems.Count <> 0 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tRs1 As ADODB.Recordset
        Dim Con As Integer
        sBuscar = "SELECT FECHA FROM VALE_CAJA WHERE ID_VENTA = " & ListView2.ListItems(1)
        Set tRs = cnn.Execute(sBuscar)
        If tRs.EOF And tRs.BOF Then
            sBuscar = "INSERT INTO VALE_CAJA (ID_VENTA, IMPORTE, FECHA, APLICADO, ID_USUARIO) VALUES (" & ListView2.ListItems(1) & ", " & Replace(Text4.Text, ",", "") & ", '" & Format(Date, "dd/mm/yyyy") & "', 'N', " & VarMen.Text1(0).Text & ");"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT ID_VALE FROM VALE_CAJA ORDER BY ID_VALE DESC"
            Set tRs = cnn.Execute(sBuscar)
            For Con = 1 To ListView2.ListItems.Count
                sBuscar = "INSERT INTO VALE_CAJA_DETALLE (ID_VALE, ID_PRODUCTO, CANTIDAD, PRECIO_VENTA) VALUES (" & tRs.Fields("ID_VALE") & ", '" & ListView1.ListItems(Con).SubItems(2) & "', " & Replace(ListView1.ListItems(Con).SubItems(4), ",", "") & ", " & Replace(ListView1.ListItems(Con).SubItems(5), ",", "") & ");"
                cnn.Execute (sBuscar)
                sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ListView1.ListItems(Con).SubItems(2) & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                Set tRs1 = cnn.Execute(sBuscar)
                If tRs1.EOF And tRs1.BOF Then
                    sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & ListView1.ListItems(Con).SubItems(2) & "', '" & Replace(ListView1.ListItems(Con).SubItems(4), ",", "") & "', " & VarMen.Text4(0).Text & ");"
                    cnn.Execute (sBuscar)
                Else
                    sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & Replace(ListView1.ListItems(Con).SubItems(4), ",", "") & " WHERE ID_PRODUCTO = '" & ListView1.ListItems(Con).SubItems(2) & "' AND SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                    cnn.Execute (sBuscar)
                End If
            Next
            Imprimir
        Else
            MsgBox "LA VENTA YA TIENE UN VALE DE CAJA REALIZADO EL DIA" & tRs.Fields("FECHA") & "!", vbInformation, "SACC"
        End If
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IndAgr = Item.Index
    Itm = Item
    SubItm1 = Item.SubItems(1)
    SubItm2 = Item.SubItems(2)
    SubItm3 = Item.SubItems(3)
    SubItm4 = Item.SubItems(4)
    SubItm5 = Item.SubItems(5)
    Text2.Text = Item.SubItems(2)
    Text3.Text = Item.SubItems(4)
    Text5.Text = Item.SubItems(1)
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    EliItm = Item.Index
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command1.Value = True
    End If
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Imprimir()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT ID_VALE FROM VALE_CAJA ORDER BY ID_VALE DESC"
    Set tRs = cnn.Execute(sBuscar)
    sBuscar = "SELECT ID_VALE, ID_VENTA, IMPORTE, FECHA, ID_PRODUCTO, CANTIDAD, PRECIO_VENTA FROM VsValeCaja WHERE ID_VALE =" & tRs.Fields("ID_VALE")
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "FECHA : " & tRs.Fields("FECHA")
        Printer.Print "SUCURSAL : " & VarMen.Text4(0).Text
        Printer.Print "TELEFONO SUCURSAL : " & VarMen.Text4(5).Text
        Printer.Print "No. DE VALE DE CAJA : " & tRs.Fields("ID_VALE")
        Printer.Print "No. DE VENTA : " & tRs.Fields("ID_VENTA")
        Printer.Print "ATENDIDO POR : " & VarMen.Text1(1).Text
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                          VALE DE CAJA"
        Printer.Print "--------------------------------------------------------------------------------"
        Dim POSY As Integer
        POSY = 2200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 1300
        Printer.Print "P./U."
        Printer.CurrentY = POSY
        Printer.CurrentX = 3000
        Printer.Print "Cant."
        If Not (tRs.BOF And tRs.EOF) Then
            Do While Not tRs.EOF
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print tRs.Fields("ID_PRODUCTO")
                Printer.CurrentY = POSY
                Printer.CurrentX = 1900
                Printer.Print tRs.Fields("PRECIO_VENTA")
                Printer.CurrentY = POSY
                Printer.CurrentX = 2900
                Printer.Print tRs.Fields("CANTIDAD")
                tRs.MoveNext
            Loop
        End If
        tRs.MoveFirst
        Printer.CurrentY = POSY + 200
        Printer.Print "TOTAL : " & tRs.Fields("IMPORTE")
        Printer.CurrentY = POSY + 400
        Printer.Print ""
        Printer.CurrentY = POSY + 600
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.CurrentY = POSY + 800
        Printer.Print "               GRACIAS POR SU COMPRA"
        Printer.CurrentY = POSY + 1000
        Printer.Print "           PRODUCTO 100% GARANTIZADO"
        Printer.CurrentY = POSY + 1200
        Printer.Print "                APLICA RESTRICCIONES"
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.EndDoc
    End If
End Sub
