VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Pedidos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PEDIDO DE SURTIDO DE ALMACEN"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10560
      TabIndex        =   27
      Top             =   6120
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
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Pedidos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Pedidos.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   -2147483644
      TabCaption(0)   =   " "
      TabPicture(0)   =   "Pedidos.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ListView1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ListView2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DTPicker1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text2(4)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text2(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Check1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Command2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text2(2)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Command1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text2(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text2(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Option2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Option1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   0
         Top             =   840
         Width           =   5415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Descripcion"
         Height          =   195
         Left            =   7200
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   7200
         TabIndex        =   2
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   5
         Top             =   3360
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   7
         Top             =   3720
         Width           =   6495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Pedir"
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
         Left            =   9120
         Picture         =   "Pedidos.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   6120
         TabIndex        =   6
         Top             =   3360
         Width           =   1815
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
         Left            =   9120
         Picture         =   "Pedidos.frx":4DDA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pedido para uso personal "
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   6600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
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
         Left            =   7920
         Picture         =   "Pedidos.frx":77AC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6720
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   3
         Left            =   8160
         TabIndex        =   16
         Top             =   3720
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   6240
         Width           =   8895
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   4
         Left            =   8400
         TabIndex        =   15
         Top             =   3720
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   8880
         TabIndex        =   3
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51183617
         CurrentDate     =   38727
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2055
         Left            =   120
         TabIndex        =   9
         Top             =   4080
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3625
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3625
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
      Begin VB.Label Label1 
         Caption         =   "Buscar Producto"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
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
         TabIndex        =   24
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Agente"
         Height          =   255
         Left            =   4680
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
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
         Left            =   5280
         TabIndex        =   22
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label6 
         Caption         =   "Clave"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   5280
         TabIndex        =   19
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "(no contar como inventario de venta)"
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   6840
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Comentarios :"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   6240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim ItmInd As Integer
Private Sub Command1_Click()
    Me.Command1.Enabled = True
    Dim tLi As ListItem
    Set tLi = ListView2.ListItems.Add(, , Text2(0).Text)
    tLi.SubItems(1) = Text2(1).Text
    tLi.SubItems(2) = Text2(2).Text
    tLi.SubItems(3) = Text2(3).Text
    tLi.SubItems(4) = Text2(4).Text
    Text2(0).Text = ""
    Text2(1).Text = ""
    Text2(2).Text = ""
    Me.Command2.Enabled = True
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tip As String
    Dim NRegistros As Integer
    Dim POSY As Integer
    Dim Con As Integer
    Dim NumeroRegistros As Integer
    Dim Conta As Integer
    If Check1.Value = 0 Then
        tip = "D"
    Else
        tip = "I"
    End If
    If Text3.Text <> "" Then
        sBuscar = "INSERT INTO PEDIDO (SUCURSAL, PIDIO, FECHA, TIPO, COMENTARIO) VALUES ('" & Label3.Caption & "', '" & VarMen.Text1(0).Text & "', DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), '" & tip & "', '" & Text3.Text & "')"
        cnn.Execute (sBuscar)
        If MsgBox("DESEA IMPRIMIR EL COMPROBANTE DE PEDIDO?", vbYesNo, "SACC") = vbYes Then
            sBuscar = "SELECT ID_PEDIDO FROM PEDIDO WHERE SUCURSAL = '" & Label3.Caption & "' ORDER BY ID_PEDIDO DESC"
            Set tRs = cnn.Execute(sBuscar)
            NumeroRegistros = ListView2.ListItems.Count
            Printer.Print "   " & VarMen.Text5(0).Text
            Printer.Print "         SURTIR SUCURSAL"
            Printer.Print "        LISTA DE PRODUCTOS "
            Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
            POSY = 1400
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print "Producto"
            Printer.CurrentY = POSY
            Printer.CurrentX = 1900
            Printer.Print "Cant."
            Printer.CurrentY = POSY
            Printer.CurrentX = 2400
            Printer.Print "SUCURSAL"
            For Conta = 1 To NumeroRegistros
                sBuscar = "INSERT INTO DETALLE_PEDIDO (ID_PEDIDO, CANTIDAD, ID_PRODUCTO, ENTREGADO, DESCRIPCION, ALMACEN, MARCA) VALUES ('" & tRs.Fields("ID_PEDIDO") & "', '" & ListView2.ListItems(Conta).SubItems(2) & "', '" & ListView2.ListItems(Conta) & "', 0, '" & ListView2.ListItems(Conta).SubItems(1) & "', '" & ListView2.ListItems(Conta).SubItems(3) & "', '" & ListView2.ListItems(Conta).SubItems(4) & "')"
                cnn.Execute (sBuscar)
                If POSY > 16000 Then
                    Printer.NewPage
                    Printer.Print "     " & VarMen.Text5(0).Text
                    Printer.Print "         SURTIR SUCURSAL"
                    Printer.Print "        LISTA DE PRODUCTOS "
                    Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
                    POSY = 1400
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 100
                    Printer.Print "Producto"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 1900
                    Printer.Print "Cant."
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 2400
                    Printer.Print "SUCURSAL"
                    POSY = POSY + 200
                Else
                    POSY = POSY + 200
                End If
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print ListView2.ListItems.Item(Conta)
                Printer.CurrentY = POSY
                Printer.CurrentX = 1900
                Printer.Print ListView2.ListItems(Conta).SubItems(2)
                Printer.CurrentY = POSY
                Printer.CurrentX = 2400
                Printer.Print VarMen.Text4(0).Text
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 2400
                Printer.Print "---------------------------------------------------------"
            Next Conta
            Printer.Print ""
            Printer.Print "FIN DEL LISTADO"
            Printer.EndDoc
            ListView1.ListItems.Clear
            ListView2.ListItems.Clear
            Text1.Text = ""
            Text2(0).Text = ""
            Text2(1).Text = ""
            Text2(2).Text = ""
            Check1.Value = 0
            Me.Command2.Enabled = False
        Else
            sBuscar = "SELECT ID_PEDIDO FROM PEDIDO WHERE SUCURSAL = '" & Label3.Caption & "' ORDER BY ID_PEDIDO DESC"
            Set tRs = cnn.Execute(sBuscar)
            NumeroRegistros = ListView2.ListItems.Count
            For Conta = 1 To NumeroRegistros
                sBuscar = "INSERT INTO DETALLE_PEDIDO (ID_PEDIDO, CANTIDAD, ID_PRODUCTO, ENTREGADO, DESCRIPCION, ALMACEN, MARCA) VALUES ('" & tRs.Fields("ID_PEDIDO") & "', '" & ListView2.ListItems(Conta).SubItems(2) & "', '" & ListView2.ListItems(Conta) & "', 0, '" & ListView2.ListItems(Conta).SubItems(1) & "', '" & ListView2.ListItems(Conta).SubItems(3) & "', '" & ListView2.ListItems(Conta).SubItems(4) & "')"
                cnn.Execute (sBuscar)
            Next Conta
            ListView1.ListItems.Clear
            ListView2.ListItems.Clear
            Text1.Text = ""
            Text2(0).Text = ""
            Text2(1).Text = ""
            Text2(2).Text = ""
            Check1.Value = 0
            Me.Command2.Enabled = False
        End If
    Else
        MsgBox "NO PUEDE DAR DE ALTA UN PEDIDO SIN UN COMENTARIO", vbCritical, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command3_Click()
    If ItmInd > 0 Then
        ListView2.ListItems.Remove (ItmInd)
        ItmInd = 0
    Else
        MsgBox "NO HA SELECCIONADO NINGUN ARTICULO PARA ELIMINAR!", vbInformation, "SACC"
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Me.Command1.Enabled = False
    Me.Command2.Enabled = False
    Me.Command2.Enabled = False
    DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    Label3.Caption = VarMen.Text4(0).Text
    Label5.Caption = VarMen.Text1(1).Text
    Option2.Value = True
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
        .ColumnHeaders.Add , , "Clave del Producto", 3200
        .ColumnHeaders.Add , , "Descripcion", 7000
        .ColumnHeaders.Add , , "ALMACEN", 0
        .ColumnHeaders.Add , , "MARCA", 0
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 3200
        .ColumnHeaders.Add , , "Descripcion", 7000
        .ColumnHeaders.Add , , "CANTIDAD", 1500
        .ColumnHeaders.Add , , "ALMACEN", 0
        .ColumnHeaders.Add , , "MARCA", 0
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text2(0).Text = ListView1.SelectedItem
    Text2(1).Text = ListView1.SelectedItem.SubItems(1)
    Text2(3).Text = ListView1.SelectedItem.SubItems(2)
    Text2(4).Text = ListView1.SelectedItem.SubItems(3)
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2(2).SetFocus
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ItmInd = Item.Index
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        ListView1.ListItems.Clear
        If Option1.Value Then
            sBuscar = "SELECT * FROM ALMACEN3 WHERE Descripcion LIKE '%" & Text1.Text & "%' AND PEDIDO_SUCURSAL = 'S'"
        Else
            sBuscar = "SELECT * FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' AND PEDIDO_SUCURSAL = 'S'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If Not (.BOF And .EOF) Then
                .MoveFirst
                Do While Not .EOF
                    Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    tLi.SubItems(1) = .Fields("Descripcion") & ""
                    tLi.SubItems(2) = "A3"
                    tLi.SubItems(3) = .Fields("MARCA") & ""
                    .MoveNext
                Loop
                ListView1.SetFocus
            End If
        End With
        If Option1.Value Then
            sBuscar = "SELECT * FROM ALMACEN2 WHERE Descripcion LIKE '%" & Text1.Text & "%' AND PEDIDO_SUCURSAL = 'S'"
        Else
            sBuscar = "SELECT * FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' AND PEDIDO_SUCURSAL = 'S'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If Not (.BOF And .EOF) Then
                .MoveFirst
                Do While Not .EOF
                    Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    tLi.SubItems(1) = .Fields("Descripcion") & ""
                    tLi.SubItems(2) = "A2"
                    tLi.SubItems(3) = .Fields("MARCA") & ""
                    .MoveNext
                Loop
                ListView1.SetFocus
            End If
        End With
        
        If Option1.Value Then
            sBuscar = "SELECT * FROM ALMACEN1 WHERE Descripcion LIKE '%" & Text1.Text & "%' AND PEDIDO_SUCURSAL = 'S'"
        Else
            sBuscar = "SELECT * FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' AND PEDIDO_SUCURSAL = 'S'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If Not (.BOF And .EOF) Then
                .MoveFirst
                Do While Not .EOF
                    Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    tLi.SubItems(1) = .Fields("Descripcion") & ""
                    tLi.SubItems(2) = "A1"
                    tLi.SubItems(3) = .Fields("MARCA") & ""
                    .MoveNext
                Loop
                ListView1.SetFocus
            End If
        End With
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text2_Change(Index As Integer)
    If Text2(1).Text <> "" And Text2(2).Text <> "" Then
        Me.Command1.Enabled = True
    End If
End Sub
Private Sub Text2_GotFocus(Index As Integer)
    Text2(Index).BackColor = &HFFE1E1
    Text2(Index).SetFocus
    Text2(Index).SelStart = 0
    Text2(Index).SelLength = Len(Text2(Index).Text)
End Sub
Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 2 Then
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii = 13 Then
        Command1.Value = True
        Text1.SetFocus
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End If
End Sub
Private Sub Text2_LostFocus(Index As Integer)
    Text2(Index).BackColor = &H80000005
End Sub
