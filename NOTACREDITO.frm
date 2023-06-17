VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form NotaCredito 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Notas de Credito"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10200
      TabIndex        =   37
      Top             =   5760
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "NOTACREDITO.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "NOTACREDITO.frx":030A
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
         TabIndex        =   38
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10200
      TabIndex        =   35
      Top             =   4560
      Width           =   975
      Begin VB.Image Command4 
         Height          =   720
         Left            =   120
         MouseIcon       =   "NOTACREDITO.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "NOTACREDITO.frx":26F6
         Top             =   240
         Width           =   675
      End
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
         TabIndex        =   36
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "NOTACREDITO.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtIDCLIENTE"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame Frame1 
         Caption         =   "Numero de Venta"
         Height          =   1215
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   2655
         Begin VB.OptionButton Option6 
            Caption         =   "Por Factura"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Por Nota"
            Height          =   255
            Left            =   1560
            TabIndex        =   39
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   1215
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
            Left            =   1440
            Picture         =   "NOTACREDITO.frx":40D4
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos de la venta"
         Height          =   5295
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   5175
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox Text4 
            Height          =   765
            Left            =   720
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   2880
            Width           =   2415
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   3960
            TabIndex        =   4
            Top             =   2880
            Width           =   1095
         End
         Begin VB.CommandButton Command5 
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
            Picture         =   "NOTACREDITO.frx":6AA6
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   3240
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   2520
            Width           =   3015
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1455
            Left            =   120
            TabIndex        =   6
            Top             =   3720
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   2566
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
            Height          =   1695
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   2990
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
            Caption         =   "Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Motivo"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   2880
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   3240
            TabIndex        =   31
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Producto"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   2520
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sanción"
         Height          =   1695
         Left            =   5400
         TabIndex        =   24
         Top             =   5040
         Width           =   4455
         Begin VB.TextBox Text5 
            Height          =   795
            Left            =   720
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   16
            Text            =   "0.00"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Motivo"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Total"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo"
         Height          =   855
         Left            =   5400
         TabIndex        =   23
         Top             =   480
         Width           =   2535
         Begin VB.OptionButton Option1 
            Caption         =   "De Venta"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Por Sanción "
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   8
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Contar como"
         Height          =   1215
         Left            =   2880
         TabIndex        =   22
         Top             =   120
         Width           =   2415
         Begin VB.OptionButton Option2 
            Caption         =   "Sanción"
            Height          =   195
            HelpContextID   =   1
            Index           =   2
            Left            =   600
            TabIndex        =   43
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Mermas"
            Height          =   195
            HelpContextID   =   1
            Index           =   1
            Left            =   600
            TabIndex        =   42
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Existencias"
            Height          =   195
            HelpContextID   =   1
            Index           =   0
            Left            =   600
            TabIndex        =   41
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Cliente que sanciono"
         Height          =   3495
         Left            =   5400
         TabIndex        =   20
         Top             =   1440
         Width           =   4455
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   2775
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Por Nombre"
            Height          =   255
            Left            =   3000
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Por Calve"
            Height          =   255
            Left            =   3000
            TabIndex        =   12
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   3120
            Width           =   3135
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   2175
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   3836
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
         Begin VB.Label Label7 
            Caption         =   "Seleccionado"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   3120
            Width           =   975
         End
      End
      Begin VB.TextBox txtIDCLIENTE 
         Height          =   285
         Left            =   9720
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Frame Frame7 
         Caption         =   "Folio"
         Height          =   855
         Left            =   8040
         TabIndex        =   18
         Top             =   480
         Width           =   1815
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "NotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim LimCant As String
Dim DesProd As String
Dim PrecProd As String
Dim Posi As Integer
Dim CLVCLIEN As String
Dim RFCClien As String
Dim ClvVenta As String
Private Sub Command1_Click()
    BuscaVenta
End Sub
Private Sub Command4_Click()
On Error GoTo ManejaError
    Dim sqlComanda As String
    Dim tRs As ADODB.Recordset
    Dim ClvVenta As String
    If Option1(0).Value = True Then
        If Text5.Text = "" Or CDbl(Text6.Text) <= 0 Then
            MsgBox "DEBE ANOTAR UN MOTIVO A LA SANCIÓN Y UNA CANTIDAD MAYOR A CERO!", vbInformation, "SACC"
        Else
            'sqlComanda = "SELECT FOLIO, ID_CLIENTE, ID_VENTA FROM VENTAS WHERE FOLIO = '" & Text10.Text & "'"
            'Set tRs = cnn.Execute(sqlComanda)
            'If Not (tRs.EOF And tRs.BOF) Then
                Text6.Text = Replace(Text6.Text, ",", "")
            '    ClvVenta = tRs.Fields("ID_VENTA")
                sqlComanda = "INSERT INTO NOTA_CREDITO (IMPORTE, NOMBRE, TOTAL, FECHA, MOTIVOCAMBIO, ID_VENTA, ID_USUARIO, ID_CLIENTE, APLICADA) VALUES (" & Text6.Text & ", '" & Text2.Text & "', " & Text6.Text & ", '" & Format(Date, "dd/mm/yyyy") & "', '" & Text5.Text & "', 0, '" & VarMen.Text1(0).Text & "', " & CLVCLIEN & ", 'N');"
                cnn.Execute (sqlComanda)
                Text1.Text = ""
                Text2.Text = ""
                Text3.Text = ""
                Text4.Text = ""
                Text5.Text = ""
                Text6.Text = "0.00"
                Text7.Text = ""
                Text8.Text = ""
                Text9.Text = ""
                Text10.Text = ""
                ListView1.ListItems.Clear
                ListView2.ListItems.Clear
                ListView3.ListItems.Clear
                Option1(0).Enabled = True
                Option1(1).Enabled = True
            'Else
            '    MsgBox "EL NUMERO DE FOLIO YA EXISTE, NO SE PUEDE REGISTRAR LA NOTA!", vbInformation, "SACC"
            '    Exit Sub
            'End If
        End If
    Else
        If ListView2.ListItems.Count = 0 Then
            MsgBox "DEBE AGREGAR ARTICULOS A LA NOTA DE CREDITO!", vbInformation, "SACC"
        Else
            Dim tRs1 As ADODB.Recordset
            Dim CanEx As String
            Dim PreProd As String
            Dim NumeroRegistros As Integer
            Dim Conta As Integer
            Dim IdNota As String
            Text6.Text = Replace(Text6.Text, ",", "")
            If Option6.Value Then
                sqlComanda = "SELECT ID_VENTA, ID_CLIENTE FROM VENTAS WHERE FOLIO = '" & ClvVenta & "'"
                Set tRs1 = cnn.Execute(sqlComanda)
                sqlComanda = "INSERT INTO NOTA_CREDITO (IMPORTE, NOMBRE, TOTAL, FECHA, MOTIVOCAMBIO, ID_VENTA, ID_USUARIO, ID_CLIENTE, FOLIO) VALUES (" & CDbl(Text6.Text) * (CDbl(CDbl(VarMen.Text4(7).Text) / 100) + 1) & ", '" & Text3.Text & "', " & CDbl(Text6.Text) * (CDbl(CDbl(VarMen.Text4(7).Text) / 100) + 1) & ", '" & Format(Date, "dd/mm/yyyy") & "', '" & Text5.Text & "', 0, '" & VarMen.Text1(0).Text & "', " & tRs1.Fields("ID_CLIENTE") & ", '" & Text10.Text & "');"
                cnn.Execute (sqlComanda)
            Else
                sqlComanda = "INSERT INTO NOTA_CREDITO (IMPORTE, NOMBRE, TOTAL, FECHA, MOTIVOCAMBIO, ID_VENTA, ID_USUARIO, ID_CLIENTE, FOLIO) VALUES (" & CDbl(Text6.Text) * (CDbl(CDbl(VarMen.Text4(7).Text) / 100) + 1) & ", '" & Text3.Text & "', " & CDbl(Text6.Text) * (CDbl(CDbl(VarMen.Text4(7).Text) / 100) + 1) & ", '" & Format(Date, "dd/mm/yyyy") & "', '" & Text5.Text & "', 0, '" & VarMen.Text1(0).Text & "', " & tRs1.Fields("ID_CLIENTE") & ", '" & Text10.Text & "');"
                cnn.Execute (sqlComanda)
            End If
            sqlComanda = "SELECT ID_NOTA FROM NOTA_CREDITO WHERE FECHA = '" & Format(Date, "dd/mm/yyyy") & "'ORDER BY ID_NOTA DESC"
            Set tRs = cnn.Execute(sqlComanda)
            If Not (tRs.EOF And tRs.BOF) Then
                IdNota = tRs.Fields("ID_NOTA")
                NumeroRegistros = ListView2.ListItems.Count
                For Conta = 1 To NumeroRegistros
                    CanEx = Format(CDbl(ListView2.ListItems.Item(Conta).SubItems(3)), "###,###,##0.00")
                    CanEx = Replace(CanEx, ",", "")
                    PreProd = Format(CDbl(ListView2.ListItems.Item(Conta).SubItems(2)), "###,###,##0.00")
                    PreProd = Replace(PreProd, ",", "")
                    sqlComanda = "INSERT INTO NOTA_CREDITO_PRODUCTO (ID_NOTA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO) VALUES (" & IdNota & ", '" & ListView2.ListItems.Item(Conta) & "', '" & ListView2.ListItems.Item(Conta).SubItems(1) & "', " & CanEx & ", " & PreProd & ");"
                    cnn.Execute (sqlComanda)
                    If ListView2.ListItems.Item(Conta).SubItems(5) = "EXISTENCIA" Then
                        sqlComanda = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ListView2.ListItems.Item(Conta) & "' AND SUCURSAL = 'BODEGA'"
                        Set tRs1 = cnn.Execute(sqlComanda)
                        If (tRs1.EOF And tRs1.BOF) Then
                            CanEx = Format(CDbl(ListView2.ListItems.Item(Conta).SubItems(3)), "###,###,##0.00")
                            CanEx = Replace(CanEx, ",", "")
                            sqlComanda = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & ListView2.ListItems.Item(Conta) & "', " & CanEx & ", 'BODEGA');"
                            cnn.Execute (sqlComanda)
                        Else
                            CanEx = Format(CDbl(ListView2.ListItems.Item(Conta).SubItems(3)) + CDbl(tRs1.Fields("CANTIDAD")), "###,###,##0.00")
                            CanEx = Replace(CanEx, ",", "")
                            sqlComanda = "UPDATE EXISTENCIAS SET CANTIDAD = " & CanEx & " WHERE ID_PRODUCTO = '" & ListView2.ListItems.Item(Conta) & "' AND SUCURSAL = 'BODEGA'"
                            Set tRs = cnn.Execute(sqlComanda)
                        End If
                    Else
                        sqlComanda = "INSERT INTO MERMAS (ID_NOTA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, FECHA) VALUES (" & IdNota & ", '" & ListView2.ListItems.Item(Conta) & "', '" & ListView2.ListItems.Item(Conta).SubItems(1) & "', " & CanEx & ", '" & Format(Date, "dd/mm/yyyy") & "');"
                        cnn.Execute (sqlComanda)
                    End If
                Next Conta
            Else
                MsgBox " Ocurrio un error al guardar la nota, Favor de reportar a Soporte", vbCritical, "SACC"
            End If
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text5.Text = ""
            Text6.Text = "0.00"
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            ListView1.ListItems.Clear
            ListView2.ListItems.Clear
            ListView3.ListItems.Clear
            Option1(0).Enabled = True
            Option1(1).Enabled = True
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command5_Click()
    AgregarProd
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Text1.Enabled = True
    Command1.Enabled = False
    Text2.Enabled = True
    ListView1.Enabled = True
    Text4.Enabled = True
    Text8.Enabled = True
    Command5.Enabled = False
    ListView2.Enabled = True
    Option2(0).Enabled = True
    Option2(1).Enabled = True
    Option2(2).Enabled = True
    Text7.Enabled = False
    Option3.Enabled = False
    Option4.Enabled = False
    ListView3.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
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
        .ColumnHeaders.Add , , "Clave del Producto", 2100
        .ColumnHeaders.Add , , "Descripcion", 7000
        .ColumnHeaders.Add , , "PRECIO", 1500
        .ColumnHeaders.Add , , "CANTIDAD", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 2100
        .ColumnHeaders.Add , , "Descripcion", 7000
        .ColumnHeaders.Add , , "PRECIO", 1500
        .ColumnHeaders.Add , , "CANTIDAD", 1500
        .ColumnHeaders.Add , , "MOTIVO", 7500
        .ColumnHeaders.Add , , "ALMACEN", 1500
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID CLIENTE", 0
        .ColumnHeaders.Add , , "NOMBRE", 7000
        .ColumnHeaders.Add , , "RFC", 2500
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
    Text3.Text = Item
    DesProd = Item.SubItems(1)
    PrecProd = Item.SubItems(2)
    LimCant = Item.SubItems(3)
    Posi = Item.Index
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    CLVCLIEN = Item
    Text9.Text = Item.SubItems(1)
    RFCClien = Item.SubItems(2)
    txtIDCliente = Item
End Sub
Private Sub Option1_Click(Index As Integer)
    If ListView2.ListItems.Count = 0 Then
        If Option1(0).Value = True Then
            Text1.Enabled = False
            Command1.Enabled = False
            Text2.Enabled = False
            ListView1.Enabled = False
            Text4.Enabled = False
            Text8.Enabled = False
            ListView2.Enabled = False
            Option2(0).Enabled = False
            Option2(1).Enabled = False
            Option2(2).Enabled = False
            Text7.Enabled = True
            Option3.Enabled = True
            Option4.Enabled = True
            ListView3.Enabled = True
            Text5.Enabled = True
            Text6.Enabled = True
        Else
            Text1.Enabled = True
            If Text1.Text <> "" Then
                Command1.Enabled = True
            Else
                Command1.Enabled = False
            End If
            Text2.Enabled = True
            ListView1.Enabled = True
            Text4.Enabled = True
            Text8.Enabled = True
            ListView2.Enabled = True
            Option2(0).Enabled = True
            Option2(1).Enabled = True
            Option2(2).Enabled = True
            Text7.Enabled = False
            Option3.Enabled = False
            Option4.Enabled = False
            ListView3.Enabled = False
            Text5.Enabled = False
            Text6.Enabled = False
        End If
    End If
    Text6.Text = "0.00"
    Text5.Text = ""
    Text9.Text = ""
End Sub
Private Sub Option3_Click()
    Text7.Text = ""
End Sub
Private Sub Option4_Click()
    Text7.Text = ""
End Sub
Private Sub Text1_Change()
    If Text1.Text <> "" Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
    End If
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text1.Text <> "" Then
        BuscaVenta
    End If
    Dim Valido As String
    If Option6.Value Then
        Valido = "1234567890ABDEFGHIJKLMNOPQRSTUVWXYX- "
    Else
        Valido = "1234567890"
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text3_Change()
    If Text4.Text <> "" And Text8.Text <> "" And Text3.Text <> "" Then
        Me.Command5.Enabled = True
    Else
        Me.Command5.Enabled = False
    End If
End Sub
Private Sub Text4_Change()
    If Text4.Text <> "" And Text8.Text <> "" And Text3.Text <> "" Then
        Me.Command5.Enabled = True
    Else
        Me.Command5.Enabled = False
    End If
End Sub
Private Sub Text4_GotFocus()
    Text4.BackColor = &HFFE1E1
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text4.Text <> "" And Text8.Text <> "" And Text3.Text <> "" Then
        AgregarProd
    End If
End Sub
Private Sub Text4_LostFocus()
    Text4.BackColor = &H80000005
End Sub
Private Sub Text5_GotFocus()
    Text5.BackColor = &HFFE1E1
End Sub
Private Sub Text5_LostFocus()
    Text5.BackColor = &H80000005
End Sub
Private Sub Text6_GotFocus()
    Text6.BackColor = &HFFE1E1
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text6_LostFocus()
    Text6.BackColor = &H80000005
End Sub
Private Sub Text7_GotFocus()
    Text7.BackColor = &HFFE1E1
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And Text7.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        Dim sBus As String
        If Me.Option3.Value = True Then
            sBus = "SELECT ID_CLIENTE, NOMBRE, RFC FROM CLIENTE WHERE NOMBRE LIKE '%" & Text7.Text & "%'"
        Else
            sBus = "SELECT ID_CLIENTE, NOMBRE, RFC FROM CLIENTE WHERE ID_CLIENTE = " & Text7.Text
        End If
        Set tRs = cnn.Execute(sBus)
        ListView3.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            With tRs
                Do While Not .EOF
                    Set tLi = ListView3.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & ""
                    If Not IsNull(.Fields("RFC")) Then tLi.SubItems(2) = .Fields("RFC") & ""
                    .MoveNext
                Loop
            End With
        End If
    End If
    If Option4.Value = True Then
        Dim Valido As String
        Valido = "1234567890"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text7_LostFocus()
    Text7.BackColor = &H80000005
End Sub
Private Sub Text8_Change()
    If Text4.Text <> "" And Text8.Text <> "" And Text3.Text <> "" Then
        Me.Command5.Enabled = True
    Else
        Me.Command5.Enabled = False
    End If
End Sub
Private Sub Text8_GotFocus()
    Text8.BackColor = &HFFE1E1
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text4.Text <> "" And Text8.Text <> "" And Text3.Text <> "" Then
        AgregarProd
    End If
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub BuscaVenta()
On Error GoTo ManejaError
    ClvVenta = Text1.Text
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBus As String
    If Option6.Value Then
        sBus = "SELECT NOMBRE, ID_CLIENTE, ID_VENTA FROM VENTAS WHERE FOLIO = '" & Text1.Text & "' AND FACTURADO = 1"
    Else
        sBus = "SELECT NOMBRE, ID_CLIENTE, ID_VENTA FROM VENTAS WHERE ID_VENTA = " & Text1.Text & " AND FACTURADO = 1"
    End If
    ListView1.ListItems.Clear
    Set tRs = cnn.Execute(sBus)
    If Not (tRs.EOF And tRs.BOF) Then
        Text2.Text = tRs.Fields("NOMBRE")
        txtIDCliente.Text = tRs.Fields("ID_CLIENTE")
        sBus = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_VENTA, CANTIDAD FROM VENTAS_DETALLE WHERE ID_VENTA = " & tRs.Fields("ID_VENTA")
        Set tRs = cnn.Execute(sBus)
        If Not (tRs.EOF And tRs.BOF) Then
            With tRs
                Do While Not .EOF
                    Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                    If Not IsNull(.Fields("PRECIO_VENTA")) Then tLi.SubItems(2) = .Fields("PRECIO_VENTA") & ""
                    If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(3) = .Fields("CANTIDAD") & ""
                    .MoveNext
                Loop
            End With
        End If
    Else
        MsgBox "NO SE ENCUENTRA REGISTRO DE LA VENTA, PUDO SER ELIMINADA O AUN NO ESTA REGISTRADA!", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub AgregarProd()
    If Text8.Text > LimCant Then
        MsgBox "LA CANTIDAD NO PUEDE SER MAYOR A LA DE LA COMPRA!", vbInformation, "SACC"
    Else
        Option1(0).Enabled = False
        Option1(1).Enabled = True
        Dim tLi As ListItem
        Set tLi = ListView2.ListItems.Add(, , Text3.Text)
        tLi.SubItems(1) = DesProd
        tLi.SubItems(2) = PrecProd
        tLi.SubItems(3) = Text8.Text
        tLi.SubItems(4) = Text4.Text
        If Option2(1).Value = True Then
            tLi.SubItems(5) = "MERMA"
        Else
            If Option2(2).Value = True Then
                tLi.SubItems(5) = "SANCION"
            Else
                tLi.SubItems(5) = "EXISTENCIA"
            End If
        End If
        ListView1.ListItems.Item(Posi).SubItems(3) = Format(CDbl(ListView1.ListItems.Item(Posi).SubItems(3)) - CDbl(Text8.Text), "###,###,##0.00")
        Text6.Text = Format(CDbl(Text6.Text) + (CDbl(Text8.Text) * CDbl(PrecProd)), "###,###,##0.00")
        Text3.Text = ""
        Text8.Text = ""
    End If
End Sub
Private Sub Text8_LostFocus()
    Text8.BackColor = &H80000005
End Sub
