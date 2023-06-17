VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmProduccionMaxMin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ordenes por Maximos y Minimos"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7800
      TabIndex        =   4
      Top             =   3000
      Width           =   975
      Begin VB.Label Label7 
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image5 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmProduccionMaxMin.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmProduccionMaxMin.frx":030A
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7800
      TabIndex        =   2
      Top             =   4200
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmProduccionMaxMin.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "FrmProduccionMaxMin.frx":1FD6
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Pendientes"
      TabPicture(0)   =   "FrmProduccionMaxMin.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Ordenes Abiertas"
      TabPicture(1)   =   "FrmProduccionMaxMin.frx":40D4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView2"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView ListView2 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   8281
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4680
         TabIndex        =   9
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4800
         Width           =   2655
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7011
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
      Begin VB.Label Label2 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   4800
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   4800
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmProduccionMaxMin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
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
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Producto", 3500
        .ColumnHeaders.Add , , "Cantidad", 1500
    End With
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Comanda", 1500
        .ColumnHeaders.Add , , "Producto", 3500
        .ColumnHeaders.Add , , "Cantidad", 1500
    End With
    MaxMin
End Sub
Private Sub MaxMin()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT ALMACEN3.ID_PRODUCTO, ALMACEN3.C_MAXIMA - ISNULL(EXISTENCIAS.CANTIDAD, 0) AS FALTANTE FROM ALMACEN3 LEFT OUTER JOIN EXISTENCIAS ON ALMACEN3.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE (ALMACEN3.C_MAXIMA > 0) AND (EXISTENCIAS.SUCURSAL = 'BODEGA') AND (ALMACEN3.C_MINIMA >= EXISTENCIAS.CANTIDAD) AND (ALMACEN3.ID_PRODUCTO NOT IN (SELECT ID_PRODUCTO From COMANDAS_DETALLES_2 WHERE (CLASIFICACION = 'P') AND (ESTADO_ACTUAL NOT IN ('C', 'I', 'N', 'L', '0')))) AND (ALMACEN3.ID_PRODUCTO LIKE '%COMAP' OR ALMACEN3.ID_PRODUCTO LIKE '%CAMAP') AND ALMACEN3.TIPO = 'COMPUESTO' UNION SELECT ID_PRODUCTO, (C_MAXIMA) AS FALTANTE FROM ALMACEN3 WHERE ID_PRODUCTO NOT IN (SELECT ID_PRODUCTO FROM EXISTENCIAS WHERE SUCURSAL = 'BODEGA') AND C_MAXIMA >0 AND ALMACEN3.TIPO = 'COMPUESTO' AND (ALMACEN3.ID_PRODUCTO LIKE '%COMAP' OR ALMACEN3.ID_PRODUCTO LIKE '%CAMAP')"
    'sBuscar = "SELECT ALMACEN3.ID_PRODUCTO, ALMACEN3.C_MAXIMA - ISNULL(EXISTENCIAS.CANTIDAD, 0) AS FALTANTE FROM ALMACEN3 LEFT OUTER JOIN EXISTENCIAS ON ALMACEN3.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE (ALMACEN3.C_MAXIMA > 0) AND (EXISTENCIAS.SUCURSAL = 'BODEGA') AND (ALMACEN3.C_MINIMA >= EXISTENCIAS.CANTIDAD) AND (ALMACEN3.ID_PRODUCTO LIKE '%COMAP' OR ALMACEN3.ID_PRODUCTO LIKE '%CAMAP')"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            tLi.SubItems(1) = tRs.Fields("FALTANTE")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT ID_COMANDA, ID_PRODUCTO, CANTIDAD From COMANDAS_DETALLES_2 WHERE (CLASIFICACION = 'P') AND (ESTADO_ACTUAL NOT IN ('C', 'I', 'N', 'L', '0', 'O'))"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_COMANDA"))
            tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image5_Click()
On Error GoTo ManejaError
    If Text1.Text = "" Then
        MsgBox "INGRESE UN PRODUCTO PARA PODER CERRAR LA ORDEN DE PRODUCCION!", vbInformation, "SACC"
    Else
        If Text2.Text <> "" And CDbl(Text2.Text) <> 0 Then
            Dim Cont As Integer
            Dim nComanda As Integer
            Dim cTipo As String
            Dim tRs As ADODB.Recordset
            Dim sBuscar As String
            sBuscar = "INSERT INTO COMANDAS_2 (FECHA_INICIO, ID_AGENTE, ID_SUCURSAL, TIPO, COMENTARIO, SUCURSAL) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', " & VarMen.Text1(0).Text & ", " & VarMen.Text1(5).Text & ", 'P','PRODUCCION POR MAXIMOS Y MINIMOS SACC', '" & VarMen.Text4(0).Text & "')"
            cnn.Execute (sBuscar)
            DoEvents
            sBuscar = "SELECT TOP 1 ID_COMANDA FROM COMANDAS_2 ORDER BY ID_COMANDA DESC"
            Set tRs = cnn.Execute(sBuscar)
            nComanda = tRs.Fields("ID_COMANDA")
            DoEvents
            If Mid(Text1.Text, 3, 1) = "T" Then
                cTipo = "T" 'Toner
            Else
                If Mid(Text1.Text, 3, 1) = "I" Then
                    cTipo = "I" 'Tinta
                Else
                    cTipo = "X" 'Error
                End If
            End If
            sBuscar = "INSERT INTO COMANDAS_DETALLES_2 (ID_COMANDA, ARTICULO, ID_PRODUCTO, CANTIDAD, TIPO, CLASIFICACION) VALUES (" & nComanda & ", " & Cont & ", '" & Text1.Text & "', " & Text2.Text & ", '" & cTipo & "','P');"
            cnn.Execute (sBuscar)
            sBuscar = "INSERT INTO PRODPEND (ID_COMANDA, ARTICULO) VALUES (" & nComanda & ", " & Cont & ");"
            cnn.Execute (sBuscar)
            DoEvents
            Imprimir_Ticket (nComanda)
            Imprimir_Ticket (nComanda)
            Text1.Text = ""
            Text2.Text = ""
            MaxMin
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Imprimir_Ticket(cNoCom As Integer)
On Error GoTo ManejaError
    Printer.Print "        " & VarMen.Text5(0).Text
    Printer.Print "           ORDEN DE PRODUCCIÓN"
    Printer.Print "FECHA : " & Now
    Printer.Print "No. DE ORDEN DE PRODUCCCION : " & cNoCom
    Printer.Print "ORDEN HECHA POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
    Printer.Print "COMENTARIO : " & Text1.Text
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                           ORDEN DE TINTA"
    Dim Con As Integer
    Dim POSY As Integer
    POSY = 2200
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3000
    Printer.Print "Cant."
    If Mid(Text1.Text, 3, 1) = "I" Then
        POSY = POSY + 200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print Text1.Text
        Printer.CurrentY = POSY
        Printer.CurrentX = 2900
        Printer.Print Text2.Text
    End If
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                           ORDEN DE TONER"
    POSY = POSY + 600
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3000
    Printer.Print "Cant."
    If Mid(Text1.Text, 3, 1) = "T" Then
        POSY = POSY + 200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print Text1.Text
        Printer.CurrentY = POSY
        Printer.CurrentX = 2900
        Printer.Print Text2.Text
    End If
    Printer.Print ""
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print ""
    Printer.EndDoc
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item
    Text2.Text = Item.SubItems(1)
End Sub
