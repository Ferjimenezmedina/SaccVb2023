VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSalidaInvProd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salida de Material Almacen"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   2040
      TabIndex        =   17
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmSalidaInvProd.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTitulo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListView5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListView4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ListView2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ListView3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.CommandButton Command3 
         Caption         =   "Guardar"
         Enabled         =   0   'False
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
         Left            =   3360
         Picture         =   "frmSalidaInvProd.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7320
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   5
         Top             =   5040
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         Enabled         =   0   'False
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
         Left            =   6840
         Picture         =   "frmSalidaInvProd.frx":29EE
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4980
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Quitar"
         Enabled         =   0   'False
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
         Left            =   6960
         Picture         =   "frmSalidaInvProd.frx":53C0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Top             =   2520
         Width           =   4455
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2295
         Left            =   240
         TabIndex        =   2
         Top             =   2520
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   4048
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1575
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2778
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   1575
         Left            =   240
         TabIndex        =   7
         Top             =   5640
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   2778
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   2055
         Left            =   4080
         TabIndex        =   4
         Top             =   2880
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Detalle de la produccion"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Juego de Reparacion de"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   4080
         TabIndex        =   21
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2160
         TabIndex        =   20
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Productos a descontar de inventario"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   5400
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Producto"
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   2280
         Width           =   2535
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   960
      TabIndex        =   15
      Top             =   6720
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmSalidaInvProd.frx":7D92
         MousePointer    =   99  'Custom
         Picture         =   "frmSalidaInvProd.frx":809C
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
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
      Begin VB.Label Label6 
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblID4 
         Height          =   135
         Left            =   10560
         TabIndex        =   13
         Top             =   6600
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblID3 
         Height          =   135
         Left            =   6840
         TabIndex        =   12
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   6720
         TabIndex        =   11
         Top             =   6120
         Width           =   1815
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   8281
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "frmSalidaInvProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    Dim tRs As ADODB.Recordset
    Dim sBus As String
    Dim tLi As ListItem
    If InStr(1, lblID3.Caption, "+") > 0 Then
        lblID3.Caption = Replace(lblID3.Caption, "+", "0")
        If CDbl(Text2.Text) <= CDbl(lblID3.Caption) Then
            Set tLi = ListView4.ListItems.Add(, , Label3.Caption)
            tLi.SubItems(1) = Text2.Text
        Else
            MsgBox "LA CANTIDAD SOLICITADA ES MAYOR A LA EXISTENCIA ACTUAL"
        End If
    Else
        If CDbl(Text2.Text) Mod CDbl(ListView3.ListItems.Item(lblID3.Caption).SubItems(1)) > 0 Then
            If CDbl(Text2.Text) <= CDbl(ListView3.ListItems.Item(lblID3.Caption).SubItems(2)) Then
                If CDbl(Text2.Text) <= CDbl(ListView3.ListItems.Item(lblID3.Caption).SubItems(3)) Then
                    Set tLi = ListView4.ListItems.Add(, , Label3.Caption)
                        tLi.SubItems(1) = Text2.Text
                Else
                    MsgBox "LA CANTIDAD SOLICITADA ES MAYOR A LA EXISTENCIA ACTUAL"
                End If
            Else
                MsgBox "LA CANTIDAD SOLICITADA ES MAYOR A LA REQUERIDA"
            End If
        Else
            MsgBox "LA CANTIDAD DEBE SER MULTIPLO DE LA CANTIDAD INDIVIDUAL NESESARIA PARA EL JUEGO DE REPARACION"
        End If
    End If
    If ListView4.ListItems.Count > 0 Then
        Command3.Enabled = True
        Command4.Enabled = True
    End If
End Sub
Private Sub Command3_Click()
    Dim sBus As String
    Dim Cont As Integer
    Dim Sel As Integer
    'tiene q checar existencias
    Sel = ListView1.SelectedItem
    For Cont = 1 To ListView4.ListItems.Count
        sBus = "UPDATE EXISTENCIAS CANTIDAD = CANTIDAD - " & ListView4.ListItems.Item(Cont).SubItems(1) & " WHERE ID_PRODUCTO = '" & ListView4.ListItems.Item(Cont) & "' AND SUCURSAL = 'BODEGA'"
        cnn.Execute (sBus)
        sBus = "INSERT INTO SALIDASINVPROD (ID_PRODUCCION, ID_PRODUCTO, CANTIDAD) VALUES (" & ListView1.ListItems.Item(Sel) & ", '" & ListView4.ListItems.Item(Cont) & "', " & ListView4.ListItems.Item(Cont).SubItems(1) & ");"
        cnn.Execute (sBus)
    Next Cont
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
    ListView4.ListItems.Clear
    ListView5.ListItems.Clear
End Sub
Private Sub Command4_Click()
    If Label4.Caption <> "" Then
        ListView4.ListItems.Remove (lblID4.Caption)
    End If
    If ListView4.ListItems.Count = 0 Then
        Command3.Enabled = False
        Command4.Enabled = False
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Dim tRs As ADODB.Recordset
    Dim sBus As String
    Dim tLi As ListItem
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "#PRODUCCION", 1200
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "PRODUCTO", 1200
        .ColumnHeaders.Add , , "CANTIDAD", 1200
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "PRODUCTO", 1200
        .ColumnHeaders.Add , , "CANTIDAD INDIV.", 1200
        .ColumnHeaders.Add , , "CANTIDAD", 1200
        .ColumnHeaders.Add , , "EXISTENCIA", 1200
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "PRODUCTO", 1200
        .ColumnHeaders.Add , , "CANTIDAD", 1200
    End With
    With ListView5
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "PRODUCTO", 1200
        .ColumnHeaders.Add , , "DESCRIP", 2200
        .ColumnHeaders.Add , , "CANTIDAD", 1200
    End With
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    ListView1.ListItems.Clear
    sBus = "SELECT * FROM PRODUCCION WHERE EFECTUADO = 'N'"
    Set tRs = cnn.Execute(sBus)
    If tRs.EOF And tRs.BOF Then
        MsgBox "No hay producciones pendientes"
    Else
        With tRs
            Do While Not (.EOF)
                Set tLi = lvwGarantia.ListItems.Add(, , .Fields("ID") & "")
                .MoveNext
            Loop
        End With
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim tRs As ADODB.Recordset
    Dim sBus As String
    Dim tLi As ListItem
    ListView2.ListItems.Clear
    sBus = "SELECT * FROM PRODUCCION_ENTRADAS WHERE ID_PRODUCCION = " & Item
    Set tRs = cnn.Execute(sBus)
    If tRs.EOF And tRs.BOF Then
        MsgBox "No hay productos pendientes"
    Else
        With tRs
            Do While Not (.EOF)
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    tLi.SubItems(1) = .Fields("CANTIDAD")
                .MoveNext
            Loop
        End With
        Label6.Caption = Item
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim tRs As ADODB.Recordset
    Dim sBus As String
    Dim tLi As ListItem
    Dim Clave As String
    Clave = Replace(Item, "CAMAP", "REMA")
    ListView3.ListItems.Clear
    sBus = "SELECT J.ID_PRODUCTO,J.CANTIDAD, V.CANTIDAD AS EXIST FROM JUEGO_REPARACION AS J LEFT JOIN VSEXISALMACEN2 AS V ON J.ID_PRODUCTO = V.ID_PRODUCTO WHERE J.ID_REPARACION = " & Clave & " AND V.SUCURSAL = 'BODEGA'"
    Set tRs = cnn.Execute(sBus)
    If tRs.EOF And tRs.BOF Then
        MsgBox "No hay productos pendientes"
    Else
        With tRs
            Do While Not (.EOF)
                Set tLi = ListView3.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    tLi.SubItems(1) = .Fields("CANTIDAD")
                    tLi.SubItems(2) = CDbl(.Fields("CANTIDAD")) * CDbl(Item.SubItems(1))
                .MoveNext
            Loop
        End With
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim tRs As ADODB.Recordset
    Dim sBus As String
    Dim tLi As ListItem
    sBus = "SELECT CANTIDAD FROM VSEXISALAMACEN2 WHERE ID_PRODUCTO = '%" & Item & "%' AND SUCURSAL = 'BODEGA'"
    Set tRs = cnn.Execute(sBus)
    Label3.Caption = Item
    Text2.Text = Item.SubItems(2)
    lblID3.Caption = Item.Index
End Sub
Private Sub ListView4_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Label4.Caption = Item
    lblID4.Caption = Item.Index
End Sub
Private Sub ListView5_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Label3.Caption = Item
    Text2.Text = "1"
    lblID3.Caption = "00+" & Item.SubItems(2)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim tRs As ADODB.Recordset
    Dim sBus As String
    Dim tLi As ListItem
    If KeyAscii = 13 And Text1.Text <> "" Then
        sBus = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM VSEXISALAMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' AND SUCURSAL = 'BODEGA'"
        Set tRs = cnn.Execute(sBus)
        With tRs
            If Not (.EOF And .BOF) Then
                Do While Not .EOF
                    Set tLi = ListView5.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                        tLi.SubItems(1) = .Fields("Descripcion")
                        tLi.SubItems(2) = .Fields("CANTIDAD")
                    .MoveNext
                Loop
            End If
        End With
    End If
End Sub
