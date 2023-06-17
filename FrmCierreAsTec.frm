VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCierreAsTec 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cerrar Asistencia Técnicas"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8640
      TabIndex        =   12
      Top             =   3360
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
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmCierreAsTec.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCierreAsTec.frx":030A
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8640
      TabIndex        =   10
      Top             =   4560
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCierreAsTec.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "FrmCierreAsTec.frx":1FD6
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Cierre"
      TabPicture(0)   =   "FrmCierreAsTec.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListView2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "FrmCierreAsTec.frx":40D4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "ListView3"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "Command2"
      Tab(1).ControlCount=   4
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   6000
         TabIndex        =   6
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton Command2 
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
         Left            =   -67920
         Picture         =   "FrmCierreAsTec.frx":40F0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
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
         Left            =   7080
         Picture         =   "FrmCierreAsTec.frx":6AC2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Información de la Asistencia Técnica :"
         Height          =   2295
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   8175
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   24
            Top             =   1440
            Width           =   7695
         End
         Begin VB.Label Label7 
            Caption         =   "Caracteristicas del atriculo :"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label11 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   22
            Top             =   720
            Width           =   6975
         End
         Begin VB.Label Label10 
            Caption         =   "Cliente :"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   20
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   6960
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   " Asistencia Técnica numero :"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   480
            Width           =   2175
         End
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   8
         Top             =   3120
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3413
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
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
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
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   4560
         TabIndex        =   4
         Top             =   2880
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   4560
         TabIndex        =   3
         Top             =   2640
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   2760
         Width           =   3375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   120
         TabIndex        =   5
         Top             =   3120
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3413
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label13 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   5160
         TabIndex        =   25
         Top             =   5160
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Articulos agregados al cobro de la Asistencia Técnica :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   16
         Top             =   2880
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Asistencia Técnica :"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2880
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmCierreAsTec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Attribute cnn.VB_VarHelpID = -1
Dim ClvProdu As String
Dim DesPordu As String
Dim PreProdu As String
Dim ProdElim As Integer
Private Sub Command1_Click()
    If Text2.Text = "" Then
        Text2.Text = "1"
    End If
    If ClvProdu <> "" Then
        Dim tLi As ListItem
        Set tLi = ListView3.ListItems.Add(, , ClvProdu)
            tLi.SubItems(1) = DesPordu
            tLi.SubItems(2) = Text2.Text
            tLi.SubItems(3) = PreProdu
        ClvProdu = ""
        DesPordu = ""
        PreProdu = ""
    Else
        MsgBox "NO SE HA SELECCIONADO UN PRODUCTO!", vbExclamation, "SACC"
    End If
End Sub
Private Sub Command2_Click()
    If ProdElim > 0 Then
        ListView3.ListItems.Remove (ProdElim)
        ProdElim = 0
    Else
        MsgBox "NO SE HA SELECCIONADO UN PRODUCTO!", vbExclamation, "SACC"
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
        .ColumnHeaders.Add , , "Clave", 2500
        .ColumnHeaders.Add , , "Descripcion", 5700
        .ColumnHeaders.Add , , "Precio de venta", 1200
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. Asistencia", 500
        .ColumnHeaders.Add , , "Cliente", 5700
        .ColumnHeaders.Add , , "Modelo", 1200
        .ColumnHeaders.Add , , "Marca", 1200
        .ColumnHeaders.Add , , "Tipo", 1200
        .ColumnHeaders.Add , , "Descripcion de las Piezas", 5700
        .ColumnHeaders.Add , , "Garantia", 0
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave", 2500
        .ColumnHeaders.Add , , "Descripcion", 5700
        .ColumnHeaders.Add , , "Cantidad", 1200
        .ColumnHeaders.Add , , "Precio de venta", 1200
    End With
    Buscar
    'BuscaArticulo
End Sub
Private Sub Image8_Click()
    If Label6.Caption <> "" Then
        If ListView3.ListItems.COUNT <> 0 Then
            Dim sBuscar As String
            Dim Cont As Integer
            Dim tRs As ADODB.Recordset
            For Cont = 1 To ListView3.ListItems.COUNT
                sBuscar = "SELECT ID_PRODUCTO FROM COBRO_ASISTENCIA_TECNICA WHERE ID_PRODUCTO = '" & ListView3.ListItems.Item(Cont) & "' AND ID_AS_TEC = " & Label6.Caption
                Set tRs = cnn.Execute(sBuscar)
                If tRs.EOF And tRs.BOF Then
                    sBuscar = "INSERT INTO COBRO_ASISTENCIA_TECNICA (ID_AS_TEC, ID_PRODUCTO, CANTIDAD) VALUES (" & Label6.Caption & ", '" & ListView3.ListItems.Item(Cont) & "', " & ListView3.ListItems.Item(Cont).SubItems(2) & ");"
                    cnn.Execute (sBuscar)
                Else
                    If MsgBox("EL PRODUCTO " & ListView3.ListItems.Item(Cont) & " YA ESTA REGISTRADO EN ESTA ASISTENCIA (" & Label6.Caption & ") ¿DESEA REEMPLAZAR CON LA NUEVA CANTIDAD DADA?", vbYesNo + vbInformation + vbDefaultButton1, "SACC") = vbYes Then
                        sBuscar = "UPDATE COBRO_ASISTENCIA_TECNICA SET CANTIDAD =  " & ListView3.ListItems.Item(Cont).SubItems(2) & " WHERE ID_PRODUCTO = '" & ListView3.ListItems.Item(Cont) & "' AND ID_AS_TEC = " & Label6.Caption
                        cnn.Execute (sBuscar)
                    End If
                End If
            Next Cont
            sBuscar = "UPDATE ASISTENCIA_TECNICA SET ATENDIDO =  3 WHERE ID_AS_TEC = " & Label6.Caption
            cnn.Execute (sBuscar)
            Label5.Caption = ""
            Label6.Caption = ""
            Label11.Caption = ""
            Label12.Caption = ""
            ListView3.ListItems.Clear
        Else
            If MsgBox("ESTA CERRANDO LA ASISTENCIA TÉCNICA SIN AGREGAR ARTICULOS, ¿DESEA REGISTRARLA SIN COBRO?", vbYesNo + vbInformation + vbDefaultButton1, "SACC") = vbYes Then
                sBuscar = "UPDATE ASISTENCIA_TECNICA SET ATENDIDO = 3 WHERE ID_AS_TEC = " & Label6.Caption
                cnn.Execute (sBuscar)
            End If
        End If
    Else
        MsgBox "NO SE HA SELECCIONADO LA ASISTENCIA TÉCNICA!", vbExclamation, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Buscar()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT * FROM ASISTENCIA_TECNICA WHERE ATENDIDO = '0' ORDER BY ID_AS_TEC"
    Set tRs = cnn.Execute(sBuscar)
    ListView2.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_AS_TEC"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("MODELO")) Then tLi.SubItems(2) = tRs.Fields("MODELO")
                If Not IsNull(tRs.Fields("MARCA")) Then tLi.SubItems(3) = tRs.Fields("MARCA")
                If Not IsNull(tRs.Fields("TIPO_ARTICULO")) Then tLi.SubItems(4) = tRs.Fields("TIPO_ARTICULO")
                If Not IsNull(tRs.Fields("Descripcion_PIEZAS")) Then tLi.SubItems(5) = tRs.Fields("Descripcion_PIEZAS")
                If Not IsNull(tRs.Fields("GARANTIA")) Then
                    If tRs.Fields("GARANTIA") = "1" Then
                        tLi.SubItems(6) = "SI"
                    Else
                        tLi.SubItems(6) = "NO"
                    End If
                Else
                    tLi.SubItems(6) = "NO"
                End If
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub BuscaArticulo()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If Option1.value = True Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' ORDER BY ID_PRODUCTO"
    Else
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO, GANANCIA FROM ALMACEN3 WHERE Descripcion LIKE '%" & Text1.Text & "%' ORDER BY ID_PRODUCTO"
    End If
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
                If Not IsNull(tRs.Fields("PRECIO_COSTO")) And Not IsNull(tRs.Fields("GANANCIA")) Then tLi.SubItems(2) = (CDbl(tRs.Fields("GANANCIA")) + 1) * CDbl(tRs.Fields("PRECIO_COSTO"))
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ClvProdu = Item
    DesPordu = Item.SubItems(1)
    PreProdu = Item.SubItems(2)
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.SubItems(6) = "SI" Then
        Label5.Caption = "Garantia"
    Else
        Label5.Caption = ""
    End If
    Label6.Caption = Item
    Label11.Caption = Item.SubItems(1)
    Label12.Caption = "Modelo : " & Item.SubItems(2) & ", Marca : " & Item.SubItems(3) & " Tipo : " & Item.SubItems(4)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView3.ListItems.Clear
    sBuscar = "SELECT * FROM VsATCerrada WHERE ID_AS_TEC = " & Item
    Set tRs = cnn.Execute(sBuscar)
    If tRs.EOF And tRs.BOF Then
        Do While Not tRs.EOF
            Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            tLi.SubItems(1) = tRs.Fields("Descripcion")
            tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            tLi.SubItems(3) = tRs.Fields("PRECIO_VENTA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ProdElim = Item.Index
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BuscaArticulo
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
