VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmSanciones 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aplicación de sanción a notas de venta"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   8
      Top             =   3120
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmSanciones.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmSanciones.frx":030A
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
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   6
      Top             =   4320
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmSanciones.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "FrmSanciones.frx":1FD6
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmSanciones.frx":40B8
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
      Tab(0).Control(5)=   "ListView1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAgregar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   3840
         TabIndex        =   17
         Top             =   4680
         Width           =   1815
         Begin VB.OptionButton Option4 
            Caption         =   "Descuento"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Sanción"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   120
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Top             =   4920
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   4320
         Width           =   6975
      End
      Begin VB.CommandButton cmdAgregar 
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
         Left            =   6960
         Picture         =   "FrmSanciones.frx":40D4
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   5760
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "No. Nota"
         Height          =   255
         Left            =   5760
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   4455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2775
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4895
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   3840
         Width           =   6975
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Motivo :"
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
         Left            =   240
         TabIndex        =   13
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Monto de la sanción :"
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
         Left            =   240
         TabIndex        =   11
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar :"
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
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmSanciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub cmdAgregar_Click()
    Dim sBuscar As String
    Dim tLi As ListItem
    Dim tRs As ADODB.Recordset
    ListView1.ListItems.Clear
    If Option1.Value Then
        If IsNumeric(Text1.Text) Then
            sBuscar = "SELECT * FROM VENTAS WHERE ID_VENTA = " & Text1.Text & " AND FACTURADO = 0"
        Else
            MsgBox "FAVOR DE PONER SOLO VALORES NUMERICOS PARA BUSCAR POR NOTA!", vbExclamation, "SACC"
            Text1.Text = ""
            Exit Sub
        End If
    Else
        sBuscar = "SELECT * FROM VENTAS WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND FACTURADO = 0"
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("SUBTOTAL")) Then tLi.SubItems(2) = tRs.Fields("SUBTOTAL")
            If Not IsNull(tRs.Fields("IVA")) Then tLi.SubItems(3) = tRs.Fields("IVA")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(4) = tRs.Fields("TOTAL")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(5) = tRs.Fields("FECHA")
            tRs.MoveNext
        Loop
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
        .ColumnHeaders.Add , , "Nota", 1000
        .ColumnHeaders.Add , , "Cliente", 5000
        .ColumnHeaders.Add , , "Subtotal", 1500
        .ColumnHeaders.Add , , "IVA", 1000
        .ColumnHeaders.Add , , "Total", 1500
        .ColumnHeaders.Add , , "Fecha", 1500
    End With
End Sub
Private Sub Image8_Click()
On Error GoTo ManejaError
    If Label4.Caption <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        sBuscar = "SELECT SUBTOTAL, IVA, TOTAL, UNA_EXIBICION FROM VENTAS WHERE ID_VENTA = " & Label4.Caption
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF Or tRs.BOF) Then
            If tRs.Fields("UNA_EXIBICION") = "S" Then
                If Option3.Value Then
                    sBuscar = "INSERT INTO VENTAS_DETALLE (ID_VENTA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_VENTA, PRECIO_COSTO, GANANCIA, IMPORTE) VALUES (" & Label4.Caption & ", 'SANCIÓN', '" & Text2.Text & "', 1, -" & Text3.Text & ", -" & Text3.Text & ", 0, -" & Text3.Text & ");"
                Else
                    sBuscar = "INSERT INTO VENTAS_DETALLE (ID_VENTA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_VENTA, PRECIO_COSTO, GANANCIA, IMPORTE) VALUES (" & Label4.Caption & ", 'DESCUENTO', '" & Text2.Text & "', 1, -" & Text3.Text & ", -" & Text3.Text & ", 0, -" & Text3.Text & ");"
                End If
                cnn.Execute (sBuscar)
                sBuscar = "UPDATE VENTAS SET SUBTOTAL = SUBTOTAL - " & Text3.Text & ", IVA = (SUBTOTAL  -  " & Text3.Text & ") * " & CDbl(CDbl(VarMen.Text4(7).Text) / 100) & ", TOTAL  = (SUBTOTAL  -  " & Text3.Text & ") * (" & CDbl(CDbl(VarMen.Text4(7).Text) / 100) + 1 & ") WHERE ID_VENTA = " & Label4.Caption
                cnn.Execute (sBuscar)
            Else
                If Option3.Value Then
                    sBuscar = "INSERT INTO VENTAS_DETALLE (ID_VENTA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_VENTA, PRECIO_COSTO, GANANCIA, IMPORTE) VALUES (" & Label4.Caption & ", 'SANCIÓN', '" & Text2.Text & "', 1, -" & Text3.Text & ", -" & Text3.Text & ", 0, -" & Text3.Text & ");"
                Else
                    sBuscar = "INSERT INTO VENTAS_DETALLE (ID_VENTA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_VENTA, PRECIO_COSTO, GANANCIA, IMPORTE) VALUES (" & Label4.Caption & ", 'DESCUENTO', '" & Text2.Text & "', 1, -" & Text3.Text & ", -" & Text3.Text & ", 0, -" & Text3.Text & ");"
                End If
                cnn.Execute (sBuscar)
                sBuscar = "UPDATE VENTAS SET SUBTOTAL = SUBTOTAL - " & Text3.Text & ", IVA = (SUBTOTAL  -  " & Text3.Text & ") * " & CDbl(CDbl(VarMen.Text4(7).Text) / 100) & " , TOTAL  = (SUBTOTAL  -  " & Text3.Text & ") * " & (CDbl(CDbl(VarMen.Text4(7).Text) / 100) + 1) & " WHERE ID_VENTA = " & Label4.Caption
                cnn.Execute (sBuscar)
                sBuscar = "SELECT ID_CUENTA FROM CUENTA_VENTA WHERE ID_VENTA = " & Label4.Caption
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF Or tRs.BOF) Then
                    If Option3.Value Then
                        sBuscar = "INSERT INTO CUENTA_DETALLE (ID_CUENTA, ID_PRODUCTO, CANTIDAD, PRECIO_VENTA) VALUES (" & tRs.Fields("ID_CUENTA") & ", 'SANCIÓN', 1, -" & Text3.Text & ");"
                    Else
                        sBuscar = "INSERT INTO CUENTA_DETALLE (ID_CUENTA, ID_PRODUCTO, CANTIDAD, PRECIO_VENTA) VALUES (" & tRs.Fields("ID_CUENTA") & ", 'DESCUENTO', 1, -" & Text3.Text & ");"
                    End If
                    cnn.Execute (sBuscar)
                    sBuscar = "UPDATE CUENTAS SET TOTAL_COMPRA = TOTAL_COMPRA - " & Text3.Text * (1 + (VarMen.Text4(7).Text / 100)) & ", DEUDA = DEUDA  - " & Text3.Text * (1 + (VarMen.Text4(7).Text / 100)) & ", DEUDA_ACTUAL = DEUDA - " & Text3.Text * (1 + (VarMen.Text4(7).Text / 100)) & " WHERE ID_CUENTA  = " & tRs.Fields("ID_CUENTA")
                    cnn.Execute (sBuscar)
                Else
                    sBuscar = "SELECT ID_CUENTA FROM CUENTAS WHERE ID_VENTA = " & Label4.Caption
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF Or tRs.BOF) Then
                        If Option3.Value Then
                            sBuscar = "INSERT INTO CUENTA_DETALLE (ID_CUENTA, ID_PRODUCTO, CANTIDAD, PRECIO_VENTA) VALUES (" & tRs.Fields("ID_CUENTA") & ", 'SANCIÓN', 1, -" & Text3.Text & ");"
                        Else
                            sBuscar = "INSERT INTO CUENTA_DETALLE (ID_CUENTA, ID_PRODUCTO, CANTIDAD, PRECIO_VENTA) VALUES (" & tRs.Fields("ID_CUENTA") & ", 'DESCUENTO', 1, -" & Text3.Text & ");"
                        End If
                        cnn.Execute (sBuscar)
                        sBuscar = "UPDATE CUENTAS SET TOTAL_COMPRA = TOTAL_COMPRA - " & Text3.Text * (1 + (VarMen.Text4(7).Text / 100)) & ", DEUDA = DEUDA  - " & Text3.Text * (1 + (VarMen.Text4(7).Text / 100)) & ", DEUDA_ACTUAL = DEUDA - " & Text3.Text * (1 + (VarMen.Text4(7).Text / 100)) & " WHERE ID_CUENTA  = " & tRs.Fields("ID_CUENTA")
                        cnn.Execute (sBuscar)
                    Else
                        MsgBox "La cuenta de deuda no fue encontrda, favor de reportar el error al Depto. de sistemas", vbCritical, "SACC"
                    End If
                End If
                MsgBox "La sanción ha sido aplicada crrectamente!", vbExclamation, "SACC"
            End If
            Text2.Text = ""
            Text3.Text = ""
        End If
    End If
Exit Sub
ManejaError:
    MsgBox Err.Number & ": " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Label4.Caption = Item
    Label5.Caption = Item.SubItems(1)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT Descripcion FROM VENTAS_DETALLE WHERE ID_PRODUCTO = 'SANCIÓN' AND ID_VENTA = " & Item
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            MsgBox "La venta tiene una sanción con la Descripcion: " & tRs.Fields("Descripcion")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    If Option1.Value Then
        Valido = "1234567890"
    Else
        Valido = "1234567890.abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+"
    End If
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
