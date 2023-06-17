VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmPagoOrdenRapida 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagar Orden de Compra Rapida"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   12
      Top             =   3120
      Width           =   975
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Modificar"
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
      Begin VB.Image Image7 
         Height          =   810
         Left            =   120
         MouseIcon       =   "FrmPagoOrdenRapida.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmPagoOrdenRapida.frx":030A
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   4
      Top             =   4320
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmPagoOrdenRapida.frx":2434
         MousePointer    =   99  'Custom
         Picture         =   "FrmPagoOrdenRapida.frx":273E
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   2
      Top             =   5520
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
         MouseIcon       =   "FrmPagoOrdenRapida.frx":4100
         MousePointer    =   99  'Custom
         Picture         =   "FrmPagoOrdenRapida.frx":440A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Ordenes Pendientes de Cerrar"
      TabPicture(0)   =   "FrmPagoOrdenRapida.frx":64EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListView1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Combo2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Combo1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtTrans"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCheque"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.TextBox txtCheque 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   5640
         Width           =   2055
      End
      Begin VB.TextBox txtTrans 
         Height          =   285
         Left            =   5640
         TabIndex        =   16
         Top             =   5640
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   15
         Top             =   5160
         Width           =   2895
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   5040
         TabIndex        =   14
         Top             =   5160
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8040
         TabIndex        =   10
         Top             =   6120
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   6120
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   6120
         Width           =   2055
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4695
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8281
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label7 
         Caption         =   "No. DE CHEQUE"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   5640
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "No. DE TRANSFERENCIA"
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   5640
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "TIPO DE PAGO"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "BANCO"
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label lblID 
         Height          =   255
         Left            =   6600
         TabIndex        =   11
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "TOTAL A PAGAR"
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   6120
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL DE ORDEN:"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   6120
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmPagoOrdenRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim NumOrden As String
Dim NomProv As String
Dim sPendiente As String
Dim VarTot As Double
Private Sub Command1_Click()
    Text2.Enabled = True
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
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
        .ColumnHeaders.Add , , "Numero", 1000
        .ColumnHeaders.Add , , "Proveedor", 5500
        .ColumnHeaders.Add , , "Total", 1500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Id_proveedor", 0
    End With
    sBuscar = "SELECT * FROM BANCOS"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Combo2.Clear
            Do While Not .EOF
                Combo2.AddItem (.Fields("NOMBRE"))
                .MoveNext
            Loop
        Else
            MsgBox "NO EXISTEN BANCOS REGISTRADOS, NO PUEDE REGISTRAR PAGOS", vbInformation, "SACC"
        End If
        .Close
    End With
    sBuscar = "SELECT * FROM TPAGOS_OC"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Combo1.Clear
            Do While Not .EOF
                Combo1.AddItem (.Fields("Descripcion"))
                .MoveNext
            Loop
        Else
            MsgBox "FALLO DE INFORMACION, FAVOR DE LLAMAR A SOPORTE", vbInformation, "SACC"
        End If
        .Close
    End With
    Actualiza
End Sub
Private Sub Image7_Click()
    Dim sBuscar As String
    Dim Cont As Integer
    For Cont = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(Cont).Checked Then
            sBuscar = "UPDATE ORDEN_RAPIDA SET ESTADO = 'M' WHERE ID_ORDEN_RAPIDA = " & ListView1.ListItems(Cont)
            cnn.Execute (sBuscar)
        End If
    Next Cont
    Actualiza
End Sub
Private Sub Image8_Click()
    If Not ArchivoEnUso(App.Path & "\Cheque.pdf") Then
        If Combo1.Text <> "" And Combo2.Text <> "" And (txtCheque.Text <> "" Or txtTrans.Text <> "") And Text2.Text <> "" Then
            Dim sBuscar As String
            Dim tRs As ADODB.Recordset
            Dim Con As Integer
            Dim tot As Double
            Dim subto As Double
            tot = CDbl(Text1.Text) - CDbl(Text2.Text)
            subto = CDbl(sPendiente) - CDbl(Text2.Text)
            If MsgBox("ESTA POR CERRAR LA(S) ORDEN(ES) DE COMPRA SELECCIONADA(S), REGISTRARA UN PAGO ¿ESTA SEGURO QUE DESEA CONTINUAR?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
                For Con = 1 To ListView1.ListItems.Count
                   If ListView1.ListItems(Con).Checked = True Then
                        NumOrden = ListView1.ListItems(Con) & ", "
                        If Text2.Text <> Text1.Text Then
                            sBuscar = "INSERT INTO ABONOS_PAGO_OC (ID_PROVEEDOR, ID_ORDEN, NUM_ORDEN, PROVEEDOR, CANT_ABONO, FECHA, TIPO, TIPOPAGO, BANCO, NUMCHEQUE, NUMTRANS) VALUES ('" & IdProve & "', " & ListView1.ListItems(Con) & ", " & ListView1.ListItems(Con) & ", '" & NomProv & "', " & Text2.Text & ",'" & Format(Date, "dd/mm/yyyy") & "','R', '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtCheque.Text & "', '" & txtTrans.Text & "');"
                            cnn.Execute (sBuscar)
                        End If
                        If tot = 0 Or subto = 0 Then
                            sBuscar = "UPDATE ORDEN_RAPIDA SET ESTADO = 'F' WHERE ID_ORDEN_RAPIDA = " & ListView1.ListItems(Con)
                            Set tRs = cnn.Execute(sBuscar)
                            sBuscar = "INSERT INTO ABONOS_PAGO_OC (ID_PROVEEDOR, ID_ORDEN, NUM_ORDEN, PROVEEDOR, CANT_ABONO, FECHA, TIPO, TIPOPAGO, BANCO, NUMCHEQUE, NUMTRANS) VALUES ('" & IdProve & "', " & ListView1.ListItems(Con) & ", " & ListView1.ListItems(Con) & ", '" & NomProv & "', " & Text2.Text & ",'" & Format(Date, "dd/mm/yyyy") & "','R', '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtCheque.Text & "', '" & txtTrans.Text & "');"
                            cnn.Execute (sBuscar)
                        End If
                  End If
                Next
                FrmCheque.TxtIdProv.Text = IdProve
                FrmCheque.TxtNUM_ORDEN.Text = lblID.Caption
                FrmCheque.TxtTIPO_ORDEN.Text = "RAPIDA" 'tipo de orden de compra
                FrmCheque.txtNum2Let(0).Text = Text2.Text 'total de la orden de compra
                FrmCheque.TxtNOMBRE.Text = NomProv 'nombre del proveedor a recibir el pago
                FrmCheque.TxtNUM_CHEQUE.Text = txtCheque.Text 'numero de cheque
                FrmCheque.Combo1.Text = Combo2.Text 'banco
                If txtCheque.Text = "" Then
                    FrmCheque.TxtNUM_CHEQUE.Text = txtTrans.Text 'numero de cheque
                Else
                    FrmCheque.TxtNUM_CHEQUE.Text = txtCheque.Text 'numero de cheque
                End If
                FrmCheque.Combo1.Text = Combo2.Text 'banco
                FrmCheque.Show vbModal
                lblID.Caption = ""
                Actualiza
            End If
            For Con = 1 To ListView1.ListItems.Count
                ListView1.ListItems(Con).Checked = False
            Next
        Else
            MsgBox "Debe dar la información del pago antes de continuar", vbExclamation, "SACC"
        End If
    Else
        MsgBox "Cierre la poliza de cheque que tiene abierta para poder continuar", vbExclamation, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim txtotal As Double
    Dim IdProve As String
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim Con As Integer
    Dim ConSel As Integer
    Text1.Text = "0.00"
    Text2.Text = "0.00"
    lblID.Caption = ""
    ConSel = 0
    NomProv = Item.SubItems(1)
    VarTot = Item.SubItems(2)
    txtotal = 0
    IdProve = ""
    For Con = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(Con).Checked Then
            If IdProve = "" Then IdProve = ListView1.ListItems.Item(Con).SubItems(4)
            If IdProve = ListView1.ListItems.Item(Con).SubItems(4) Then
                lblID.Caption = lblID.Caption & ListView1.ListItems(Con) & ", "
                Text1.Text = CDbl(Text1.Text) + CDbl(ListView1.ListItems(Con).SubItems(2))
                sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN='" & ListView1.ListItems.Item(Con) & "' AND TIPO='R'"
                Set tRs = cnn.Execute(sBusca)
                If Not (tRs.EOF And tRs.BOF) Then
                    If Not IsNull(tRs.Fields("CANT_ABONO")) Then
                        Text2.Text = CDbl(Text1.Text) - CDbl(tRs.Fields("CANT_ABONO"))
                    Else
                        Text2.Text = Text1.Text
                    End If
                Else
                    Text2.Text = Text1.Text
                End If
                sPendiente = Text2.Text
            Else
                ListView1.ListItems(Con).Checked = False
                MsgBox "TODAS LAS ORDENES DEBEN SER DEL MISMO PROVEEDOR!", vbExclamation, "SACC"
            End If
            ConSel = ConSel + 1
            If CDbl(ConSel) > 1 Then
                Command1.Enabled = False
                Text2.Enabled = False
            Else
                Command1.Enabled = True
            End If
        End If
    Next
End Sub
Private Sub Actualiza()
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBuscar As String
    Dim CompDolar As Double
    Dim totalreal As Double
    sBuscar = "SELECT COMPRA FROM DOLAR WHERE FECHA = '" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        If Not IsNull(tRs.Fields("COMPRA")) Then
            CompDolar = tRs.Fields("COMPRA")
        Else
            CompDolar = InputBox("DE EL PRECIO DE VENTA DEL DOLAR HOY!")
            sBuscar = "INSERT INTO DOLAR (FECHA, COMPRA, VENTA) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', " & CompDolar & ", " & InputBox("CON FIN DE ACTUALIZAR EL TIPO DE CAMBIO A LA FECHA, DE EL PRECIO DE COMPRA DEL DOLAR HOY!") & ");"
            cnn.Execute (sBuscar)
        End If
    Else
        CompDolar = InputBox("DE EL PRECIO DE VENTA DEL DOLAR HOY!")
        sBuscar = "INSERT INTO DOLAR (FECHA, COMPRA, VENTA) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', " & CompDolar & ", " & InputBox("CON FIN DE ACTUALIZAR EL TIPO DE CAMBIO A LA FECHA, DE EL PRECIO DE COMPRA DEL DOLAR HOY!") & ");"
        cnn.Execute (sBuscar)
    End If
    ListView1.ListItems.Clear
    sBuscar = "SELECT ID_ORDEN_RAPIDA, NOMBRE, FECHA, ID_PROVEEDOR, MONEDA, SUM(RETENCION) AS RETENCION, SUM(TOTAL) AS TOTAL, SUM(IVADIEZ) AS IVADIEZ FROM VsOrdenRapida WHERE ESTADO = 'A' GROUP BY ID_ORDEN_RAPIDA, NOMBRE, FECHA, ID_PROVEEDOR, MONEDA ORDER BY ID_ORDEN_RAPIDA "
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set ItMx = ListView1.ListItems.Add(, , tRs.Fields("ID_ORDEN_RAPIDA"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then ItMx.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FECHA")) Then ItMx.SubItems(3) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(4) = tRs.Fields("ID_PROVEEDOR")
            If tRs.Fields("MONEDA") = "DOLARES" Then
                totalreal = Format(CDbl(tRs.Fields("TOTAL")) - CDbl(tRs.Fields("RETENCION") - CDbl(tRs.Fields("IVADIEZ"))), "###,###,##0.00")
                If Not IsNull(totalreal) Then ItMx.SubItems(2) = Trim(Format(CDbl(totalreal) * CDbl(CompDolar), "###,###,##0.00"))
                totalreal = 0
            End If
            If tRs.Fields("MONEDA") = "PESOS" Then
                sBuscar = "SELECT SUM (TOTAL) AS TOTAL FROM VsOrdenRapida WHERE ID_ORDEN_RAPIDA = " & tRs.Fields("ID_ORDEN_RAPIDA")
                Set tRs1 = cnn.Execute(sBuscar)
                totalreal = Format(CDbl(tRs1.Fields("TOTAL")), "###,###,##0.00")
                ItMx.SubItems(2) = Trim(totalreal)
                totalreal = 0
            End If
            tRs.MoveNext
        Loop
    End If
End Sub
Public Function ArchivoEnUso(ByVal sFileName As String) As Boolean
    Dim filenum As Integer, errnum As Integer
    On Error Resume Next
    filenum = FreeFile()
    Open sFileName For Input Lock Read As #filenum
    Close filenum
    errnum = Err
    On Error GoTo 0
    Select Case errnum
        Case 0
            ArchivoEnUso = False
        Case 70
            ArchivoEnUso = True
        Case Else
            Error errnum
    End Select
End Function

