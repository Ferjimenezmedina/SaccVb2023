VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPagoOrden 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ORDENES DE COMPRA PENDIENTES DE PAGO"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Reporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8160
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   29
      Top             =   5280
      Visible         =   0   'False
      Width           =   135
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   5520
      TabIndex        =   10
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmPagoOrden.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFolio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblID"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbldeuda"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "opnIndirecta"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "opnInternacional"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "opnNacional"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtTotal"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "textsalpago"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.TextBox textsalpago 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   32
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   195
         Left            =   3000
         TabIndex        =   31
         Top             =   2160
         Width           =   255
      End
      Begin VB.Frame Frame3 
         Height          =   3135
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   3135
         Begin VB.CommandButton Command2 
            Caption         =   "Fact.Prove"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   30
            Top             =   2760
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.TextBox txtCheque 
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Top             =   1680
            Width           =   2895
         End
         Begin VB.TextBox txtTrans 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   2280
            Width           =   2895
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   2895
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label7 
            Caption         =   "NUMERO DE CHEQUE"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label Label8 
            Caption         =   "NUMERO DE TRANSFERENCIA"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2040
            Width           =   2415
         End
         Begin VB.Label Label3 
            Caption         =   "TIPO DE PAGO"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "BANCO"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Text            =   "0"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton opnNacional 
         Caption         =   "Nacional"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton opnInternacional 
         Caption         =   "Internacional"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton opnIndirecta 
         Caption         =   "Indirecta"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbldeuda 
         Caption         =   "PAGO A REALIZAR"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblID 
         Height          =   255
         Left            =   2040
         TabIndex        =   28
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL"
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "FOLIO"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblFolio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   840
         TabIndex        =   25
         Top             =   1200
         Width           =   2415
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   3135
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9120
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   8
      Top             =   4560
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
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmPagoOrden.frx":001C
         MousePointer    =   99  'Custom
         Picture         =   "frmPagoOrden.frx":0326
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   6
      Top             =   3360
      Width           =   975
      Begin VB.Label Label1 
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
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmPagoOrden.frx":2408
         MousePointer    =   99  'Custom
         Picture         =   "frmPagoOrden.frx":2712
         Top             =   240
         Width           =   675
      End
   End
   Begin MSComctlLib.ListView lvwOCInternacionales 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwOCNacionales 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwOCIndirectas 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indirectas :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nacionales :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Nacionales 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Internacionales :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmPagoOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As adodb.Connection
Dim TipoOrden As String
Dim sPendiente As String
Dim tip As String
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Command1_Click()
    FRORPEN.Show vbModal
End Sub
Private Sub Command2_Click()
    frmfactpro.Show vbModal
End Sub
Private Sub Command3_Click()
    textsalpago.Enabled = True
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New adodb.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With lvwOCInternacionales
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_ORDEN", 0
        .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
        .ColumnHeaders.Add , , "FOLIO", 500
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
        .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
        .ColumnHeaders.Add , , "COMENTARIO", 500
        .ColumnHeaders.Add , , "DEUDA PENDIENTE", 1440
    End With
    With lvwOCNacionales
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_ORDEN", 0
        .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
        .ColumnHeaders.Add , , "FOLIO", 500
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
        .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
        .ColumnHeaders.Add , , "COMENTARIO", 500
         .ColumnHeaders.Add , , "DEUDA PENDIENTE", 1440
    End With
    With lvwOCIndirectas
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_ORDEN", 0
        .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
        .ColumnHeaders.Add , , "FOLIO", 500
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
        .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
        .ColumnHeaders.Add , , "COMENTARIO", 1440
         .ColumnHeaders.Add , , "DEUDA PENDIENTE", 1440
    End With
    If Hay_Ordenes_Compra Then
        Llenar_Lista_Compras "Internacionales"
        Llenar_Lista_Compras "Nacionales"
        Llenar_Lista_Compras "Indirectas"
    End If
    Dim sBuscar As String
    Dim tRs As Recordset
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
                Combo1.AddItem (.Fields("DESCRIPCION"))
                .MoveNext
            Loop
        Else
            MsgBox "FALLO DE INFORMACION, FAVOR DE LLAMAR A SOPORTE", vbInformation, "SACC"
        End If
        .Close
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Public Function Hay_Ordenes_Compra() As Boolean
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    sBuscar = "SELECT  count(*) as Orden_Compra From ORDEN_COMPRA WHERE Confirmada='X'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If .Fields("ORDEN_COMPRA") <> 0 Then
            Hay_Ordenes_Compra = True
        Else
            Hay_Ordenes_Compra = False
        End If
        .Close
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Lista_Compras(Tipo As String)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim CompDolar As Double
    Dim NumOrden As Integer
    Dim tRs2 As Recordset
    Dim sBusca As String
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
    sBuscar = "SELECT OC.Id_Orden_Compra, OC.NUM_ORDEN, OC.Id_Proveedor, P.Nombre, ((OC.Total - OC.Discount) + OC.TAX + OC.Freight + OC.Otros_Cargos) AS Total_Pagar, OC.COMENTARIO, MONEDA FROM ORDEN_COMPRA AS OC JOIN PROVEEDOR AS P ON P.Id_Proveedor = OC.Id_Proveedor WHERE OC.Confirmada = 'X' AND OC.Tipo = '"
    Select Case Tipo
        Case "Internacionales":
            Me.lvwOCInternacionales.ListItems.Clear
            sBuscar = sBuscar & "I'"
            Set tRs = cnn.Execute(sBuscar)
            With tRs
                While Not .EOF
                    Set ItMx = Me.lvwOCInternacionales.ListItems.Add(, , .Fields("ID_ORDEN_COMPRA"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(1) = Trim(.Fields("ID_PROVEEDOR"))
                    If Not IsNull(.Fields("NUM_ORDEN")) Then ItMx.SubItems(2) = Trim(.Fields("NUM_ORDEN"))
                  ''''''modificacion line de abajo
                    NumOrden = .Fields("NUM_ORDEN")
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(3) = Trim(.Fields("Nombre"))
                    If .Fields("MONEDA") = "DOLARES" Then
                        If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(Format(CDbl(.Fields("Total_Pagar")) * CDbl(CompDolar), "0.00"))
                    Else
                        If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(Format(CDbl(.Fields("Total_Pagar")), "0.00"))
                    End If
                    If Not IsNull(.Fields("COMENTARIO")) Then ItMx.SubItems(5) = Trim(.Fields("COMENTARIO"))
                    .MoveNext
                Wend
            .Close
            End With
        Case "Nacionales":
            Me.lvwOCNacionales.ListItems.Clear
            sBuscar = sBuscar & "N'"
            Set tRs = cnn.Execute(sBuscar)
            With tRs
                While Not .EOF
                    Set ItMx = Me.lvwOCNacionales.ListItems.Add(, , .Fields("ID_ORDEN_COMPRA"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(1) = Trim(.Fields("ID_PROVEEDOR"))
                    If Not IsNull(.Fields("NUM_ORDEN")) Then ItMx.SubItems(2) = Trim(.Fields("NUM_ORDEN"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(3) = Trim(.Fields("Nombre"))
                    If .Fields("MONEDA") = "DOLARES" Then
                        If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(Format(CDbl(.Fields("Total_Pagar")) * CDbl(CompDolar), "0.00"))
                    Else
                        If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(.Fields("Total_Pagar"))
                    End If
                    If Not IsNull(.Fields("COMENTARIO")) Then ItMx.SubItems(5) = Trim(.Fields("COMENTARIO"))
                    .MoveNext
                Wend
            .Close
            End With
        Case "Indirectas":
            Me.lvwOCIndirectas.ListItems.Clear
            sBuscar = sBuscar & "X'"
            Set tRs = cnn.Execute(sBuscar)
            With tRs
                While Not .EOF
                    Set ItMx = Me.lvwOCIndirectas.ListItems.Add(, , .Fields("ID_ORDEN_COMPRA"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(1) = Trim(.Fields("ID_PROVEEDOR"))
                    If Not IsNull(.Fields("NUM_ORDEN")) Then ItMx.SubItems(2) = Trim(.Fields("NUM_ORDEN"))
                    If Not IsNull(.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(3) = Trim(.Fields("Nombre"))
                    If .Fields("MONEDA") = "DOLARES" Then
                        If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(Format(CDbl(.Fields("Total_Pagar")) * CDbl(CompDolar), "0.00"))
                    Else
                        If Not IsNull(.Fields("Total_Pagar")) Then ItMx.SubItems(4) = Trim(.Fields("Total_Pagar"))
                    End If
                    If Not IsNull(.Fields("COMENTARIO")) Then ItMx.SubItems(5) = Trim(.Fields("COMENTARIO"))
                    .MoveNext
                Wend
            .Close
            End With
        Case Else:
            MsgBox "ERROR GRAVE. LA APLICACIÓN TERMINARA", vbCritical, "SACC"
            End
    End Select
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image8_Click()
    Dim sBuscar As String
    Dim tot As Double
    tot = CDbl(txtTotal) - CDbl(textsalpago.Text)
    Dim Con As Integer
    If Combo2.Text <> "" And txtCheque.Text = "" Then
        MsgBox "Se debe dar un numero de cheque", vbExclamation, "SACC"
        Exit Sub
    End If
    If MsgBox("ESTA POR REGISTRARA UN PAGO A LA ORDEN DE COMPRA SELECCIONADA,¿ESTA SEGURO QUE DESEA CONTINUAR?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
        If lblID.Caption <> "" Then
            If Combo1.Text <> "" And Combo2.Text <> "" And txtCheque.Text <> "" Then
                FrmCheque.TxtNUM_ORDEN.Text = lblFolio.Caption 'numero de orden de compra
                FrmCheque.TxtTIPO_ORDEN.Text = TipoOrden 'tipo de orden de compra
                FrmCheque.txtNum2Let(0).Text = textsalpago
                FrmCheque.TxtNOMBRE.Text = Label5.Caption 'nombre del proveedor a recibir el pago
                FrmCheque.TxtNUM_CHEQUE.Text = txtCheque.Text 'numero de cheque
                FrmCheque.Combo1.Text = Combo2.Text 'banco
                FrmCheque.Show vbModal
                If textsalpago <> txtTotal.Text Then
                    If TipoOrden = "INTERNACIONAL" Then
                        If lvwOCInternacionales.ListItems(Con).Selected Then
                            sBuscar = "INSERT INTO ABONOS_PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANT_ABONO, FECHA,TIPO,NUM_ORDEN) VALUES (" & lblID.Caption & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "','I','" & lblFolio.Caption & "');"
                            cnn.Execute (sBuscar)
                        End If
                    End If
                    If TipoOrden = "NACIONAL" Then
                        If lvwOCNacionales.ListItems(Con).Selected Then
                            sBuscar = "INSERT INTO ABONOS_PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANT_ABONO, FECHA,TIPO,NUM_ORDEN) VALUES (" & lblID.Caption & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "','N','" & lblFolio.Caption & "');"
                            cnn.Execute (sBuscar)
                        End If
                    End If
                    If TipoOrden = "INDIRECTA" Then
                        If lvwOCIndirectas.ListItems(Con).Selected Then
                            sBuscar = "INSERT INTO ABONOS_PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANT_ABONO, FECHA,TIPO,NUM_ORDEN) VALUES (" & lblID.Caption & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "','M','" & lblFolio.Caption & "');"
                            cnn.Execute (sBuscar)
                        End If
                    End If
               End If
                ' si el pago escrito es igual al pendiente entonces lo marca como pagado
                If textsalpago.Text = sPendiente Or tot < 0 Then
                    sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'Y' WHERE ID_ORDEN_COMPRA =' " & lblFolio.Caption & "' AND TIPO='" & tip & "' "
                    cnn.Execute (sBuscar)
                End If
                ' insert que generaba el pago por el total de la factura (sin parecialidades)
                ' insert nuevo que permite el pago en parcialidades dado en "textsalpago"
                sBuscar = "INSERT INTO PAGO_OC (ID_ORDEN, TIPOPAGO, BANCO, NUMTRANS, NUMCHEQUE, CANTIDAD, FECHA) VALUES (" & lblID.Caption & ", '" & Combo1.Text & "', '" & Combo2.Text & "', '" & txtTrans.Text & "', '" & txtCheque.Text & "', " & textsalpago.Text & ", '" & Format(Date, "dd/mm/yyyy") & "');"
                cnn.Execute (sBuscar)
                TipoOrden = ""
            Else
                MsgBox "NO PUEDE REGISTRAR PAGOS SIN LA INFORMACION COMPLETA", vbInformation, "SACC"
            End If
            MsgBox "DEBE SELECCIONAR UNA ORDEN DE COMPRA A PAGAR", vbInformation, "SACC"
        End If
    Else
        MsgBox "NO PUEDE REGISTRAR UN PAGO MAYOR QUE EL PENDIENTE  " & sPendiente, vbExclamation, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwOCIndirectas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwOCIndirectas.ListItems.Count > 0 Then
        opnInternacional.Value = True
        txtTotal.Text = Item.SubItems(4)
        Label5.Caption = Item.SubItems(3)
        lblFolio.Caption = Item.SubItems(2)
        lblID.Caption = Item
        TipoOrden = "INDIRECTA"
        Dim sBuscar As String
        Dim tRs As Recordset
        sBuscar = "SELECT SUM (CANTIDAD) AS TOTAL FROM PAGO_OC WHERE ID_ORDEN = " & lblID
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            If Not IsNull(tRs.Fields("TOTAL")) Then
                textsalpago.Text = CDbl(txtTotal.Text) - CDbl(tRs.Fields("TOTAL"))
            Else
                textsalpago.Text = txtTotal.Text
            End If
        End If
        sPendiente = textsalpago.Text
    End If
End Sub
Private Sub lvwOCInternacionales_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwOCInternacionales.ListItems.Count > 0 Then
        opnIndirecta.Value = True
        txtTotal.Text = Item.SubItems(4)
        textsalpago.Text = Item.SubItems(6)
        Label5.Caption = Item.SubItems(3)
        lblFolio.Caption = Item.SubItems(2)
        lblID.Caption = Item
        TipoOrden = "INTERNACIONAL"
        tip = "I"
        Dim sBuscar As String
        Dim tRs As Recordset
        '''lo mod debido que cuando busca  el folio en la tabla  no hay que lo identifi si era INT,NAC
        '''CREO QUE LA TABLA DE CHEQUE TAMBIEN SE PUEDE TOMAR LA INFO  DE CUANDO SE LE CLICK
        'AL ITEM DELALISVIEW JALE  LOS ABONOS QUE TIENE  ESA  ORDEN
        sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN='" & lblFolio.Caption & "' AND TIPO='I'"
        Set tRs = cnn.Execute(sBusca)
        If Not (tRs.EOF And tRs.BOF) Then
            If Not IsNull(tRs.Fields("CANT_ABONO")) Then
                textsalpago.Text = CDbl(txtTotal.Text) - CDbl(tRs.Fields("CANT_ABONO"))
            Else
                textsalpago.Text = txtTotal.Text
            End If
        End If
        sPendiente = textsalpago.Text
    End If
End Sub
Private Sub lvwOCNacionales_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwOCNacionales.ListItems.Count > 0 Then
        opnNacional.Value = True
        txtTotal.Text = Item.SubItems(4)
        textsalpago.Text = Item.SubItems(6)
        Label5.Caption = Item.SubItems(3)
        lblFolio.Caption = Item.SubItems(2)
        lblID.Caption = Item
        TipoOrden = "NACIONAL"
        tip = "N"
        Dim sBuscar As String
        Dim tRs As Recordset
        sBusca = "SELECT SUM(CANT_ABONO) AS CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN='" & lblFolio.Caption & "' AND TIPO='N'"
        Set tRs = cnn.Execute(sBusca)
        If Not (tRs.EOF And tRs.BOF) Then
            If Not IsNull(tRs.Fields("CANT_ABONO")) Then
                textsalpago.Text = CDbl(txtTotal.Text) - CDbl(tRs.Fields("CANT_ABONO"))
            Else
                textsalpago.Text = txtTotal.Text
            End If
        End If
        sPendiente = textsalpago.Text
    End If
End Sub
