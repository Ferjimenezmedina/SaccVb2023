VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmReImprime 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reimprimir"
   ClientHeight    =   6255
   ClientLeft      =   2340
   ClientTop       =   1350
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9000
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9000
      TabIndex        =   14
      Top             =   4920
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmReImprime.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmReImprime.frx":030A
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
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Filtros"
      TabPicture(0)   =   "FrmReImprime.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.TextBox Text3 
         Height          =   1935
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   3960
         Visible         =   0   'False
         Width           =   8415
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Por Producto"
         Height          =   255
         Left            =   7200
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Por Cliente"
         Height          =   255
         Left            =   7200
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Documento"
         Height          =   2295
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   3255
         Begin VB.OptionButton Option22 
            Caption         =   "Remisión"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1920
            Width           =   2295
         End
         Begin VB.OptionButton Option21 
            Caption         =   "P. Cheque"
            Height          =   255
            Left            =   2040
            TabIndex        =   33
            Top             =   1200
            Width           =   1095
         End
         Begin VB.OptionButton Option16 
            Caption         =   "Entrada Proveedores Varios"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1680
            Width           =   2295
         End
         Begin VB.OptionButton Option14 
            Caption         =   "Compra Proveedores Varios"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   2535
         End
         Begin VB.OptionButton Option15 
            Caption         =   "Comanda(Juego Rep.)"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   1935
         End
         Begin VB.OptionButton Option13 
            Caption         =   "Asistencia"
            Height          =   255
            Left            =   2040
            TabIndex        =   20
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton Option12 
            Caption         =   "Venta Programada"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1200
            Width           =   1695
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Vale de Caja"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Entrada de Almacen"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Cotización"
            Height          =   195
            Left            =   2040
            TabIndex        =   16
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Comanda"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Venta"
            Height          =   195
            Left            =   2040
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Garantia"
            Height          =   195
            Left            =   2040
            TabIndex        =   7
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4200
         TabIndex        =   4
         Top             =   480
         Width           =   2895
      End
      Begin VB.Frame Frame1 
         Caption         =   "Reimprimir"
         Height          =   1335
         Left            =   240
         TabIndex        =   1
         Top             =   2520
         Width           =   3255
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Remisión"
            Height          =   195
            Left            =   1560
            TabIndex        =   37
            Top             =   960
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   1560
            TabIndex        =   29
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
            Begin VB.OptionButton Option18 
               Caption         =   "Por Nota"
               Height          =   195
               Left            =   120
               TabIndex        =   32
               Top             =   120
               Width           =   1215
            End
            Begin VB.OptionButton Option19 
               Caption         =   "Por Factura"
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Top             =   360
               Width           =   1215
            End
            Begin VB.OptionButton Option20 
               Caption         =   "Por Comanda"
               Height          =   195
               Left            =   120
               TabIndex        =   30
               Top             =   600
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   1560
            TabIndex        =   24
            Top             =   120
            Width           =   1575
            Begin VB.OptionButton Option8 
               Caption         =   "Nacional"
               Height          =   195
               Left            =   120
               TabIndex        =   28
               Top             =   120
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.OptionButton Option9 
               Caption         =   "Internacional"
               Height          =   195
               Left            =   120
               TabIndex        =   27
               Top             =   360
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.OptionButton Option10 
               Caption         =   "Indirecta"
               Height          =   195
               Left            =   120
               TabIndex        =   26
               Top             =   600
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.OptionButton Option17 
               Caption         =   "Rapida"
               Height          =   195
               Left            =   120
               TabIndex        =   25
               Top             =   840
               Visible         =   0   'False
               Width           =   1215
            End
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdBorrar 
            Caption         =   "Reimpimir"
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
            Left            =   240
            Picture         =   "FrmReImprime.frx":2408
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   720
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   3600
         TabIndex        =   12
         Top             =   960
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4683
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   1935
         Left            =   240
         TabIndex        =   13
         Top             =   3960
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3413
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
         Caption         =   "Comentarios :"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   3960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   9000
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FrmReImprime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim NOM As String
Dim fech As Date
Private Sub cmdBorrar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If Text1.Text <> "" Then
        If Option1.Value = True Then
            If Option19.Value Then
                sBuscar = "SELECT NO_COM_AT FROM VENTAS WHERE FOLIO = '" & Text1.Text & "'"
                Set tRs = cnn.Execute(sBuscar)
                If (tRs.EOF And tRs.BOF) Then
                    MsgBox "EL FOLIO DE LA FACTURA NO FUE ENCONTRADO, ES POSIBLE QUE ESTE CANCELADO", vbExclamation, "SACC"
                Else
                    Do While Not tRs.EOF
                        Text1.Text = Replace(tRs.Fields("NO_COM_AT"), "C", "")
                        ImpComanda
                    Loop
                    tRs.MoveNext
                    Text1.Text = ""
                End If
            Else
                If Option20.Value Then
                    sBuscar = "SELECT * FROM COMANDAS_2 WHERE ID_COMANDA = " & Text1.Text
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        If tRs.Fields("TIPO") = "C" Then
                            ImpComanda
                        Else
                            Imprimir_Produccion
                        End If
                    Else
                        MsgBox "La comanda o produccion no existe!", vbInformation, "SACC"
                    End If
                    Text1.Text = ""
                Else
                    sBuscar = "SELECT NO_COM_AT FROM VENTAS_DETALLE WHERE ID_VENTA = " & Text1.Text & " GROUP BY NO_COM_AT"
                    Set tRs = cnn.Execute(sBuscar)
                    If (tRs.EOF And tRs.BOF) Then
                        MsgBox "LA VENTA NO FUE ENCONTRADA!", vbExclamation, "SACC"
                    Else
                        Do While Not tRs.EOF
                            Text1.Text = tRs.Fields("NO_COM_AT")
                            ImpComanda
                        Loop
                        tRs.MoveNext
                        Text1.Text = ""
                    End If
                End If
            End If
        End If
        If Option2.Value = True Then
            If Option19.Value Then
                sBuscar = "SELECT ID_VENTA FROM VENTAS WHERE FOLIO = '" & Text1.Text & "'"
                Set tRs = cnn.Execute(sBuscar)
                If (tRs.EOF And tRs.BOF) Then
                    MsgBox "EL FOLIO DE LA FACTURA NO FUE ENCONTRADO, ES POSIBLE QUE ESTE CANCELADO", vbExclamation, "SACC"
                Else
                    Do While Not tRs.EOF
                        Text1.Text = tRs.Fields("ID_VENTA")
                        ReImpVenta
                        If Check1.Value = 1 Then
                            FunRemision
                        End If
                        tRs.MoveNext
                    Loop
                    Text1.Text = ""
                End If
            Else
                If Option20.Value Then
                    sBuscar = "SELECT ID_VENTA FROM VENTAS_DETALLE WHERE NO_COM_AT = 'C" & Text1.Text & "' GROUP BY ID_VENTA"
                    Set tRs = cnn.Execute(sBuscar)
                    If (tRs.EOF And tRs.BOF) Then
                        MsgBox "LA COMANDA NO FUE ENCONTRADA COMO EXTRAIDA EN NOTA DE VENTA!", vbExclamation, "SACC"
                    Else
                        Do While Not tRs.EOF
                            Text1.Text = tRs.Fields("ID_VENTA")
                            ReImpVenta
                            If Check1.Value = 1 Then
                                FunRemision
                            End If
                            tRs.MoveNext
                        Loop
                        Text1.Text = ""
                    End If
                Else
                    ReImpVenta
                    If Check1.Value = 1 Then
                        FunRemision
                    End If
                    Text1.Text = ""
                End If
            End If
        End If
        If Option3.Value = True Then
            ReImpGtia
        End If
        If Option6.Value = True Then
            ImprimeCotiza
        End If
        If Option7.Value = True Then
            If Option17.Value = True Then
                RECIBIDO
            Else
                ReImpEnt
            End If
        End If
        If Option11.Value = True Then
            ReImpValeCaja
        End If
        If Option12.Value = True Then
            ReImpVentaProgramada
        End If
        If Option13.Value = True Then
            FunImpATec
            'ReImpAsistencia
        End If
        If Option15.Value = True Then
            juegopdf
        End If
        If Option14.Value = True Then
            ImpRecep
        End If
        If Option16.Value = True Then
            ImpEntProvVarios
        End If
        If Option21.Value = True Then
            'imrPolizaCheque
            ImrPolizaChequeVarios
        End If
        If Option22.Value = True Then
            ImpRemision
        End If
    Else
        MsgBox "ES NECESARIO DAR UN NUMERO", vbInformation, "SACC"
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
        .ColumnHeaders.Add , , "Clave", 1440
        .ColumnHeaders.Add , , "Cliente", 4040
         'Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 1440
        .ColumnHeaders.Add , , "Descripcion", 3040
        .ColumnHeaders.Add , , "Centidad", 1040
    End With
    Option8.Visible = False
    Option9.Visible = False
    Option10.Visible = False
    Option17.Visible = False
    Frame4.Visible = True
End Sub
Private Sub ReImpGtia()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT GARANTIAS.FECHA, GARANTIAS.ID_VENTA, GARANTIAS.ID_PRODUCTO, GARANTIAS.PRECIO, GARANTIAS.CANTIDAD, VENTAS.NOMBRE FROM GARANTIAS, VENTAS WHERE GARANTIAS.ID_VENTA = " & Text1.Text & " AND GARANTIAS.ID_VENTA = VENTAS.ID_VENTA"
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
        Printer.Print "No. DE VENTA : " & tRs.Fields("ID_VENTA")
        Printer.Print "ATENDIDO POR : " & VarMen.Text1(1).Text
        Printer.Print "CLIENTE : " & tRs.Fields("NOMBRE")
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                          TICKET DE GARANTIA"
        Printer.Print "--------------------------------------------------------------------------------"
        Dim POSY As Integer
        POSY = 2200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 1300
        Printer.Print "Precio unitario"
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
                Printer.Print tRs.Fields("PRECIO")
                Printer.CurrentY = POSY
                Printer.CurrentX = 2900
                Printer.Print tRs.Fields("CANTIDAD")
                tRs.MoveNext
            Loop
        End If
        Printer.Print ""
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "               GRACIAS POR SU COMPRA"
        Printer.Print "           PRODUCTO 100% GARANTIZADO"
        Printer.Print "                APLICA RESTRICCIONES"
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.EndDoc
    Else
        MsgBox "NO SE ENCONTRO LA GARANTIA BUSCADA!", vbInformation, "SACC"
    End If
End Sub
Private Sub ReImpVenta()
    Dim sBuscar As String
    Dim Acum As String
    Dim tRs As ADODB.Recordset
    Dim tRs8 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim POSY As Integer
    Dim Usuario As String
    Dim Usu As String
    Dim Cliente As String
    Dim Sucu As String
    Dim sExibicion As String
    Dim sSubtotal As String
    Dim sIVA As String
    Dim sTotal As String
    Dim sTipoVenta As String
    Dim Facturado As String
    If Text1.Text <> "" Then
        sBuscar = "SELECT ID_USUARIO, NOMBRE, SUCURSAL, FECHA, UNA_EXIBICION, SUBTOTAL, IVA, TOTAL, TIPO_PAGO, FACTURADO FROM VENTAS WHERE ID_VENTA = " & Text1.Text
        Set tRs = cnn.Execute(sBuscar)
        sSubtotal = tRs.Fields("SUBTOTAL")
        sTotal = tRs.Fields("TOTAL")
        sIVA = tRs.Fields("IVA")
        sExibicion = tRs.Fields("UNA_EXIBICION")
        sTipoVenta = tRs.Fields("TIPO_PAGO")
        If Not (tRs.EOF And tRs.BOF) Then
            Usuario = tRs.Fields("ID_USUARIO")
            Cliente = tRs.Fields("NOMBRE")
            Sucu = tRs.Fields("SUCURSAL")
            fech = tRs.Fields("FECHA")
            Facturado = tRs.Fields("FACTURADO")
            tRs.Close
            sBuscar = "SELECT NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & Usuario
            Set tRs = cnn.Execute(sBuscar)
            If tRs.EOF And tRs.BOF Then
                Usuario = VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
            Else
                Usuario = tRs.Fields("NOMBRE") & " " & tRs.Fields("APELLIDOS")
            End If
            tRs.Close
            Acum = "0"
            sBuscar = "SELECT * FROM SUCURSALES WHERE NOMBRE = '" & VarMen.Text4(0).Text & "'"
            Set tRs8 = cnn.Execute(sBuscar)
            sBuscar = "SELECT * FROM EMPRESA"
            Set tRs3 = cnn.Execute(sBuscar)
            '********************************IMPRIMIR TICKET********************************************
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
            Printer.Print VarMen.Text5(0).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
            Printer.Print "R.F.C. " & VarMen.Text5(8).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(tRs8.Fields("CALLE") & " COL. " & tRs8.Fields("COLONIA"))) / 2
            Printer.Print tRs8.Fields("CALLE") & " COL. " & tRs8.Fields("COLONIA")
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & tRs8.Fields("CP"))) / 2
            Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & tRs8.Fields("CP")
            Printer.Print "FECHA : " & fech
            Printer.Print "SUCURSAL : " & Sucu
            Printer.Print "TELEFONO SUCURSAL : " & tRs8.Fields("TELEFONO")
            Printer.Print "No. DE VENTA : " & Text1.Text
            If sTipoVenta = "C" Then
                Printer.Print "FORMA DE PAGO : EFECTIVO"
            Else
                If sTipoVenta = "H" Then
                    Printer.Print "FORMA DE PAGO : TRANSFERENCIA"
                Else
                    If sTipoVenta = "T" Then
                        Printer.Print "FORMA DE PAGO : TARJETA DE CREDITO"
                    Else
                        If sTipoVenta = "E" Then
                            Printer.Print "FORMA DE PAGO : TRANSFERENCIA ELECTRONICA"
                        Else
                            Printer.Print "FORMA DE PAGO : NO INDICADO"
                        End If
                    End If
                End If
            End If
            Printer.Print "ATENDIDO POR : " & Usuario
            Printer.Print "CLIENTE : " & Cliente
            If sExibicion = "N" Then
                Printer.Print "VENTA A CREDITO"
            Else
                Printer.Print "VENTA A CONTADO"
            End If
            Printer.Print "--------------------------------------------------------------------------------"
            Printer.Print "                          NOTA DE FACTURA"
            Printer.Print "--------------------------------------------------------------------------------"
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
            
            sBuscar = "SELECT VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.PRECIO_VENTA, VENTAS_DETALLE.CANTIDAD, VENTAS.SUBTOTAL, VENTAS.IVA, VENTAS.TOTAL FROM VENTAS_DETALLE, VENTAS WHERE VENTAS.ID_VENTA = VENTAS_DETALLE.ID_VENTA AND VENTAS.ID_VENTA = " & Text1.Text
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                sSubtotal = Format(tRs.Fields("SUBTOTAL"), "###,###,##0.00")
                sTotal = Format(tRs.Fields("TOTAL"), "###,###,##0.00")
                sIVA = Format(tRs.Fields("IVA"), "###,###,##0.00")
                Do While Not tRs.EOF
                    POSY = POSY + 200
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 100
                    Printer.Print tRs.Fields("ID_PRODUCTO")
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 1900
                    Printer.Print tRs.Fields("CANTIDAD")
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 2900
                    Printer.Print Format(CDbl(tRs.Fields("PRECIO_VENTA")), "###,###,##0.00")
                    Acum = CDbl(Acum) + CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("CANTIDAD"))
                    tRs.MoveNext
                Loop
            End If
            Printer.Print ""
            Printer.Print "SUBTOTAL : " & Format(CDbl(sSubtotal), "###,###,##0.00")
            Printer.Print "IVA              : " & Format(CDbl(sIVA), "###,###,##0.00")
            Printer.Print "TOTAL        : " & Format(CDbl(sTotal), "###,###,##0.00")
            Printer.Print ""
            If Facturado = "2" Then
                Printer.Print "--------------------------------------------------------------------------------"
                Printer.Print "                       VENTA CANCELADA"
                Printer.Print "--------------------------------------------------------------------------------"
            Else
                Printer.Print "--------------------------------------------------------------------------------"
                Printer.Print "               GRACIAS POR SU COMPRA"
                Printer.Print "LA GARANTÍA SERÁ VÁLIDA HASTA 15 DIAS"
                Printer.Print "     DESPUES DE HABER EFECTUADO SU "
                Printer.Print "                                COMPRA"
                Printer.Print "--------------------------------------------------------------------------------"
            End If
            Printer.EndDoc
        Else
            MsgBox "LA VENTA NO EXISTE!", vbInformation, "SACC"
        End If
    Else
        MsgBox "DEBE DAR EL NUMERO DE VENTA!", vbInformation, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If Option1.Value = True Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM VsReImpCOMANDA WHERE ID_COMANDA = " & Item
    End If
    If Option2.Value = True Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM VsReImpVENTA WHERE ID_VENTA = " & Item
    End If
    If Option3.Value = True Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM GARANTIAS WHERE ID_VENTA = " & Item
    End If
    Set tRs = cnn.Execute(sBuscar)
    ListView2.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Option1_Click()
    Option8.Visible = False
    Option9.Visible = False
    Option10.Visible = False
    Option17.Visible = False
    Option20.Value = True
    Frame4.Visible = True
    Check1.Visible = False
    Label2.Visible = False
    Text3.Visible = False
    ListView2.Visible = True
End Sub
Private Sub Option11_Click()
    Option8.Visible = False
    Option9.Visible = False
    Option10.Visible = False
    Option17.Visible = False
    Frame4.Visible = False
    Check1.Visible = False
    Label2.Visible = False
    Text3.Visible = False
    ListView2.Visible = True
End Sub
Private Sub Option12_Click()
    Option8.Visible = False
    Option9.Visible = False
    Option10.Visible = False
    Option17.Visible = False
    Frame4.Visible = False
    Check1.Visible = False
    Label2.Visible = False
    Text3.Visible = False
    ListView2.Visible = True
End Sub
Private Sub Option13_Click()
    Option8.Visible = False
    Option9.Visible = False
    Option10.Visible = False
    Option17.Visible = False
    Frame4.Visible = False
    Check1.Visible = False
    Label2.Visible = False
    Text3.Visible = False
    ListView2.Visible = True
End Sub
Private Sub Option14_Click()
    Option8.Visible = False
    Option9.Visible = False
    Option10.Visible = False
    Option17.Visible = False
    Frame4.Visible = False
    Check1.Visible = False
    Label2.Visible = False
    Text3.Visible = False
    ListView2.Visible = True
End Sub
Private Sub Option15_Click()
    Option8.Visible = False
    Option9.Visible = False
    Option10.Visible = False
    Option17.Visible = False
    Frame4.Visible = False
    Check1.Visible = False
    Label2.Visible = False
    Text3.Visible = False
    ListView2.Visible = True
End Sub
Private Sub Option16_Click()
    Option8.Visible = False
    Option9.Visible = False
    Option10.Visible = False
    Option17.Visible = False
    Frame4.Visible = False
    Check1.Visible = False
    Label2.Visible = False
    Text3.Visible = False
    ListView2.Visible = True
End Sub
Private Sub Option2_Click()
    Option8.Visible = False
    Option9.Visible = False
    Option10.Visible = False
    Option17.Visible = False
    Option18.Value = True
    Frame4.Visible = True
    Check1.Visible = True
    Label2.Visible = False
    Text3.Visible = False
    ListView2.Visible = True
End Sub
Private Sub Option21_Click()
    Option8.Visible = True
    Option9.Visible = True
    Option10.Visible = True
    Option17.Visible = True
    Option20.Value = False
    Frame4.Visible = False
    Check1.Visible = False
    Label2.Visible = False
    Text3.Visible = False
    ListView2.Visible = True
End Sub
Private Sub Option22_Click()
    Option8.Visible = False
    Option9.Visible = False
    Option10.Visible = False
    Option17.Visible = False
    Frame4.Visible = False
    Check1.Visible = False
    Label2.Visible = True
    Text3.Visible = True
    ListView2.Visible = False
End Sub
Private Sub Option3_Click()
    Option8.Visible = False
    Option9.Visible = False
    Option10.Visible = False
    Option17.Visible = False
    Frame4.Visible = False
    Check1.Visible = False
    Label2.Visible = False
    Text3.Visible = False
    ListView2.Visible = True
End Sub
Private Sub Option6_Click()
    Option8.Visible = False
    Option9.Visible = False
    Option10.Visible = False
    Option17.Visible = False
    Frame4.Visible = False
    Check1.Visible = False
    Label2.Visible = False
    Text3.Visible = False
    ListView2.Visible = True
End Sub
Private Sub Option7_Click()
    Option8.Visible = True
    Option9.Visible = True
    Option10.Visible = True
    Option17.Visible = True
    Option20.Value = False
    Frame4.Visible = False
    Check1.Visible = False
    Label2.Visible = False
    Text3.Visible = False
    ListView2.Visible = True
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    If (Option1.Value And Option19.Value) Or (Option2.Value And Option19.Value) Then
        Valido = "1234567890ABCDEFGHIJKLMNÑOPQRSTUVWXYZ"
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
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Option1.Value = True Or Option2.Value = True Then
            Dim sBuscar As String
            Dim tRs As ADODB.Recordset
            Dim tLi As ListItem
            sBuscar = "SELECT NOMBRE, "
            If Option1.Value = True Then
                sBuscar = sBuscar & "ID_COMANDA FROM VsReImpCOMANDA WHERE "
            End If
            If Option2.Value = True Or Option3.Value = True Then
                sBuscar = sBuscar & "ID_VENTA FROM VsReImpVENTA WHERE "
            End If
            If Option4.Value = True Then
                sBuscar = sBuscar & "NOMBRE LIKE '%" & Text2.Text & "%'"
            Else
                sBuscar = sBuscar & "ID_PRODUCTO LIKE '%" & Text2.Text & "%'"
            End If
            If Option1.Value = True Then
                sBuscar = sBuscar & " GROUP BY ID_COMANDA, NOMBRE"
            End If
            If Option2.Value = True Or Option3.Value = True Then
                sBuscar = sBuscar & " GROUP BY ID_VENTA, NOMBRE"
            End If
            ListView1.ListItems.Clear
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not (tRs.EOF)
                    If Option1.Value = True Then
                        Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_COMANDA"))
                    Else
                        Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
                    End If
                    If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                    tRs.MoveNext
                Loop
            End If
        End If
    End If
End Sub
Private Sub ImpComanda()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT * FROM VsReImpCOMANDA WHERE ID_COMANDA = " & Text1.Text
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
        Printer.Print "FECHA : " & tRs.Fields("FECHA_INICIO")
        Printer.Print "SUCURSAL : " & Trim(tRs.Fields("Sucursal"))
        Printer.Print "TELEFONO SUCURSAL : " & Trim(tRs.Fields("TelSuc"))
        Printer.Print "No. DE COMANDA : " & Text1.Text
        Printer.Print "ATENDIDO POR : " & Trim(tRs.Fields("Usuario")) & " " & Trim(tRs.Fields("APELLIDOS"))
        Printer.Print "CLIENTE : " & tRs.Fields("NOMBRE")
        Printer.Print "TELEFONO : " & tRs.Fields("TELEFONO")
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                           RECARGA DE TINTA"
        Dim POSY As Integer
        POSY = 2400
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3000
        Printer.Print "Cant."
        Do While Not tRs.EOF
            If Mid(tRs.Fields("ID_PRODUCTO"), 3, 1) = "I" Then
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print tRs.Fields("ID_PRODUCTO")
                Printer.CurrentY = POSY
                Printer.CurrentX = 2900
                Printer.Print tRs.Fields("CANTIDAD")
            End If
            tRs.MoveNext
        Loop
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                           RECARGA DE TONER"
        POSY = POSY + 600
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3000
        Printer.Print "Cant."
        tRs.MoveFirst
        Do While Not tRs.EOF
            If Mid(tRs.Fields("ID_PRODUCTO"), 3, 1) <> "I" Then
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print tRs.Fields("ID_PRODUCTO")
                Printer.CurrentY = POSY
                Printer.CurrentX = 2900
                Printer.Print tRs.Fields("CANTIDAD")
            End If
            tRs.MoveNext
        Loop
        Printer.Print ""
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "               GRACIAS POR SU COMPRA"
        Printer.Print "           PRODUCTO 100% GARANTIZADO"
        Printer.Print "SI SU PEDIDO NO SE RECOGE EN 30 DIAS LA"
        Printer.Print "   LA EMPRESA NO SE HACE RESPONSABLE"
        Printer.Print "                APLICA RESTRICCIONES"
        Printer.Print ""
        Printer.Print "Conserve su ticket"
        Printer.Print "El cobro se hará hasta la entrega del cartucho lleno"
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.EndDoc
    Else
        MsgBox "NO SE ENCONTRO LA COMANDA BUSCADA!", vbInformation, "SACC"
    End If
End Sub
Private Sub juegopdf()
   Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim loca As Integer
    Dim sBuscar As String
    Dim tRs1 As ADODB.Recordset
    Dim tRs8 As ADODB.Recordset
    Dim PosVer As Integer
    Dim Posabo As Integer
    Dim tRs  As ADODB.Recordset
    Dim tRs2  As ADODB.Recordset
    Dim tRs4  As ADODB.Recordset
    Set oDoc = New cPDF
    Dim sumdeuda As Double
    Dim sumIndi As Double
    Dim sNombre As String
    Dim sumabonos As Double
    Dim sumIndiabonos As Double
    Dim ConPag As Integer
    ConPag = 1
    If Not oDoc.PDFCreate(App.Path & "\Juegoreparacion.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
' Encabezado del reporte
    oDoc.NewPage A4_Vertical
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
    oDoc.WTextBox 90, 200, 20, 250, "Juego de Reparacion Para Comandas", "F2", 10, hCenter
    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
' Encabezado de pagina
    oDoc.WTextBox 100, 90, 40, 200, "Cantidad en Juego", "F2", 10, hCenter
    oDoc.WTextBox 100, 250, 40, 300, "Cantidad  en comanda", "F2", 10, hLeft
' Cuerpo del reporte
    sBuscar = "SELECT ID_PRODUCTO, CANTIDAD, ID_COMANDA FROM COMANDAS_DETALLES_2 WHERE ID_COMANDA = " & Text1.Text & " AND ESTADO_ACTUAL = 'A'"
    Set tRs8 = cnn.Execute(sBuscar)
     If Not (tRs8.EOF And tRs8.BOF) Then
        If MsgBox("ESTA SEGURO QUE  QUIERE IMPRIMIR EL JUEGO DE REPARACION. " & Text1.Text & "?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
            sBuscar = "SELECT * FROM  IMPCOMAN WHERE ID_COMANDA = " & Text1.Text
            Set tRs4 = cnn.Execute(sBuscar)
            If Not (tRs4.EOF And tRs4.BOF) Then
                oDoc.WTextBox 100, 500, 40, 300, "Reimpresion", "F2", 10, hLeft
            Else
                sBuscar = "INSERT INTO IMPCOMAN (ID_COMANDA,FECHA) VALUES ('" & Text1.Text & "','" & Format(Date, "dd/mm/yyyy") & "');"
                cnn.Execute (sBuscar)
                oDoc.WTextBox 100, 500, 40, 300, "Impresion", "F2", 10, hLeft
            End If
            sumdeuda = 0
            sBuscar = "SELECT ID_PRODUCTO, CANTIDAD, CANTIDAD_NO_SIRVIO, ID_COMANDA FROM COMANDAS_DETALLES_2 WHERE ID_COMANDA = " & Text1.Text & " AND (CANTIDAD - CANTIDAD_NO_SIRVIO) > 0"
            Set tRs = cnn.Execute(sBuscar)
            oDoc.WTextBox 130, 30, 30, 60, "COMANDA :", "F2", 10, hLeft
            oDoc.WTextBox 130, 100, 30, 40, tRs.Fields("ID_COMANDA"), "F2", 10, hLeft
            Posi = 130
            oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 100
            oDoc.WLineTo 580, 100
            oDoc.LineStroke
            oDoc.MoveTo 10, 125
            oDoc.WLineTo 580, 125
            oDoc.LineStroke
            Posi = Posi + 15
            If Not (tRs.EOF And tRs.BOF) Then
                 Do While Not (tRs.EOF)
                    sBuscar = "SELECT ID_REPARACION, ID_PRODUCTO, CANTIDAD FROM JR_TEMPORALES WHERE ID_REPARACION = '" & tRs.Fields("ID_PRODUCTO") & "' AND ID_COMANDA = " & Text1.Text & " AND (CANTIDAD * " & tRs.Fields("CANTIDAD_NO_SIRVIO") * tRs.Fields("CANTIDAD") & ") > 0 GROUP BY ID_REPARACION, ID_PRODUCTO, CANTIDAD"
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
                         Do While Not (tRs1.EOF)
                            If sNombre <> tRs1.Fields("ID_REPARACION") Then
                                Posi = Posi + 15
                                oDoc.WTextBox Posi, 30, 40, 200, tRs1.Fields("ID_REPARACION"), "F2", 11, hLeft
                                oDoc.WTextBox Posi, 150, 50, 100, tRs.Fields("CANTIDAD"), "F2", 11, hLeft
                            End If
                            Posi = Posi + 10
                            oDoc.WTextBox Posi, 30, 40, 300, tRs1.Fields("ID_PRODUCTO"), "F2", 9, hLeft
                            oDoc.WTextBox Posi, 200, 40, 80, tRs1.Fields("CANTIDAD"), "F2", 9, hLeft
                            sumdeuda = CDbl(tRs.Fields("CANTIDAD") * tRs1.Fields("CANTIDAD"))
                            oDoc.WTextBox Posi, 300, 40, 300, sumdeuda, "F2", 10, hLeft
                            sNombre = tRs1.Fields("ID_REPARACION")
                            sumdeuda = 0
                            tRs1.MoveNext
                            If Posi >= 730 Then
                                oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                                ConPag = ConPag + 1
                                oDoc.NewPage A4_Vertical
                                ' Encabezado del reporte
                                Posi = 140
                                oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                                oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                                oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                                oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                                oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Pagar", "F2", 10, hCenter
                                oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
                                oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
                                ' Encabezado de pagina
                                oDoc.WTextBox 100, 20, 30, 40, "Id", "F2", 10, hCenter
                                oDoc.WTextBox 100, 30, 30, 80, "Factura", "F2", 10, hCenter
                                oDoc.WTextBox 100, 50, 50, 160, "Fecha", "F2", 10, hCenter
                                oDoc.WTextBox 100, 90, 40, 200, "Total-Fac", "F2", 10, hCenter
                                oDoc.WTextBox 100, 250, 40, 300, "Abono", "F2", 10, hLeft
                                oDoc.WTextBox 100, 350, 40, 200, "Pendiente", "F2", 10, hLeft
                                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                                oDoc.MoveTo 10, 100
                                oDoc.WLineTo 580, 100
                                oDoc.LineStroke
                                oDoc.MoveTo 10, 125
                                oDoc.WLineTo 580, 125
                                oDoc.LineStroke
                            End If
                        Loop
                    Else
                        sBuscar = "SELECT ID_REPARACION, ID_PRODUCTO, CANTIDAD FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & tRs.Fields("ID_PRODUCTO") & "' GROUP BY ID_REPARACION, ID_PRODUCTO, CANTIDAD"
                        Set tRs1 = cnn.Execute(sBuscar)
                        If Not (tRs1.EOF And tRs1.BOF) Then
                            Do While Not (tRs1.EOF)
                                If sNombre <> tRs1.Fields("ID_REPARACION") Then
                                    Posi = Posi + 15
                                    oDoc.WTextBox Posi, 30, 40, 200, tRs1.Fields("ID_REPARACION"), "F2", 11, hLeft
                                    oDoc.WTextBox Posi, 150, 50, 100, tRs.Fields("CANTIDAD"), "F2", 11, hLeft
                                End If
                                Posi = Posi + 10
                                oDoc.WTextBox Posi, 30, 40, 300, tRs1.Fields("ID_PRODUCTO"), "F2", 9, hLeft
                                oDoc.WTextBox Posi, 200, 40, 80, tRs1.Fields("CANTIDAD"), "F2", 9, hLeft
                                sumdeuda = CDbl(tRs.Fields("CANTIDAD") * tRs1.Fields("CANTIDAD"))
                                oDoc.WTextBox Posi, 300, 40, 300, sumdeuda, "F2", 10, hLeft
                                sNombre = tRs1.Fields("ID_REPARACION")
                                sumdeuda = 0
                                tRs1.MoveNext
                                If Posi >= 730 Then
                                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                                    ConPag = ConPag + 1
                                    oDoc.NewPage A4_Vertical
                                    ' Encabezado del reporte
                                    Posi = 140
                                    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                                    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                                    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                                    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                                    oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Pagar", "F2", 10, hCenter
                                    oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
                                    oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
                                    ' Encabezado de pagina
                                    oDoc.WTextBox 100, 20, 30, 40, "Id", "F2", 10, hCenter
                                    oDoc.WTextBox 100, 30, 30, 80, "Factura", "F2", 10, hCenter
                                    oDoc.WTextBox 100, 50, 50, 160, "Fecha", "F2", 10, hCenter
                                    oDoc.WTextBox 100, 90, 40, 200, "Total-Fac", "F2", 10, hCenter
                                    oDoc.WTextBox 100, 250, 40, 300, "Abono", "F2", 10, hLeft
                                    oDoc.WTextBox 100, 350, 40, 200, "Pendiente", "F2", 10, hLeft
                                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                                    oDoc.MoveTo 10, 100
                                    oDoc.WLineTo 580, 100
                                    oDoc.LineStroke
                                    oDoc.MoveTo 10, 125
                                    oDoc.WLineTo 580, 125
                                    oDoc.LineStroke
                                End If
                            Loop
                        End If
                    End If
                    Posi = Posi + 15
                    tRs.MoveNext
                    If Posi >= 730 Then
                        oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                        ConPag = ConPag + 1
                        oDoc.NewPage A4_Vertical
                        ' Encabezado del reporte
                        Posi = 140
                        oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                        oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                        oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                        oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                        oDoc.WTextBox 90, 200, 20, 250, "Reporte de Cuentas por Pagar", "F2", 10, hCenter
                        oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
                        oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
                        ' Encabezado de pagina
                        oDoc.WTextBox 100, 20, 30, 40, "Id", "F2", 10, hCenter
                        oDoc.WTextBox 100, 30, 30, 80, "Factura", "F2", 10, hCenter
                        oDoc.WTextBox 100, 50, 50, 160, "Fecha", "F2", 10, hCenter
                        oDoc.WTextBox 100, 90, 40, 200, "Total-Fac", "F2", 10, hCenter
                        oDoc.WTextBox 100, 250, 40, 300, "Abono", "F2", 10, hLeft
                        oDoc.WTextBox 100, 350, 40, 200, "Pendiente", "F2", 10, hLeft
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, 100
                        oDoc.WLineTo 580, 100
                        oDoc.LineStroke
                        oDoc.MoveTo 10, 125
                        oDoc.WLineTo 580, 125
                        oDoc.LineStroke
                    End If
                Loop
                Posi = Posi + 30
                Cont = Cont + 1
            End If
        Else
            MsgBox "ESTA  COMANDA ESTA  COMO DAÑADA  O YA ESTA EN PROCESO", vbInformation, "SACC"
        End If
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Function Imprimir_Produccion()
    Dim NRegistros As Integer
    Dim Con As Integer
    Dim POSY As Integer
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    sBuscar = "SELECT * FROM VsReImpCOMANDA WHERE ID_COMANDA = " & Text1.Text
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Printer.Print "        " & VarMen.Text5(0).Text
        Printer.Print "           ORDEN DE PRODUCCIÓN"
        Printer.Print "FECHA : " & Now
        Printer.Print "No. DE ORDEN DE PRODUCCCION : " & Text1.Text
        Printer.Print "ORDEN HECHA POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
        Printer.Print "COMENTARIO : " & Text1.Text
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                           ORDEN DE TINTA"
        POSY = 2200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3000
        Printer.Print "Cant."
        Do While Not tRs.EOF
            If Mid(tRs.Fields("ID_PRODUCTO"), 3, 1) = "I" Then
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print tRs.Fields("ID_PRODUCTO")
                Printer.CurrentY = POSY
                Printer.CurrentX = 2900
                Printer.Print tRs.Fields("CANTIDAD")
            End If
            tRs.MoveNext
        Loop
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                           ORDEN DE TONER"
        POSY = POSY + 600
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3000
        Printer.Print "Cant."
        tRs.MoveFirst
        Do While Not tRs.EOF
            If Mid(tRs.Fields("ID_PRODUCTO"), 3, 1) = "T" Then
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print tRs.Fields("ID_PRODUCTO")
                Printer.CurrentY = POSY
                Printer.CurrentX = 2900
                Printer.Print tRs.Fields("CANTIDAD")
            End If
            tRs.MoveNext
        Loop
        Printer.Print ""
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print ""
        Printer.EndDoc
    End If
End Function
Private Sub ImprimeCotiza()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT NOMBRE, FECHA, SUCURSAL, USUARIO, APELLIDOS, ID_COTIZA_CLIEN, ID_PRODUCTO, DESCRIPCION, PRECIO_VENTA, CANTIDAD, SUBTOTAL, TOTAL FROM VsCotizaCliente WHERE ID_COTIZA_CLIEN = " & Text1.Text
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        CommonDialog1.Flags = 64
        CommonDialog1.CancelError = True
        CommonDialog1.ShowPrinter
        Dim cant As String
        Dim P_ven As String
        Dim IdVentAut As String
        Dim POSY As Integer
        '********************************IMPRIMIR TICKET********************************************
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print ""
        Printer.Print ""
        Printer.Print "     CLIENTE : " & tRs.Fields("NOMBRE")
        Printer.Print "     FECHA : " & tRs.Fields("FECHA")
        Printer.Print "     SUCURSAL : " & tRs.Fields("SUCURSAL")
        Printer.Print "     ATENDIDO POR : " & tRs.Fields("USUARIO") & " " & tRs.Fields("APELLIDOS")
        Printer.Print "     No. FOLIO   : " & tRs.Fields("ID_COTIZA_CLIEN")
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("COTIZACION")) / 2
        Printer.Print "COTIZACION"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        POSY = 3400
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 1600
        Printer.Print "Descripcion"
        Printer.CurrentY = POSY
        Printer.CurrentX = 9000
        Printer.Print "Cant."
        Printer.CurrentY = POSY
        Printer.CurrentX = 10500
        Printer.Print "P/U"
        Do While Not tRs.EOF
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print tRs.Fields("ID_PRODUCTO")
            Printer.CurrentY = POSY
            Printer.CurrentX = 1600
            Printer.Print tRs.Fields("Descripcion")
            Printer.CurrentY = POSY
            Printer.CurrentX = 9130
            Printer.Print tRs.Fields("CANTIDAD")
            Printer.CurrentY = POSY
            Printer.CurrentX = 10200
            Printer.Print Format(Val(tRs.Fields("PRECIO_VENTA")) / Val(tRs.Fields("CANTIDAD")), "###,###,##0.00")
            P_ven = Format(Val(P_ven) + Val(tRs.Fields("PRECIO_VENTA")), "###,###,##0.00")
            If POSY >= 14200 Then
                Printer.NewPage
                Printer.Print ""
                Printer.Print ""
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                Printer.Print VarMen.Text5(0).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print ""
                Printer.Print ""
                Printer.Print "     CLIENTE : " & tRs.Fields("NOMBRE")
                Printer.Print "     FECHA : " & tRs.Fields("FECHA")
                Printer.Print "     SUCURSAL : " & tRs.Fields("SUCURSAL")
                Printer.Print "     ATENDIDO POR : " & tRs.Fields("USUARIO") & " " & tRs.Fields("APELLIDOS")
                Printer.Print "     No. FOLIO   : " & tRs.Fields("ID_COTIZA_CLIEN")
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("COTIZACION")) / 2
                Printer.Print "COTIZACION"
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                POSY = 3400
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print "Producto"
                Printer.CurrentY = POSY
                Printer.CurrentX = 1600
                Printer.Print "Descripcion"
                Printer.CurrentY = POSY
                Printer.CurrentX = 9000
                Printer.Print "Cant."
                Printer.CurrentY = POSY
                Printer.CurrentX = 10500
                Printer.Print "P/U"
            End If
            tRs.MoveNext
        Loop
        Printer.Print ""
        tRs.MoveFirst
        Printer.CurrentX = 8700
        Printer.Print "         S U B T O T A L :    " & tRs.Fields("SUBTOTAL")
        Printer.CurrentX = 8700
        Printer.Print "         I V A                   :    " & Format(Val(tRs.Fields("SUBTOTAL")) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        Printer.CurrentX = 8700
        Printer.Print "         T O T A L           :    " & tRs.Fields("TOTAL")
        Printer.Print ""
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "                                                                                  COTIZACIÓN SUJETA A CAMBIOS SIN PREVIO AVISO."
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "Comentarios : "
        Printer.EndDoc
        IdCotizacion = ""
        CommonDialog1.Copies = 1
    End If
    Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub ReImpEnt()
On Error GoTo ManejaError
    Dim Total As Double
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim VarTipo As String
    Dim sFolios As String
    If Option8.Value = True Then
        VarTipo = "N"
    End If
    If Option9.Value = True Then
        VarTipo = "I"
    End If
    If Option10.Value = True Then
        VarTipo = "X"
    End If
    Total = 0
    sBuscar = "SELECT * FROM VsOrdenReimp WHERE NUM_ORDEN = " & Text1.Text & " AND TIPO = '" & VarTipo & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        sBuscar = "SELECT ID_ENTRADA FROM VsOrdenReimp WHERE NUM_ORDEN = " & Text1.Text & " AND TIPO = '" & VarTipo & "'"
        Set tRs1 = cnn.Execute(sBuscar)
        Do While Not tRs1.EOF
            sFolios = sFolios & tRs1.Fields("ID_ENTRADA") & ", "
            tRs1.MoveNext
        Loop
        CommonDialog1.Flags = 64
        CommonDialog1.CancelError = True
        CommonDialog1.ShowPrinter
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        If Option8.Value Then
            Printer.Print "             ORDEN : " & Text1.Text & " Nacional"
        End If
        If Option9.Value Then
            Printer.Print "             ORDEN : " & Text1.Text & " Internacional"
        End If
        If Option10.Value Then
            Printer.Print "             ORDEN : " & Text1.Text & " Indirecta"
        End If
        Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
        Printer.Print "             SUCURSAL : BODEGA"
        Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
        Printer.Print "             FOLIO: " & sFolios 'tRs.Fields("ID_ENTRADA")
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO")) / 2
        Printer.Print "COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("NOMBRE DEL PROVEEDOR:  " & tRs.Fields("NOMBRE"))) / 2
        Printer.Print "NOMBRE DEL PROVEEDOR:  " & tRs.Fields("NOMBRE")
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Dim POSY As Integer
        POSY = 3800
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Clave del Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 2000
        Printer.Print "Cant. Registrada"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3300
        Printer.Print "Precio"
        Printer.CurrentY = POSY
        Printer.CurrentX = 4100
        Printer.Print "Sucursal"
        Printer.CurrentY = POSY
        Printer.CurrentX = 5300
        Printer.Print "Entrada"
        Printer.CurrentY = POSY
        Printer.CurrentX = 6100
        Printer.Print "No. Orden"
        Printer.CurrentY = POSY
        Printer.CurrentX = 7000
        Printer.Print "Factura"
        POSY = POSY + 200
        sBuscar = "SELECT * FROM VsEntradaDetalle WHERE ID_ORDEN_COMPRA = " & tRs.Fields("ID_ORDEN_COMPRA") & " GROUP BY  ID_ORDEN_COMPRA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO, FECHA, SURTIDO, ID_ENTRADA, FACT_PROVE, NUM_ORDEN"
        Set tRs1 = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs1.EOF
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print tRs1.Fields("ID_PRODUCTO")
                Printer.CurrentY = POSY
                Printer.CurrentX = 2300
                Printer.Print Format(tRs1.Fields("SURTIDO"), "0.00")
                Printer.CurrentY = POSY
                Printer.CurrentX = 3300
                Printer.Print Format(tRs1.Fields("PRECIO"), "0.00")
                Printer.CurrentY = POSY
                Printer.CurrentX = 4100
                Printer.Print "BODEGA"
                Printer.CurrentY = POSY
                Printer.CurrentX = 5500
                Printer.Print tRs1.Fields("ID_ENTRADA")
                Printer.CurrentY = POSY
                Printer.CurrentX = 6100
                Printer.Print tRs1.Fields("NUM_ORDEN")
                Printer.CurrentY = POSY
                Printer.CurrentX = 7000
                Printer.Print tRs1.Fields("FACT_PROVE")
                tRs1.MoveNext
                If POSY >= 14200 Then
                    POSY = 100
                    Printer.NewPage
                    Printer.Print ""
                    Printer.Print ""
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                    Printer.Print VarMen.Text5(0).Text
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    If Option8.Value Then
                        Printer.Print "             ORDEN : " & Text1.Text & " Nacional"
                    End If
                    If Option9.Value Then
                        Printer.Print "             ORDEN : " & Text1.Text & " Internacional"
                    End If
                    If Option8.Value Then
                        Printer.Print "             ORDEN : " & Text1.Text & " Indirecta"
                    End If
                    Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
                    Printer.Print "             SUCURSAL : BODEGA"
                    Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
                    Printer.Print "             FOLIO: " & sFolios 'tRs.Fields("ID_ENTRADA")
                    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO")) / 2
                    Printer.Print "COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO"
                    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth("NOMBRE DEL PROVEEDOR:  " & Text1.Text)) / 2
                    Printer.Print "NOMBRE DEL PROVEEDOR:  " & tRs.Fields("NOMBRE")
                    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    POSY = 3800
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 100
                    Printer.Print "Clave del Producto"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 2000
                    Printer.Print "Cant. Registrada"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 3300
                    Printer.Print "Precio"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 4100
                    Printer.Print "Sucursal"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 5300
                    Printer.Print "Entrada"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 6100
                    Printer.Print "No. Orden"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 7000
                    Printer.Print "Factura"
                    POSY = POSY + 200
                End If
            Loop
        End If
        Printer.Print ""
        Printer.Print "             Total = " & Format(tRs.Fields("TOTAL"), "###,###,##0.00")
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.EndDoc
        CommonDialog1.Copies = 1
    Else
        MsgBox "NO SE ENCONTRO EL REGISTRO DE LA ORDEN DE COMPRA O ENTRADA", vbInformation, "SACC"
        sBuscar = "SELECT CONFIRMADA FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text1.Text & " AND TIPO = '" & VarTipo & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            If tRs.Fields("CONFIRMADA") = "P" Then
                MsgBox "LA ORDEN ESTA PENDIENTE DE APROBAR"
            End If
            If tRs.Fields("CONFIRMADA") = "S" Then
                MsgBox "LA ORDEN ESTA EN ESPERA DE SER CERRADA"
            End If
            If tRs.Fields("CONFIRMADA") = "N" Then
                MsgBox "LA ORDEN ESTA EN PREORDEN"
            End If
        End If
    End If
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub ReImpValeCaja()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT ID_VALE, ID_VENTA, IMPORTE, FECHA, ID_PRODUCTO, CANTIDAD, PRECIO_VENTA FROM VsValeCaja WHERE ID_VALE =" & Text1.Text
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
    Else
        MsgBox "EL VALE DE CAJA NO EXISTE EN ESE NUMERO!", vbInformation, "SACC"
    End If
End Sub
Private Sub ReImpVentaProgramada()
On Error GoTo CancelaError
    If Text1.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim POSY As Integer
        POSY = 2800
        Dim NomClien As String
        Dim NoPed As String
        Dim fecha As String
        Dim NoOrden As String
        sBuscar = "SELECT * FROM VsVentProg WHERE NO_PEDIDO = " & Text1.Text
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            CommonDialog1.Flags = 64
            CommonDialog1.CancelError = True
            CommonDialog1.ShowPrinter
            Printer.Print ""
            Printer.Print ""
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
            Printer.Print VarMen.Text5(0).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
            Printer.Print "R.F.C. " & VarMen.Text5(8).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
            Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
            Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print ""
            Printer.Print "     Cliente : " & tRs.Fields("NOMBRE")
            Printer.Print "     No. Venta Programada : " & tRs.Fields("NO_PEDIDO") & "                                No. Orden : " & tRs.Fields("NO_ORDEN")
            Printer.Print "     VENTA PROGRAMADA"
            Printer.Print "     Fecha de Entrega : " & tRs.Fields("FECHA")
            Printer.Print "     No. de Orden : " & Text1.Text
            NomClien = tRs.Fields("NOMBRE")
            NoPed = tRs.Fields("NO_PEDIDO")
            fecha = tRs.Fields("FECHA")
            NoOrden = tRs.Fields("NO_ORDEN")
            Printer.Print ""
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print ""
            Printer.Print ""
            sBuscar = "SELECT * FROM VsVentProgDet WHERE NO_PEDIDO = " & Text1.Text
            Set tRs = cnn.Execute(sBuscar)
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print "Producto"
            Printer.CurrentY = POSY
            Printer.CurrentX = 2200
            Printer.Print "Descripcion"
            Printer.CurrentY = POSY
            Printer.CurrentX = 8800
            Printer.Print "C. PEDIDA"
            Printer.CurrentY = POSY
            Printer.CurrentX = 10000
            Printer.Print "C. PENDIENTE"
            POSY = POSY + 400
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not (tRs.EOF)
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 100
                    Printer.Print tRs.Fields("ID_PRODUCTO")
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 2200
                    Printer.Print tRs.Fields("Descripcion")
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 8800
                    Printer.Print tRs.Fields("CANTIDAD_PEDIDA")
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 10000
                    Printer.Print tRs.Fields("CANTIDAD_PENDIENTE")
                    tRs.MoveNext
                    POSY = POSY + 200
                    If POSY >= 14200 Then
                        Printer.NewPage
                        POSY = 2800
                        Printer.Print ""
                        Printer.Print ""
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                        Printer.Print VarMen.Text5(0).Text
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                        Printer.Print ""
                        Printer.Print "     Cliente : " & NomClien
                        Printer.Print "     No. Pedido : " & NoPed & "                                No. Orden : " & NoOrden
                        Printer.Print "     VENTA PROGRAMADA"
                        Printer.Print "     Fecha de Entrega : " & fecha
                        Printer.Print "     No. de Orden : " & Text1.Text
                        Printer.Print ""
                        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                        Printer.Print ""
                        Printer.Print ""
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 100
                        Printer.Print "Producto"
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 2200
                        Printer.Print "Descripcion"
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 8800
                        Printer.Print "C. PEDIDA"
                        Printer.CurrentY = POSY
                        Printer.CurrentX = 10000
                        Printer.Print "C. PENDIENTE"
                        POSY = POSY + 400
                    End If
                Loop
            End If
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print "FIN DEL LISTADO"
            Printer.EndDoc
            CommonDialog1.Copies = 1
        Else
            MsgBox "NO EXISTE UNA VENTA CON ESE FOLIO!", vbInformation, "SACC"
        End If
    Else
        MsgBox "DEBE DAR EL NUMERO DE VENTA A IMPRIMIR!", vbInformation, "SACC"
    End If
    Exit Sub
CancelaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub ReImpAsistencia()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT * FROM VsASISTENCIA_TECNICA WHERE ID_AS_TEC = " & Text1.Text
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        On Error GoTo ManejaError
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "FECHA : " & Format(tRs.Fields("FECHA_CAPTURA"))
        Printer.Print "SUCURSAL : " & tRs.Fields("SUCURSAL")
        Printer.Print "TELEFONO SUCURSAL : " & tRs.Fields("TEL_SUC")
        Printer.Print "No. DE ASISTENCIA : " & Text1.Text
        Printer.Print "ATENDIDO POR : " & tRs.Fields("USUARIO") & " " & tRs.Fields("APELLIDOS")
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                       ASISTENCIA TECNICA"
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "Cliente : " & tRs.Fields("NOMBRE")
        Printer.Print "Telefono : " & tRs.Fields("TELEFONO")
        Printer.Print "Calle : " & tRs.Fields("DIRECCION") & " # " & tRs.Fields("NUM_EXT") & "-" & tRs.Fields("NUM_INT")
        Printer.Print "Colonia : " & tRs.Fields("COLONIA")
        Printer.Print "Fecha a atender : " & tRs.Fields("FECHA_DEBE_ATENDER")
        Printer.Print ""
        Printer.Print "Marca : " & tRs.Fields("MARCA")
        Printer.Print "Modelo : " & tRs.Fields("MODELO")
        Printer.Print "Decripción : " & Mid(tRs.Fields("Descripcion_PIEZAS"), 1, 30) & "-"
        If Len(tRs.Fields("Descripcion_PIEZAS")) > 30 Then
            Printer.Print "Decripcion : " & Mid(tRs.Fields("Descripcion_PIEZAS"), 31, 70) & "-"
            If Len(tRs.Fields("Descripcion_PIEZAS")) > 70 Then
                Printer.Print "Decripcion : " & Mid(tRs.Fields("Descripcion_PIEZAS"), 71, 111) & "-"
            End If
        End If
        Printer.Print "Comentarios : " & Mid(tRs.Fields("COMENTARIOS_TECNICOS"), 1, 30)
        If Len(tRs.Fields("COMENTARIOS_TECNICOS")) > 30 Then
            Printer.Print "Decripcion : " & Mid(tRs.Fields("COMENTARIOS_TECNICOS"), 31, 70) & "-"
            If Len(tRs.Fields("COMENTARIOS_TECNICOS")) > 70 Then
                Printer.Print "Decripcion : " & Mid(tRs.Fields("COMENTARIOS_TECNICOS"), 71, 111) & "-"
            End If
        End If
        Printer.Print "Articulo : " & Mid(tRs.Fields("TIPO_ARTICULO"), 1, 30)
        If Len(tRs.Fields("TIPO_ARTICULO")) > 30 Then
            Printer.Print "Decripcion : " & Mid(tRs.Fields("TIPO_ARTICULO"), 31, 70) & "-"
            If Len(tRs.Fields("TIPO_ARTICULO")) > 70 Then
                Printer.Print "Decripcion : " & Mid(tRs.Fields("TIPO_ARTICULO"), 71, 111) & "-"
            End If
        End If
        Printer.Print ""
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "               GRACIAS POR SU COMPRA"
        Printer.Print "           PRODUCTO 100% GARANTIZADO"
        Printer.Print "LA GARANTÍA SERÁ VÁLIDA HASTA 15 DIAS"
        Printer.Print "     DESPUES DE HABER EFECTUADO SU "
        Printer.Print "                                COMPRA"
        Printer.Print "SI NO RECOGE SU EQUIPO DESPUES DE 15 "
        Printer.Print "DIAS DE FINALIZADO EL SERVICIO, LA "
        Printer.Print "EMPRESA NO SE HACE RESPONSABLE POR "
        Printer.Print "                               EXTRABIO"
        Printer.Print "                APLICA RESTRICCIONES"
        Printer.Print ""
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.EndDoc
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ImpRecep()
    Dim oDoc  As cPDF
    Dim Posi As Integer
    Dim sBuscar As String
    Dim tRs  As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Set oDoc = New cPDF
    Posi = 185
    Image4.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Not oDoc.PDFCreate(App.Path & "\CartVac.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Courier_Bold, MacRomanEncoding
' Encabezado del reporte
    oDoc.LoadImage Image4, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 80, 40, 43, 161, "Logo"
    sBuscar = "SELECT * FROM REV_COMPRA_ALMACEN1 WHERE GRUPO = " & Text1.Text & " ORDER BY ID_REVISION DESC"
    Set tRs = cnn.Execute(sBuscar)
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F1", 7, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F1", 7, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F1", 7, hCenter
    oDoc.WTextBox 90, 200, 20, 250, "RECEPCION DE PRODUCTOS PARA REVISION", "F2", 10, hCenter
    oDoc.WTextBox 60, 400, 20, 250, "FECHA DE RECEPCION", "F2", 8, hCenter
    oDoc.WTextBox 70, 510, 20, 250, Format(Date, "dd/mm/yyyy"), "F2", 8, hLeft
' Encabezado de pagina
    sBuscar = "SELECT * FROM PROVEEDOR_ALMACEN1 WHERE ID_PROVEEDOR = " & tRs.Fields("ID_PROVEEDOR")
    Set tRs1 = cnn.Execute(sBuscar)
    oDoc.WTextBox 100, 10, 30, 400, "PROVEEDOR", "F2", 10, hLeft
    oDoc.WTextBox 110, 10, 30, 400, tRs1.Fields("NOMBRE"), "F2", 10, hLeft
    oDoc.WTextBox 130, 10, 30, 400, "TELEFONO", "F2", 10, hLeft
    oDoc.WTextBox 140, 10, 30, 400, tRs1.Fields("TELEFONO"), "F2", 10, hLeft
    oDoc.WTextBox 100, 450, 50, 400, "FOLIO: " & tRs.Fields("GRUPO"), "F2", 10, hLeft
    oDoc.WTextBox 130, 400, 40, 400, "SUCURSAL", "F2", 10, hLeft
    oDoc.WTextBox 140, 400, 40, 400, tRs.Fields("SUCURSAL"), "F2", 10, hLeft
' Cuerpo del reporte
    oDoc.WTextBox 170, 5, 40, 145, "CLAVE", "F2", 10, hLeft
    oDoc.WTextBox 170, 100, 40, 300, "DESCRIPCION", "F2", 10, hLeft
    oDoc.WTextBox 170, 350, 40, 70, "CANTIDAD", "F2", 10, hRight
    oDoc.WTextBox 170, 370, 40, 100, "P. UNIT.", "F2", 10, hRight
    oDoc.WTextBox 170, 470, 40, 80, "SUBTOTAL", "F2", 10, hRight
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 185
    oDoc.WLineTo 580, 185
    Do While Not tRs.EOF
        sBuscar = "SELECT Descripcion FROM ALMACEN1 WHERE ID_PRODUCTO = '" & Trim(tRs.Fields("ID_PRODUCTO")) & "'"
        Set tRs2 = cnn.Execute(sBuscar)
        If (tRs2.EOF And tRs2.BOF) Then
            sBuscar = "SELECT Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO = '" & Trim(tRs.Fields("ID_PRODUCTO")) & "'"
            Set tRs2 = cnn.Execute(sBuscar)
            If (tRs2.EOF And tRs2.BOF) Then
                sBuscar = "SELECT Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO = '" & Trim(tRs.Fields("ID_PRODUCTO")) & "'"
                Set tRs2 = cnn.Execute(sBuscar)
            End If
        End If
        oDoc.WTextBox Posi, 5, 40, 145, tRs.Fields("ID_PRODUCTO"), "F1", 9, hLeft
        oDoc.WTextBox Posi, 100, 9, 300, tRs2.Fields("Descripcion"), "F1", 9, hLeft
        oDoc.WTextBox Posi, 350, 40, 70, tRs.Fields("CANTIDAD"), "F1", 9, hRight
        oDoc.WTextBox Posi, 370, 40, 100, tRs.Fields("PRECIO_COMPRA"), "F1", 9, hRight
        oDoc.WTextBox Posi, 470, 40, 80, CDbl(tRs.Fields("CANTIDAD")) * CDbl(tRs.Fields("PRECIO_COMPRA")), "F1", 9, hRight
        tRs.MoveNext
        Posi = Posi + 10
    Loop
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, Posi + 20
    oDoc.WLineTo 580, Posi + 20
    oDoc.WTextBox 690, 5, 60, 580, "EL SR(A). " & tRs1.Fields("NOMBRE") & " CON DOMICILIO " & tRs1.Fields("DIRECCION") & " QUIEN SE IDENTIFICA CON " & tRs1.Fields("IDENTIFICACION") & " CON FOTOGRAFIA NO. " & tRs1.Fields("NUMERO_ID") & " VENDE A " & VarMen.TxtEmp(0).Text & " LOS ARTICULOS ARRIBA MENCIONADOS, MANIFESTANDO BAJO PROTESTA DE DECIR LA VERDAD QUE LA MERCANCIA ANTES DESCRITA SON DE SU EXCLUSIVA CONCESION Y DOMINIO, QUE LAS MISMAS FUERON ADQUIRIDAS CON SU PROPIO PECULIO DE MANERA LEGAL SIN PROCEDER DE HECHO ILICITO, POR LO QUE LIBERO DE CUALQUIER RESPONSABIIDAD PENAL O CIVIL AL COMPRADOR.", "F1", 10, hLeft
    oDoc.WTextBox 750, 5, 10, 580, "CLIENTE " & tRs1.Fields("NOMBRE") & " " & VarMen.Text4(3).Text & ", " & VarMen.Text4(4).Text & " A " & Format(Date, "dd/mm/yyyy"), "F1", 10, hLeft
    oDoc.WTextBox 770, 5, 10, 580, "ESTE DOCUMENTO AMPARA LA RECEPCION  DE LO CARTUCHOS ANTERIORMENTE DESCRITOS, Y NO GARANTIZA LA COMPRA TOTAL O PARCIAL DE LOS MISMOS. EL PRECIO MOSTRADO ES UN ESTIMADO DEL PRECIO MAXIMO DE COMPRA, Y ESTA SUJETO A CAMBIOS POR MOTIVOS DE FUNCIONAMIENTO O ESTADO DE LOS CARTUCHOS.", "F1", 10, hLeft
     oDoc.WTextBox 790, 5, 10, 580, "SE LE RECUERDA A LOS PROVEEDORES QUE EL HORARIO DE RECEPCION DE CARTUCHOS VACIOS ES DE LUNES A VIERNES DE 8:00 A 12:00 Y DE 4:00 A 5:30.", "F1", 10, hLeft
    oDoc.LineStroke
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub ImpEntProvVarios()
On Error GoTo ManejaError
    Dim Total As Double
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Total = 0
    CommonDialog1.Flags = 64
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    sBuscar = "SELECT REV_COMPRA_ALMACEN1.FECHA, REV_COMPRA_ALMACEN1.GRUPO, REV_COMPRA_ALMACEN1.ID_PRODUCTO, REV_COMPRA_ALMACEN1.CANTIDAD_APROVADA, REV_COMPRA_ALMACEN1.PRECIO_COMPRA, PROVEEDOR_ALMACEN1.NOMBRE  FROM REV_COMPRA_ALMACEN1, PROVEEDOR_ALMACEN1 WHERE PROVEEDOR_ALMACEN1.ID_PROVEEDOR = REV_COMPRA_ALMACEN1.ID_PROVEEDOR AND GRUPO = " & Text1.Text & " AND APROVADO = 'A'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print ""
        Printer.Print ""
        Printer.Print "             FECHA : " & tRs.Fields("FECHA")
        Printer.Print "             SUCURSAL : BODEGA"
        Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE REGISTRO DE ENTRADA DE COMPRAS DE PRODUCTOS")) / 2
        Printer.Print "COMPROBANTE REGISTRO DE ENTRADA DE COMPRAS DE PRODUCTOS"
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Dim NRegistros As Integer
        NRegistros = ListView2.ListItems.Count
        Dim Con As Integer
        Dim POSY As Integer
        POSY = 3800
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Folio"
        Printer.CurrentY = POSY
        Printer.CurrentX = 1000
        Printer.Print "Proveedor"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3500
        Printer.Print "Articulo"
        Printer.CurrentY = POSY
        Printer.CurrentX = 5500
        Printer.Print "Cantidad Registrada"
        Printer.CurrentY = POSY
        Printer.CurrentX = 7500
        Printer.Print "Precio"
        POSY = POSY + 200
        Do While Not tRs.EOF
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print tRs.Fields("GRUPO")
            Printer.CurrentY = POSY
            Printer.CurrentX = 1000
            Printer.Print tRs.Fields("NOMBRE")
            Printer.CurrentY = POSY
            Printer.CurrentX = 3500
            Printer.Print tRs.Fields("ID_PRODUCTO")
            Total = Total + (CDbl(Replace(tRs.Fields("CANTIDAD_APROVADA"), ",", "")) * CDbl(Replace(tRs.Fields("PRECIO_COMPRA"), ",", "")))
            Printer.CurrentY = POSY
            Printer.CurrentX = 5500
            Printer.Print Format(tRs.Fields("CANTIDAD_APROVADA"), "###,###,##0.00")
            Printer.CurrentY = POSY
            Printer.CurrentX = 7500
            Printer.Print tRs.Fields("PRECIO_COMPRA")
            If POSY >= 14200 Then
                Printer.NewPage
                Printer.Print ""
                Printer.Print ""
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                Printer.Print VarMen.Text5(0).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                Printer.Print ""
                Printer.Print ""
                Printer.Print "             FECHA : " & tRs.Fields("FECHA")
                Printer.Print "             SUCURSAL : " & VarMen.Text4(0).Text
                Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
                Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE REGISTRO DE ENTRADA DE COMPRAS DE PRODUCTOS")) / 2
                Printer.Print "COMPROBANTE REGISTRO DE ENTRADA DE COMPRAS DE PRODUCTOS"
                Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                POSY = 3800
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print "Folio"
                Printer.CurrentY = POSY
                Printer.CurrentX = 1000
                Printer.Print "Proveedor"
                Printer.CurrentY = POSY
                Printer.CurrentX = 3500
                Printer.Print "Articulo"
                Printer.CurrentY = POSY
                Printer.CurrentX = 5500
                Printer.Print "Cantidad Registrada"
                Printer.CurrentY = POSY
                Printer.CurrentX = 7500
                Printer.Print "Precio"
            End If
            tRs.MoveNext
        Loop
        Printer.Print ""
        Printer.Print "             Total = " & Total
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.EndDoc
        CommonDialog1.Copies = 1
    Else
        MsgBox "LA ENTRADA NO FUE ENCONTRADA!", vbExclamation, "SACC"
    End If
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub RECIBIDO()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim loca As Integer
    Dim sBuscar As String
    Dim tRs1 As ADODB.Recordset
    Dim PosVer As Integer
    Dim Posabo As Integer
    Dim tRs  As ADODB.Recordset
    Dim tRs2  As ADODB.Recordset
    Set oDoc = New cPDF
    Dim sumdeuda As Double
    Dim sumIndi As Double
    Dim sNombre As String
    Dim sumabonos As Double
    Dim sumIndiabonos As Double
    Dim AcumDeudas As String
    Dim NoRe As Integer
    Dim ConPag As Integer
    ConPag = 1
    sBuscar = "SELECT ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA, ORDEN_RAPIDA_DETALLE.ID_PRODUCTO, ORDEN_RAPIDA_DETALLE.CAN_RECIBIDA, ORDEN_RAPIDA.COMENTARIO FROM ORDEN_RAPIDA_DETALLE, ORDEN_RAPIDA WHERE ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA = ORDEN_RAPIDA.ID_ORDEN_RAPIDA AND ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA = '" & Text1.Text & "' AND ORDEN_RAPIDA_DETALLE.SURTIDO = 'S'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        If Not oDoc.PDFCreate(App.Path & "\entraordenrapida.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
    ' Encabezado del reporte
        'Image2.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        'oDoc.LoadImage Image2, "Logo", False, False
        oDoc.NewPage A4_Vertical
       ' oDoc.WImage 80, 40, 43, 161, "Logo"
        oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
        oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
        oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
        oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
        oDoc.WTextBox 90, 200, 20, 250, "Comprobante de Recibido de Orden Rapida", "F2", 10, hCenter
        oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
        oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
        oDoc.WTextBox 100, 20, 30, 100, "ORDEN RAPIDA ", "F2", 9, hLeft
        oDoc.WTextBox 100, 100, 30, 80, "PRODUCTO ", "F2", 9, hLeft
        oDoc.WTextBox 100, 180, 40, 200, "CANTIDAD ", "F2", 9, hLeft
        oDoc.WTextBox 100, 240, 40, 300, "COMENTARIO ", "F2", 9, hLeft
        Posi = 110 + 10
        oDoc.WTextBox Posi, 240, 150, 300, tRs.Fields("COMENTARIO"), "F2", 9, hLeft
        Do While Not tRs.EOF
            oDoc.WTextBox Posi, 20, 30, 40, tRs.Fields("ID_ORDEN_RAPIDA"), "F2", 9, hLeft
            oDoc.WTextBox Posi, 100, 30, 150, tRs.Fields("ID_PRODUCTO"), "F2", 9, hLeft
            oDoc.WTextBox Posi, 180, 40, 200, tRs.Fields("CAN_RECIBIDA"), "F2", 9, hLeft
            'oDoc.WTextBox Posi, 240, 150, 300, tRs.Fields("COMENTARIO"), "F2", 9, hLeft
            tRs.MoveNext
            Posi = Posi + 10
        Loop
    ' Encabezado de pagina
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 100
        oDoc.WLineTo 580, 100
        oDoc.LineStroke
        Posi = Posi + 30
        oDoc.WTextBox Posi, 50, 40, 200, " RECIBIO", "F2", 9, hLeft
        oDoc.WTextBox Posi, 400, 40, 300, " ENTREGO", "F2", 9, hLeft
        If Posi >= 760 Then
            oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
            ConPag = ConPag + 1
            oDoc.NewPage A4_Vertical
            ' Encabezado del reporte
            Posi = 110
            oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
            oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
            oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
            oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
            oDoc.WTextBox 90, 200, 20, 250, "Comprobante de Recibido e Orden Rapida", "F2", 10, hCenter
            oDoc.WTextBox 80, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
            oDoc.WTextBox 90, 380, 20, 250, Date, "F3", 8, hCenter
            ' Encabezado de pagina
            oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 100
            oDoc.WLineTo 580, 100
            oDoc.LineStroke
        End If
        Cont = Cont + 1
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "LA ORDEN NO TIENE ENTRADAS REGISTRADAS", vbExclamation, "SACC"
    End If
End Sub
Private Sub ImrPolizaCheque()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim sBusca As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim totprod As Double
    Dim sqlQuery As String
    Dim sTipo As String
    Set oDoc = New cPDF
    Image4.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Option8.Value Then
        sTipo = "NACIONAL"
    End If
    If Option9.Value Then
        sTipo = "INTERNACIONAL"
    End If
    If Option10.Value Then
        sTipo = "INDIRECTA"
    End If
    If Option17.Value Then
        sTipo = "RAPIDA"
    End If
    sqlQuery = "SELECT TOP 1 ID_CHEQUE FROM VSCHEQUES2 WHERE (NUM_ORDEN LIKE '% " & Text1.Text & ",%' OR NUM_ORDEN LIKE '" & Text1.Text & ",%') AND TIPO_ORDEN = '" & sTipo & "' ORDER BY ID_CHEQUE DESC" '
    Set tRs = cnn.Execute(sqlQuery)
    If (tRs.EOF And tRs.BOF) Then
        MsgBox "NO SE ENCONTRARON CHEQUES PARA ESTA ORDEN"
        Exit Sub
    End If
    If Not oDoc.PDFCreate(App.Path & "\Cheque.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    Image4.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    oDoc.LoadImage Image4, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 50, 40, 38, 161, "Logo"
    sBusca = tRs.Fields("ID_CHEQUE")
    sqlQuery = "SELECT * FROM VSCHEQUES2 WHERE ID_CHEQUE='" & tRs.Fields("ID_CHEQUE") & "' "
    Set tRs2 = cnn.Execute(sqlQuery)
    'cuadros encabezado
    'Posi = 50
    oDoc.WTextBox 25, 280, 15, 300, "POLIZA DE TRANSFERENCIA", "F2", 20, hLeft
    ''''primer  cuadro
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 50
    oDoc.WLineTo 580, 50
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 230
    oDoc.WLineTo 580, 230
    oDoc.LineStroke
    'Posi = Posi + 6
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 50
    oDoc.WLineTo 10, 230
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 580, 50
    oDoc.WLineTo 580, 230
    oDoc.LineStroke
    '''final del primer
    ''''segundo  cuadro
    oDoc.WTextBox 240, 30, 20, 200, "CONCEPTO TRANSFERENCIA :", "F3", 8, hLeft
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 235
    oDoc.WLineTo 400, 235
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 250
    oDoc.WLineTo 400, 250
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 320
    oDoc.WLineTo 400, 320
    oDoc.LineStroke
    'Posi = Posi + 6
    'Posi = Posi + 6
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 235
    oDoc.WLineTo 10, 320
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 400, 235
    oDoc.WLineTo 400, 320
    oDoc.LineStroke
    '''final del segundo
    ''''  tercero
    oDoc.WTextBox 240, 415, 20, 200, "FIRMA TRANSFERENCIA RECIBIDO  :", "F3", 8, hLeft
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 410, 235
    oDoc.WLineTo 580, 235
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 410, 250
    oDoc.WLineTo 580, 250
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 410, 320
    oDoc.WLineTo 580, 320
    oDoc.LineStroke
    'Posi = Posi + 6
    'Posi = Posi + 6
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 410, 235
    oDoc.WLineTo 410, 320
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 580, 235
    oDoc.WLineTo 580, 320
    oDoc.LineStroke
    '''final del tercero
    ''''  cuarto
    oDoc.WTextBox 328, 30, 20, 50, "CUENTA:", "F3", 8, hLeft
    oDoc.WTextBox 328, 90, 20, 70, "SUB-CUENTA:", "F3", 8, hLeft
    oDoc.WTextBox 328, 210, 20, 50, "NOMBRE:", "F3", 8, hLeft
    oDoc.WTextBox 328, 350, 20, 50, "PARCIAL:", "F3", 8, hLeft
    oDoc.WTextBox 328, 450, 20, 50, "DEBER:", "F3", 8, hLeft
    oDoc.WTextBox 328, 540, 20, 50, "HABER:", "F3", 8, hLeft
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 325
    oDoc.WLineTo 580, 325
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 340
    oDoc.WLineTo 580, 340
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 400
    oDoc.WLineTo 580, 400
    oDoc.LineStroke
    'Posi = Posi + 6
    'Posi = Posi + 6
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 325
    oDoc.WLineTo 10, 400
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 580, 325
    oDoc.WLineTo 580, 400
    oDoc.LineStroke
    '''lineas de
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 80, 325
    oDoc.WLineTo 80, 400
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 160, 325
    oDoc.WLineTo 160, 400
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 330, 325
    oDoc.WLineTo 330, 400
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 420, 325
    oDoc.WLineTo 420, 400
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 500, 325
    oDoc.WLineTo 500, 400
    oDoc.LineStroke
    '''final del cuarto
    ''''  quinto
    oDoc.WTextBox 407, 30, 20, 70, "HECHO POR:", "F3", 8, hLeft
    oDoc.WTextBox 407, 90, 20, 60, "REVISADO:", "F3", 8, hLeft
    oDoc.WTextBox 407, 210, 20, 70, "AUTORIZADO:", "F3", 8, hLeft
    oDoc.WTextBox 440, 200, 20, 150, VarMen.TxtEmp(11).Text & ":", "F3", 8, hLeft
    oDoc.WTextBox 407, 350, 20, 50, "AUXILIARES:", "F3", 8, hLeft
    oDoc.WTextBox 407, 450, 20, 50, "DIARIO:", "F3", 8, hLeft
    oDoc.WTextBox 407, 540, 20, 50, "HABER:", "F3", 8, hLeft
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 405
    oDoc.WLineTo 580, 405
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 420
    oDoc.WLineTo 580, 420
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 460
    oDoc.WLineTo 580, 460
    oDoc.LineStroke
    'Posi = Posi + 6
    'Posi = Posi + 6
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 405
    oDoc.WLineTo 10, 460
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 580, 405
    oDoc.WLineTo 580, 460
    oDoc.LineStroke
    '''lineas de
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 80, 405
    oDoc.WLineTo 80, 460
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 160, 405
    oDoc.WLineTo 160, 460
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 330, 405
    oDoc.WLineTo 330, 460
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 420, 405
    oDoc.WLineTo 420, 460
    oDoc.LineStroke
    oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 500, 405
    oDoc.WLineTo 500, 460
    oDoc.LineStroke
    '''final del quinto
    Posi = 230
    If Not (tRs2.EOF And tRs2.BOF) Then
        Do While Not (tRs2.EOF)
            oDoc.WTextBox 70, 20, 20, 300, tRs2.Fields("NOMBRE"), "F3", 8, hLeft
            oDoc.WTextBox 55, 400, 20, 50, tRs2.Fields("FECHA"), "F3", 8, hLeft
            oDoc.WTextBox 70, 420, 20, 100, "$ " & Format(tRs2.Fields("TOTAL"), "###,###,##0.00"), "F3", 8, hLeft
            oDoc.WTextBox 88, 20, 20, 300, tRs2.Fields("TOTAL_LETRA"), "F3", 8, hLeft
            oDoc.WTextBox 200, 190, 20, 120, "TRANSFERENCIA: " & tRs2.Fields("NUM_CHEQUE"), "F3", 8, hLeft
            oDoc.WTextBox 200, 320, 20, 50, tRs2.Fields("BANCO"), "F3", 8, hLeft
            oDoc.WTextBox 260, 20, 20, 80, "PAGO DE O.C No", "F3", 8, hLeft
            sBuscar = "SELECT * FROM CHEQUES WHERE (NUM_ORDEN LIKE '%, " & Text1.Text & ", %' OR NUM_ORDEN LIKE '" & Text1.Text & ", %') AND (TIPO_ORDEN = '" & sTipo & "')"
            'sBusca = "SELECT NUM_ORDEN FROM VSCHEQUES2 WHERE NUM_CHEQUE = '" & tRs2.Fields("NUM_CHEQUE") & "' AND BANCO = '" & tRs2.Fields("BANCO") & "'"
            Set tRs3 = cnn.Execute(sBuscar)
            Dim NumerosOrdenes As String
            If Not (tRs3.EOF And tRs3.BOF) Then
                oDoc.WTextBox 260, 100, 20, 80, tRs3.Fields("NUM_ORDEN"), "F3", 8, hLeft
                Do While Not tRs3.EOF
                    NumerosOrdenes = Me.Text1.Text 'NumerosOrdenes & tRs3.Fields("NUM_ORDEN") & ", "
                    tRs3.MoveNext
                Loop
            End If
            'If Mid(NumerosOrdenes, Len(NumerosOrdenes) - 1, 1) = " " Then
            '    oDoc.WTextBox 260, 100, 20, 270, Mid(NumerosOrdenes, 1, Len(NumerosOrdenes) - 2), "F3", 8, hLeft
            'Else
            '    oDoc.WTextBox 260, 100, 20, 270, NumerosOrdenes, "F3", 8, hLeft
            'End If
            oDoc.WTextBox 270, 120, 20, 300, tRs2.Fields("NOMBRE"), "F3", 8, hLeft
            oDoc.WTextBox 300, 250, 20, 200, "TRANSFERENCIA:" & tRs2.Fields("NUM_CHEQUE"), "F3", 8, hLeft
            'oDoc.WTextBox 300, 300, 20, 80, tRs2.Fields("NUM_CHEQUE"), "F3", 8, hLeft
            tRs2.MoveNext
        Loop
    End If
    oDoc.WTextBox 470, 30, 20, 250, "REIMPRESION DEL SISTEMA, NO SE MUESTRAN LOS PAGOS REALIZADOS CON ANTERIORIDAD.", "F2", 8, hCenter
    Posi = Posi + 6
    'cierre del reporte
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub ImrPolizaChequeVarios()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim sBusca As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim totprod As Double
    Dim sqlQuery As String
    Dim sTipo As String
    Dim Con As Integer
    Set oDoc = New cPDF
    Image4.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Option8.Value Then
        sTipo = "NACIONAL"
    End If
    If Option9.Value Then
        sTipo = "INTERNACIONAL"
    End If
    If Option10.Value Then
        sTipo = "INDIRECTA"
    End If
    If Option17.Value Then
        sTipo = "RAPIDA"
    End If
    Con = 1
    sqlQuery = "SELECT ID_CHEQUE FROM VSCHEQUES2 WHERE (NUM_ORDEN LIKE '% " & Text1.Text & ",%' OR NUM_ORDEN LIKE '" & Text1.Text & ",%') AND TIPO_ORDEN = '" & sTipo & "' ORDER BY ID_CHEQUE DESC" '
    Set tRs = cnn.Execute(sqlQuery)
    If (tRs.EOF And tRs.BOF) Then
        MsgBox "NO SE ENCONTRARON CHEQUES PARA ESTA ORDEN"
        Exit Sub
    Else
        Do While Not tRs.EOF
            If Con = 1 Then
                If Not oDoc.PDFCreate(App.Path & "\Cheque" & ".pdf") Then
                    Exit Sub
                End If
                oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
                oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
                Image4.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
            End If
            oDoc.LoadImage Image4, "Logo", False, False
            oDoc.NewPage A4_Vertical
            oDoc.WImage 50, 40, 38, 161, "Logo"
            sBusca = tRs.Fields("ID_CHEQUE")
            sqlQuery = "SELECT * FROM VSCHEQUES2 WHERE ID_CHEQUE = '" & tRs.Fields("ID_CHEQUE") & "' "
            Set tRs2 = cnn.Execute(sqlQuery)
            'cuadros encabezado
            'Posi = 50
            oDoc.WTextBox 25, 280, 15, 300, "POLIZA DE TRANSFERENCIA", "F2", 20, hLeft
            ''''primer  cuadro
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 50
            oDoc.WLineTo 580, 50
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 230
            oDoc.WLineTo 580, 230
            oDoc.LineStroke
            'Posi = Posi + 6
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 50
            oDoc.WLineTo 10, 230
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 580, 50
            oDoc.WLineTo 580, 230
            oDoc.LineStroke
            '''final del primer
            ''''segundo  cuadro
            oDoc.WTextBox 240, 30, 20, 200, "CONCEPTO TRANSFERENCIA :", "F3", 8, hLeft
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 235
            oDoc.WLineTo 400, 235
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 250
            oDoc.WLineTo 400, 250
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 320
            oDoc.WLineTo 400, 320
            oDoc.LineStroke
            'Posi = Posi + 6
            'Posi = Posi + 6
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 235
            oDoc.WLineTo 10, 320
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 400, 235
            oDoc.WLineTo 400, 320
            oDoc.LineStroke
            '''final del segundo
            ''''  tercero
            oDoc.WTextBox 240, 415, 20, 200, "FIRMA TRANSFERENCIA RECIBIDO  :", "F3", 8, hLeft
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 410, 235
            oDoc.WLineTo 580, 235
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 410, 250
            oDoc.WLineTo 580, 250
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 410, 320
            oDoc.WLineTo 580, 320
            oDoc.LineStroke
            'Posi = Posi + 6
            'Posi = Posi + 6
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 410, 235
            oDoc.WLineTo 410, 320
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 580, 235
            oDoc.WLineTo 580, 320
            oDoc.LineStroke
            '''final del tercero
            ''''  cuarto
            oDoc.WTextBox 328, 30, 20, 50, "CUENTA:", "F3", 8, hLeft
            oDoc.WTextBox 328, 90, 20, 70, "SUB-CUENTA:", "F3", 8, hLeft
            oDoc.WTextBox 328, 210, 20, 50, "NOMBRE:", "F3", 8, hLeft
            oDoc.WTextBox 328, 350, 20, 50, "PARCIAL:", "F3", 8, hLeft
            oDoc.WTextBox 328, 450, 20, 50, "DEBER:", "F3", 8, hLeft
            oDoc.WTextBox 328, 540, 20, 50, "HABER:", "F3", 8, hLeft
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 325
            oDoc.WLineTo 580, 325
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 340
            oDoc.WLineTo 580, 340
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 400
            oDoc.WLineTo 580, 400
            oDoc.LineStroke
            'Posi = Posi + 6
            'Posi = Posi + 6
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 325
            oDoc.WLineTo 10, 400
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 580, 325
            oDoc.WLineTo 580, 400
            oDoc.LineStroke
            '''lineas de
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 80, 325
            oDoc.WLineTo 80, 400
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 160, 325
            oDoc.WLineTo 160, 400
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 330, 325
            oDoc.WLineTo 330, 400
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 420, 325
            oDoc.WLineTo 420, 400
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 500, 325
            oDoc.WLineTo 500, 400
            oDoc.LineStroke
            '''final del cuarto
            ''''  quinto
            oDoc.WTextBox 407, 30, 20, 70, "HECHO POR:", "F3", 8, hLeft
            oDoc.WTextBox 407, 90, 20, 60, "REVISADO:", "F3", 8, hLeft
            oDoc.WTextBox 407, 210, 20, 70, "AUTORIZADO:", "F3", 8, hLeft
            oDoc.WTextBox 440, 200, 20, 150, VarMen.TxtEmp(11).Text & ":", "F3", 8, hLeft
            oDoc.WTextBox 407, 350, 20, 50, "AUXILIARES:", "F3", 8, hLeft
            oDoc.WTextBox 407, 450, 20, 50, "DIARIO:", "F3", 8, hLeft
            oDoc.WTextBox 407, 540, 20, 50, "HABER:", "F3", 8, hLeft
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 405
            oDoc.WLineTo 580, 405
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 420
            oDoc.WLineTo 580, 420
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 460
            oDoc.WLineTo 580, 460
            oDoc.LineStroke
            'Posi = Posi + 6
            'Posi = Posi + 6
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 405
            oDoc.WLineTo 10, 460
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 580, 405
            oDoc.WLineTo 580, 460
            oDoc.LineStroke
            '''lineas de
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 80, 405
            oDoc.WLineTo 80, 460
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 160, 405
            oDoc.WLineTo 160, 460
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 330, 405
            oDoc.WLineTo 330, 460
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 420, 405
            oDoc.WLineTo 420, 460
            oDoc.LineStroke
            oDoc.SetLineFormat 0.8, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 500, 405
            oDoc.WLineTo 500, 460
            oDoc.LineStroke
            '''final del quinto
            Posi = 230
            If Not (tRs2.EOF And tRs2.BOF) Then
                Do While Not (tRs2.EOF)
                    oDoc.WTextBox 70, 20, 20, 300, tRs2.Fields("NOMBRE"), "F3", 8, hLeft
                    oDoc.WTextBox 55, 400, 20, 50, tRs2.Fields("FECHA"), "F3", 8, hLeft
                    oDoc.WTextBox 70, 420, 20, 100, "$ " & Format(tRs2.Fields("TOTAL"), "###,###,##0.00"), "F3", 8, hLeft
                    oDoc.WTextBox 88, 20, 20, 300, tRs2.Fields("TOTAL_LETRA"), "F3", 8, hLeft
                    oDoc.WTextBox 200, 190, 20, 120, "TRANSFERENCIA: " & tRs2.Fields("NUM_CHEQUE"), "F3", 8, hLeft
                    oDoc.WTextBox 200, 320, 20, 50, tRs2.Fields("BANCO"), "F3", 8, hLeft
                    oDoc.WTextBox 260, 20, 20, 80, "PAGO DE O.C No", "F3", 8, hLeft
                    sBuscar = "SELECT * FROM CHEQUES WHERE (NUM_ORDEN LIKE '%, " & Text1.Text & ", %' OR NUM_ORDEN LIKE '" & Text1.Text & ", %') AND (TIPO_ORDEN = '" & sTipo & "')"
                    'sBusca = "SELECT NUM_ORDEN FROM VSCHEQUES2 WHERE NUM_CHEQUE = '" & tRs2.Fields("NUM_CHEQUE") & "' AND BANCO = '" & tRs2.Fields("BANCO") & "'"
                    Set tRs3 = cnn.Execute(sBuscar)
                    Dim NumerosOrdenes As String
                    If Not (tRs3.EOF And tRs3.BOF) Then
                        oDoc.WTextBox 260, 100, 20, 80, tRs3.Fields("NUM_ORDEN"), "F3", 8, hLeft
                        Do While Not tRs3.EOF
                            NumerosOrdenes = Me.Text1.Text 'NumerosOrdenes & tRs3.Fields("NUM_ORDEN") & ", "
                            tRs3.MoveNext
                        Loop
                    End If
                    'If Mid(NumerosOrdenes, Len(NumerosOrdenes) - 1, 1) = " " Then
                    '    oDoc.WTextBox 260, 100, 20, 270, Mid(NumerosOrdenes, 1, Len(NumerosOrdenes) - 2), "F3", 8, hLeft
                    'Else
                    '    oDoc.WTextBox 260, 100, 20, 270, NumerosOrdenes, "F3", 8, hLeft
                    'End If
                    oDoc.WTextBox 270, 120, 20, 300, tRs2.Fields("NOMBRE"), "F3", 8, hLeft
                    oDoc.WTextBox 300, 250, 20, 200, "TRANSFERENCIA:" & tRs2.Fields("NUM_CHEQUE"), "F3", 8, hLeft
                    'oDoc.WTextBox 300, 300, 20, 80, tRs2.Fields("NUM_CHEQUE"), "F3", 8, hLeft
                    oDoc.WTextBox 470, 30, 20, 250, "REIMPRESION DEL SISTEMA, NO SE MUESTRAN LOS PAGOS REALIZADOS CON ANTERIORIDAD.", "F2", 8, hCenter
                    Posi = Posi + 6
                    tRs2.MoveNext
                    'If Not (tRs.EOF) Then
                    '    oDoc.NewPage A4_Vertical
                    'End If
                Loop
            End If
            'cierre del reporte
            Con = Con + 1
            tRs.MoveNext
        Loop
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub ImpRemision()
    Dim oDoc  As cPDF
    Dim Cont As Integer
    Dim Posi As Integer
    Dim sBuscar As String
    Dim tRs1 As ADODB.Recordset
    Dim PosVer As Integer
    Dim tRs  As ADODB.Recordset
    Set oDoc = New cPDF
    Image4.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Not oDoc.PDFCreate(App.Path & "\RepCuentasPagadas.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
    ' Encabezado del reporte
    oDoc.LoadImage Image4, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 70, 40, 43, 161, "Logo"
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
    sBuscar = "SELECT ID_VENTA, NOMBRE, SUBTOTAL, IVA, TOTAL FROM VENTAS WHERE ID_VENTA IN (" & Text1.Text & ")"
    Set tRs1 = cnn.Execute(sBuscar)
    oDoc.WTextBox 90, 10, 20, 570, tRs1.Fields("NOMBRE"), "F2", 10, hCenter
    oDoc.WTextBox 40, 500, 20, 70, "REMISION", "F3", 8, hCenter, , , 1
    oDoc.WTextBox 50, 500, 20, 70, tRs1.Fields("ID_VENTA"), "F2", 8, hCenter
    oDoc.WTextBox 70, 500, 20, 70, "Fecha", "F3", 8, hCenter
    oDoc.WTextBox 80, 500, 20, 70, Date, "F3", 8, hCenter
    ' Encabezado de pagina
    oDoc.WTextBox 120, 10, 10, 50, "CANTIDAD", "F2", 8, hCenter, , , 1
    oDoc.WTextBox 120, 60, 10, 70, "PRESENTACION", "F2", 8, hCenter, , , 1
    oDoc.WTextBox 120, 130, 10, 70, "CODIGO", "F2", 8, hCenter, , , 1
    oDoc.WTextBox 120, 200, 10, 280, "DESCRIPCION", "F2", 8, hCenter, , , 1
    oDoc.WTextBox 120, 480, 10, 50, "IMPORTE", "F2", 8, hCenter, , , 1
    oDoc.WTextBox 120, 530, 10, 60, "TOTAL", "F2", 8, hCenter, , , 1
    ' Cuerpo del reporte
    Posi = 130
    sBuscar = "SELECT VENTAS_DETALLE.ID_VENTA, VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.Descripcion, VENTAS_DETALLE.CANTIDAD, ALMACEN3.PRESENTACION , VENTAS_DETALLE.Precio_Venta FROM ALMACEN3 INNER JOIN VENTAS_DETALLE ON ALMACEN3.ID_PRODUCTO = VENTAS_DETALLE.ID_PRODUCTO WHERE VENTAS_DETALLE.ID_VENTA IN (" & Text1.Text & ")"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Posi = Posi + 10
            oDoc.WTextBox Posi, 10, 10, 580, "", "F2", 8, hRight, , , 1, vbBlue
            If Not IsNull(tRs.Fields("CANTIDAD")) Then oDoc.WTextBox Posi, 10, 10, 50, tRs.Fields("CANTIDAD"), "F2", 8, hRight
            If Not IsNull(tRs.Fields("PRESENTACION")) Then oDoc.WTextBox Posi, 60, 10, 70, tRs.Fields("PRESENTACION"), "F2", 8, hCenter
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then oDoc.WTextBox Posi, 130, 10, 70, tRs.Fields("ID_PRODUCTO"), "F3", 8, hRight
            If Not IsNull(tRs.Fields("Descripcion")) Then oDoc.WTextBox Posi, 210, 10, 280, tRs.Fields("Descripcion"), "F3", 8, hLeft
            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then oDoc.WTextBox Posi, 480, 10, 50, Format(tRs.Fields("PRECIO_VENTA"), "###,###,#0.00"), "F3", 8, hRight
            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then oDoc.WTextBox Posi, 530, 10, 50, Format(tRs.Fields("PRECIO_VENTA") * tRs.Fields("CANTIDAD"), "###,###,#0.00"), "F3", 8, hRight
            If Posi >= 760 Then
                oDoc.LoadImage Image4, "Logo", False, False
                oDoc.NewPage A4_Vertical
                oDoc.WImage 70, 40, 43, 161, "Logo"
                oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                oDoc.WTextBox 90, 10, 20, 570, tRs1.Fields("NOMBRE"), "F2", 10, hCenter
                oDoc.WTextBox 40, 500, 20, 70, "REMISION", "F3", 8, hCenter, , , 1
                oDoc.WTextBox 50, 500, 20, 70, tRs1.Fields("ID_VENTA"), "F2", 8, hCenter
                oDoc.WTextBox 70, 500, 20, 70, "Fecha", "F3", 8, hCenter
                oDoc.WTextBox 80, 500, 20, 70, Date, "F3", 8, hCenter
                ' Encabezado de pagina
                oDoc.WTextBox 120, 10, 10, 50, "CANTIDAD", "F2", 8, hCenter, , , 1
                oDoc.WTextBox 120, 60, 10, 70, "PRESENTACION", "F2", 8, hCenter, , , 1
                oDoc.WTextBox 120, 130, 10, 70, "CODIGO", "F2", 8, hCenter, , , 1
                oDoc.WTextBox 120, 200, 10, 280, "DESCRIPCION", "F2", 8, hCenter, , , 1
                oDoc.WTextBox 120, 480, 10, 50, "IMPORTE", "F2", 8, hCenter, , , 1
                oDoc.WTextBox 120, 530, 10, 60, "TOTAL", "F2", 8, hCenter, , , 1
                ' Cuerpo del reporte
                Posi = 130
            End If
            tRs.MoveNext
        Loop
        Posi = Posi + 20
        oDoc.WTextBox Posi, 470, 10, 50, "Subtotal", "F3", 8, hRight
        If Not IsNull(tRs1.Fields("SUBTOTAL")) Then oDoc.WTextBox Posi, 530, 10, 50, Format(tRs1.Fields("SUBTOTAL"), "###,###,#0.00"), "F3", 8, hRight
        oDoc.WTextBox Posi, 530, 10, 60, "", "F3", 8, hRight, , , 1
        Posi = Posi + 10
        oDoc.WTextBox Posi, 470, 10, 50, "IVA", "F3", 8, hRight
        If Not IsNull(tRs1.Fields("IVA")) Then oDoc.WTextBox Posi, 530, 10, 50, Format(tRs1.Fields("IVA"), "###,###,#0.00"), "F3", 8, hRight
        oDoc.WTextBox Posi, 530, 10, 60, "", "F3", 8, hRight, , , 1
        Posi = Posi + 10
        oDoc.WTextBox Posi, 470, 10, 50, "Total", "F3", 8, hRight
        If Not IsNull(tRs1.Fields("TOTAL")) Then oDoc.WTextBox Posi, 530, 10, 50, Format(tRs1.Fields("TOTAL"), "###,###,#0.00"), "F3", 8, hRight
        oDoc.WTextBox Posi, 530, 10, 60, "", "F3", 8, hRight, , , 1
        Cont = Cont + 1
        Posi = Posi + 50
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 380, Posi
        oDoc.WLineTo 500, Posi
        oDoc.LineStroke
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 10, 70, "Firma de recibido", "F3", 8, hCenter
        oDoc.LineStroke
    End If
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub FunRemision()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tRs5 As ADODB.Recordset
    Dim tRs6 As ADODB.Recordset
    Dim ConPag As Integer
    ConPag = 1
    Dim sBuscar As String
    sBuscar = "SELECT * FROM VENTAS WHERE ID_VENTA = " & Text1.Text & ""
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\Remision.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        Image4.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image4, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 20, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 205, 20, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "Remision : " & tRs1.Fields("ID_VENTA"), "F3", 8, hCenter
        oDoc.WTextBox 70, 340, 20, 250, "Fecha :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        
        
        'CAJA1
        sBuscar = "SELECT * FROM CLIENTE WHERE ID_CLIENTE = " & tRs1.Fields("ID_CLIENTE")
        Set tRs2 = cnn.Execute(sBuscar)
        oDoc.WTextBox 110, 20, 100, 400, "CLIENTE:", "F3", 8, hLeft
        oDoc.WTextBox 120, 20, 100, 400, "DOMICILIO", "F3", 8, hLeft
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 110, 20, 100, 400, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 120, 20, 100, 400, tRs2.Fields("DIRECCION") & "Col. " & tRs2.Fields("COLONIA"), "F3", 8, hCenter
        End If
        Posi = 150
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 20, 50, "CANTIDAD", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 55, 20, 90, "CLAVE", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 145, 20, 280, "DESCRIPCION", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 425, 20, 60, "PRESENTACION", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 485, 20, 50, "PRECIO UNITARIO", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 535, 20, 50, "TOTAL", "F2", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 20

        ' DETALLE
        sBuscar = "SELECT VENTAS_DETALLE.CANTIDAD, VENTAS_DETALLE.ID_PRODUCTO, VENTAS_DETALLE.DESCRIPCION, ALMACEN3.PRESENTACION, VENTAS_DETALLE.PRECIO_VENTA, VENTAS_DETALLE.PRECIO_VENTA * VENTAS_DETALLE.CANTIDAD AS TOTAL FROM ALMACEN3 INNER JOIN VENTAS_DETALLE ON ALMACEN3.ID_PRODUCTO = VENTAS_DETALLE.ID_PRODUCTO WHERE VENTAS_DETALLE.ID_VENTA = " & tRs1.Fields("ID_VENTA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 5, 15, 50, Format(tRs3.Fields("CANTIDAD"), "###,###,##0.00"), "F3", 7, hCenter, , , 1, vbBlack
                
                'oDoc.WTextBox Posi, 55, 15, 90, " " & tRs3.Fields("ID_PRODUCTO"), "F3", 7, hLeft, , , 1, vbBlack
                oDoc.WTextBox Posi, 55, 15, 30, " " & Mid(tRs3.Fields("ID_PRODUCTO"), 1, 3), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 85, 15, 30, " " & Mid(tRs3.Fields("ID_PRODUCTO"), 4, 3), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 115, 15, 30, " " & Mid(tRs3.Fields("ID_PRODUCTO"), 7, 3), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 145, 15, 280, " " & tRs3.Fields("DESCRIPCION"), "F3", 7, hLeft, , , 1, vbBlack
                'oDoc.WTextBox Posi, 425, 15, 60, tRs3.Fields("PRESENTACION"), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 425, 15, 60, Mid(tRs3.Fields("ID_PRODUCTO"), 11, 7), "F3", 7, hCenter, , , 1, vbBlack
                oDoc.WTextBox Posi, 485, 15, 50, Format(CDbl(tRs3.Fields("PRECIO_VENTA")), "###,###,##0.00") & " ", "F3", 7, hRight, , , 1, vbBlack
                oDoc.WTextBox Posi, 535, 15, 50, Format(CDbl(tRs3.Fields("TOTAL")), "###,###,##0.00") & " ", "F3", 7, hRight, , , 1, vbBlack
                Posi = Posi + 15
                tRs3.MoveNext
                If Posi >= 600 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    sBuscar = "SELECT * FROM VENTAS WHERE ID_VENTA = " & tRs1.Fields("ID_VENTA")
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
                        oDoc.NewPage A4_Vertical
                        oDoc.WImage 70, 40, 43, 161, "Logo"
                        oDoc.WTextBox 40, 205, 20, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
                        oDoc.WTextBox 60, 205, 20, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
                
                        oDoc.WTextBox 60, 340, 20, 250, "Remision : " & tRs1.Fields("ID_VENTA"), "F3", 8, hCenter
                        oDoc.WTextBox 70, 340, 20, 250, "Fecha :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
                        Posi = Posi + 15
                        oDoc.WTextBox 110, 20, 100, 400, "CLIENTE:", "F3", 8, hLeft
                        oDoc.WTextBox 120, 20, 100, 400, "DOMICILIO", "F3", 8, hLeft
                        If Not (tRs2.EOF And tRs2.BOF) Then
                            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 110, 20, 100, 400, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
                            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 120, 20, 100, 400, tRs2.Fields("DIRECCION") & "Col. " & tRs2.Fields("COLONIA"), "F3", 8, hCenter
                        End If
                        Posi = 210
                    End If
                End If
            Loop
        End If
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 600
        oDoc.WLineTo 580, 600
        oDoc.LineStroke
        Posi = Posi + 6
        ' TEXTO ABAJO
        'oDoc.WTextBox 620, 20, 100, 275, "Envíe por favor los artículos a la direccion de envío, si usted tiene alguna duda o pregunta sobre esta orden de compra, por favor contacte a  Dept Compras, al Tel. " & VarMen.TxtEmp(2).Text, "F3", 8, hLeft, , , 0, vbBlack
        oDoc.WTextBox 620, 400, 20, 70, "SUBTOTAL:", "F2", 8, hRight
        oDoc.WTextBox 640, 400, 20, 70, "Impuesto 1:", "F2", 8, hRight
        oDoc.WTextBox 660, 400, 20, 70, "Impuesto 2:", "F2", 8, hRight
        oDoc.WTextBox 680, 400, 20, 70, "Retencion:", "F2", 8, hRight
        oDoc.WTextBox 700, 400, 20, 70, "I.V.A:", "F2", 8, hRight
        oDoc.WTextBox 720, 400, 20, 70, "TOTAL:", "F2", 8, hRight
        'totales
        oDoc.WTextBox 620, 488, 20, 50, Format(tRs1.Fields("TOTAL"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("IMPUESTO1")) Then oDoc.WTextBox 640, 488, 20, 50, Format(tRs1.Fields("IMPUESTO1"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("IMPUESTO2")) Then oDoc.WTextBox 660, 488, 20, 50, Format(tRs1.Fields("IMPUESTO2"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("RETENCION")) Then oDoc.WTextBox 680, 488, 20, 50, Format(tRs1.Fields("RETENCION"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("IVA")) Then oDoc.WTextBox 700, 488, 20, 50, Format(tRs1.Fields("IVA"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("TOTAL")) Then oDoc.WTextBox 720, 488, 20, 50, Format(tRs1.Fields("TOTAL"), "###,###,##0.00"), "F3", 8, hRight
        'oDoc.WTextBox 720, 200, 20, 250, "Firma de autorizado", "F3", 10, hCenter
        'If tRs1.Fields("CONFIRMADA") = "E" Then
        '    oDoc.WTextBox 750, 200, 25, 250, "COPIA CANCELADA", "F3", 25, hCenter, , vbBlue
        'End If
        'oDoc.WTextBox 620, 15, 20, 250, "Precios expresados en " & tRs1.Fields("MONEDA"), "F3", 10, hCenter
        oDoc.WTextBox 620, 20, 100, 275, "OBSERVACIONES:", "F3", 8, hLeft, , , 0, vbBlack
        If Not IsNull(tRs1.Fields("COMENTARIO")) Then oDoc.WTextBox 640, 60, 100, 300, tRs1.Fields("COMENTARIO"), "F3", 8, hLeft, , , 0, vbBlack
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs6 = cnn.Execute(sBuscar)
            If Not (tRs6.EOF And tRs6.BOF) Then
                oDoc.WTextBox 700, 20, 100, 275, "RESPONSABLE : ", "F3", 8, hLeft
                oDoc.WTextBox 720, 20, 100, 275, tRs6.Fields("NOMBRE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hLeft
            End If
        End If
        ' cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "No se enc ontro la orden de compra solicitada, puede ser que este cancelda o aun no se genere el folio", vbExclamation, "SACC"
    End If
End Sub
Private Sub FunImpATec()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim ConPag As Integer
    ConPag = 1
    Dim sBuscar As String
    sBuscar = "SELECT * FROM ASISTENCIA_TECNICA WHERE ID_AS_TEC = " & Text1.Text & ""
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\AsTec.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA"
        Set tRs1 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 20, 170, tRs1.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 205, 20, 170, tRs1.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "No. Asistencia : " & tRs.Fields("ID_AS_TEC"), "F3", 8, hCenter
        oDoc.WTextBox 70, 340, 20, 250, "Fecha :" & Format(tRs.Fields("FECHA_CAPTURA"), "dd/mm/yyyy"), "F3", 8, hCenter
        
        
        'CAJA1
        'sBuscar = "SELECT * FROM CLIENTE WHERE ID_CLIENTE = " & tRs1.Fields("ID_CLIENTE")
        'Set tRs2 = cnn.Execute(sBuscar)
        oDoc.WTextBox 110, 20, 100, 585, "CLIENTE : " & tRs.Fields("NOMBRE"), "F3", 8, hLeft
        oDoc.WTextBox 120, 20, 100, 585, "TELEFONO : " & tRs.Fields("TELEFONO"), "F3", 8, hLeft
        'If Not (tRs.EOF And tRs.BOF) Then
        '    If Not IsNull(tRs.Fields("NOMBRE")) Then oDoc.WTextBox 110, 20, 100, 400, tRs.Fields("NOMBRE"), "F3", 8, hCenter
        '    If Not IsNull(tRs.Fields("TELEFONO")) Then oDoc.WTextBox 120, 20, 100, 400, tRs.Fields("TELEFONO"), "F3", 8, hCenter
        'End If
        Posi = 150
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 10, 50, "MODELO", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 55, 10, 80, "MARCA", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 135, 10, 280, "TIPO DE ARTICULO", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 415, 10, 60, "GARANTIA", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 475, 10, 50, "DOMICILIO", "F2", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 525, 10, 65, "F. COMPROMISO", "F2", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 10
        oDoc.WTextBox Posi, 5, 20, 50, tRs.Fields("MODELO"), "F3", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 55, 20, 80, tRs.Fields("MARCA"), "F3", 8, hCenter, , , 1, vbCyan
        oDoc.WTextBox Posi, 135, 20, 280, tRs.Fields("TIPO_ARTICULO"), "F3", 8, hCenter, , , 1, vbCyan
        If tRs.Fields("GARANTIA") = "1" Then
            oDoc.WTextBox Posi, 415, 20, 60, "SI", "F3", 8, hCenter, , , 1, vbCyan
        Else
            oDoc.WTextBox Posi, 415, 20, 60, "NO", "F3", 8, hCenter, , , 1, vbCyan
        End If
        If tRs.Fields("A_DOMICILIO") = "1" Then
            oDoc.WTextBox Posi, 475, 20, 50, "SI", "F3", 8, hCenter, , , 1, vbCyan
        Else
            oDoc.WTextBox Posi, 475, 20, 50, "NO", "F3", 8, hCenter, , , 1, vbCyan
        End If
        oDoc.WTextBox Posi, 525, 20, 65, Format(tRs.Fields("FECHA_DEBE_ATENDER"), "dd/mm/yyyy"), "F3", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 20
        oDoc.WTextBox Posi, 5, 10, 585, "DESCRIPCION DE PIEZAS", "F2", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 10
        oDoc.WTextBox Posi, 5, 30, 585, tRs.Fields("DESCRIPCION_PIEZAS"), "F3", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 30
        oDoc.WTextBox Posi, 5, 10, 585, "COMENTARIOS TECNICOS", "F2", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 10
        oDoc.WTextBox Posi, 5, 30, 585, tRs.Fields("COMENTARIOS_TECNICOS"), "F3", 8, hCenter, , , 1, vbCyan
        Posi = Posi + 30
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 600
        oDoc.WLineTo 580, 600
        oDoc.LineStroke
        Posi = Posi + 6
        ' cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "No se encontrò la asistencia tècnica solicitada", vbExclamation, "SACC"
    End If
End Sub
