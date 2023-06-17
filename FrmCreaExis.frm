VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCreaExis 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crear Producto de 2 o Mas (Almacen3)"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   30
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
         TabIndex        =   31
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCreaExis.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCreaExis.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   1
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmCreaExis.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmCreaExis.frx":26F6
         Top             =   240
         Width           =   675
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Seleccion"
      TabPicture(0)   =   "FrmCreaExis.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ListView3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdAgregar"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text6"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text7"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CommonDialog1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Combo1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "FrmCreaExis.frx":40D4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "Text8"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "ListView2"
      Tab(1).Control(4)=   "Label9"
      Tab(1).ControlCount=   5
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   5640
         TabIndex        =   32
         Top             =   480
         Width           =   2415
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         PrinterDefault  =   0   'False
      End
      Begin VB.CommandButton Command1 
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
         Left            =   -70680
         Picture         =   "FrmCreaExis.frx":40F0
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   6120
         Width           =   2775
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   6120
         Width           =   2775
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   5520
         TabIndex        =   24
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregar 
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
         Left            =   6840
         Picture         =   "FrmCreaExis.frx":6AC2
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   6000
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1815
         Left            =   120
         TabIndex        =   21
         Top             =   4080
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3201
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   6360
         TabIndex        =   18
         Top             =   3360
         Width           =   1695
         Begin VB.OptionButton Option4 
            Caption         =   "Por Descripción"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Por Clave"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   3720
         Width           =   4335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Producto a Formar"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   7935
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1080
            TabIndex        =   15
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   960
            Width           =   6375
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label4 
            Caption         =   "Cantidad :"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Descripción :"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Clave"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   8
         Top             =   2640
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5741
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
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   6480
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   6480
         TabIndex        =   6
         Top             =   960
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   1080
         Width           =   4815
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   7935
         _ExtentX        =   13996
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
      Begin VB.Label Label10 
         Caption         =   "Sucursal de existencias"
         Height          =   255
         Left            =   3600
         TabIndex        =   33
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Seleccionado"
         Height          =   255
         Left            =   -74760
         TabIndex        =   27
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Producto Seleccionado :"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   6120
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   4800
         TabIndex        =   23
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Productos Componentes :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Producto a Formar :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmCreaExis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim DesProd As String
Dim Exis As String
Dim InEli As Integer
Dim IdProdEl As String
Dim CantEl As String
Dim SucEl As String
Private Sub cmdAgregar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If Text6.Text <> "" And Combo1.Text <> "" And Text7.Text <> "" Then
        If CDbl(Text6.Text) > 0 Then
            If Exis >= CDbl(Text6.Text) Then
                Set tLi = ListView2.ListItems.Add(, , Text7.Text & "")
                tLi.SubItems(1) = DesProd
                tLi.SubItems(2) = Text6.Text
                tLi.SubItems(3) = Exis
                tLi.SubItems(4) = Combo1.Text
                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD - " & Text6.Text & " WHERE ID_PRODUCTO = '" & Text7.Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
                Set tRs = cnn.Execute(sBuscar)
                Text7.Text = ""
                DesProd = ""
                Text6.Text = ""
            Else
                MsgBox "NO CUENTA CON EXISTENCIA SUFICIENTE PARA SURTIR!", vbInformation, "SACC"
            End If
        End If
    Else
        MsgBox "Falta información necesaria para el registro", vbInformation, "SACC"
    End If
End Sub
Private Sub Command1_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If InEli > 0 Then
        sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & IdProdEl & "' AND SUCURSAL = '" & SucEl & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD  + " & CantEl & " WHERE ID_PRODUCTO = '" & IdProdEl & "' AND SUCURSAL = '" & SucEl & "'"
            Set tRs = cnn.Execute(sBuscar)
        Else
            sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & IdProdEl & "', " & CantEl & ", '" & SucEl & "');"
            cnn.Execute (sBuscar)
        End If
        ListView2.ListItems.Remove (InEli)
        InEli = 0
        Text8.Text = ""
    Else
        MsgBox "NO HA SELECCIONADO LOS ARTICULOS PARA ELIMINAR!", vbInformation, "SACC"
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave Producto", 2500
        .ColumnHeaders.Add , , "Descripcion", 5600
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave Producto", 2500
        .ColumnHeaders.Add , , "Descripcion", 5600
        .ColumnHeaders.Add , , "Cantidad a Tomar", 1700
        .ColumnHeaders.Add , , "Cantidad a Existencia", 0
        .ColumnHeaders.Add , , "Sucursal", 0
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave Producto", 2500
        .ColumnHeaders.Add , , "Descripcion", 5600
        .ColumnHeaders.Add , , "Existencia", 1700
    End With
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    sBuscar = "SELECT NOMBRE From SUCURSALES WHERE ELIMINADO = 'N' GROUP BY NOMBRE ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            If Trim(tRs.Fields("NOMBRE")) <> "" Then Combo1.AddItem (tRs.Fields("NOMBRE"))
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image8_Click()
On Error GoTo ManejaError
    If Text8.Text <> Text2.Text Then
        If ListView2.ListItems.Count > 0 And Text2.Text <> "" And Text4.Text <> "" And Combo1.Text <> "" Then
            Dim sBuscar As String
            Dim tRs As ADODB.Recordset
            Dim NRegistros As Integer
            Dim Con As Integer
            Dim Aux As String
            Dim NueExi As String
            Dim Folio
            Text2.Text = Replace(Text2.Text, ",", "")
            Text4.Text = Replace(Text4.Text, ",", "")
            sBuscar = "INSERT INTO PRODUCCIONES_ALMACEN3 (ID_PRODUCTO, CANTIDAD, FECHA, ID_USUARIO) VALUES ('" & Text2.Text & "', " & Text4.Text & ", '" & Format(Date, "dd/mm/yyyy") & "', " & VarMen.Text1(0).Text & " );"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT ID_FABRICACION FROM PRODUCCIONES_ALMACEN3 ORDER BY ID_FABRICACION DESC"
            Set tRs = cnn.Execute(sBuscar)
            Folio = tRs.Fields("ID_FABRICACION")
            sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Text2.Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                NueExi = tRs.Fields("CANTIDAD")
                NueExi = CDbl(NueExi) + CDbl(Text4.Text)
                NueExi = Replace(NueExi, ",", "")
                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & NueExi & " WHERE ID_PRODUCTO = '" & Text2.Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
                Set tRs = cnn.Execute(sBuscar)
            Else
                NueExi = Replace(Text4.Text, ",", "")
                sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & Text2.Text & "', " & NueExi & ", '" & Combo1.Text & "');"
                cnn.Execute (sBuscar)
            End If
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
            Printer.Print ""
            Printer.Print ""
            Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
            Printer.Print "             SUCURSAL : " & VarMen.Text4(0).Text
            Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
            Printer.Print "             FOLIO: " & Folio
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO")) / 2
            Printer.Print "COMPROBANTE DE REGISTRO DE CONVERSION DE PRODUCTOS"
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Dim POSY As Integer
            POSY = 3800
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print "Clave del Producto"
            Printer.CurrentY = POSY
            Printer.CurrentX = 3500
            Printer.Print "Cantidad Usada"
            NRegistros = ListView2.ListItems.Count
            For Con = 1 To NRegistros
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print ListView2.ListItems(Con).Text
                Printer.CurrentY = POSY
                Printer.CurrentX = 4000
                Printer.Print ListView2.ListItems(Con).SubItems(2)
                Aux = Replace(ListView2.ListItems(Con).SubItems(2), ",", "")
                NueExi = CDbl(ListView2.ListItems(Con).SubItems(3)) - CDbl(Aux)
                NueExi = Replace(NueExi, ",", "")
                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD - " & Aux & " WHERE ID_PRODUCTO = '" & ListView2.ListItems(Con).Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
                Set tRs = cnn.Execute(sBuscar)
                sBuscar = "INSERT INTO PRODUCCIONES_ALMACEN3_DETALLE (ID_PRODUCTO, CANTIDAD, ID_FABRICACION) VALUES ('" & ListView2.ListItems(Con).Text & "', " & Aux & ", " & Folio & ");"
                cnn.Execute (sBuscar)
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
                    Printer.Print VarMen
                    Printer.Print ""
                    Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
                    Printer.Print "             SUCURSAL : " & VarMen.Text4(0).Text
                    Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
                    Printer.Print "             FOLIO: " & Folio
                    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO")) / 2
                    Printer.Print "COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO"
                    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth("NOMBRE DEL PROVEEDOR:  " & Text1.Text)) / 2
                    Printer.Print "NOMBRE DEL PROVEEDOR:  " & Text1.Text
                    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    POSY = 3800
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 100
                    Printer.Print "Clave del Producto"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 3500
                    Printer.Print "Cantidad Usada"
                End If
            Next Con
            Printer.Print ""
            Printer.Print "             Producto Formado = " & Text2.Text
            Printer.Print "             Cantidad Formada = " & Text4.Text
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.EndDoc
            ListView2.ListItems.Clear
            Text7.Text = ""
            Text6.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            CommonDialog1.Copies = 1
        Else
            MsgBox "FALTA INFORMACIÓN NECESARIA PARA EL REGISTRO!", vbInformation, "SACC"
        End If
    Else
        MsgBox "No es posible transformar un producto por el mismo producto", vbExclamation, "SACC"
    End If
    Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub Image9_Click()
    If ListView2.ListItems.Count = 0 Then
        Unload Me
    Else
        MsgBox "Elimine los productos de la lista de agregados para poder salir", vbExclamation, "SACC"
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text2.Text = Item
    Text3.Text = Item.SubItems(1)
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text8.Text = Item
    InEli = Item.Index
    IdProdEl = Item
    CantEl = Item.SubItems(2)
    SucEl = Item.SubItems(4)
    Exis = Item.SubItems(3)
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text7.Text = Item
    DesProd = Item.SubItems(1)
    Exis = Item.SubItems(2)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        If Option1.Value = True Then
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%'"
        Else
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE Descripcion LIKE '%" & Text1.Text & "%'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            ListView1.ListItems.Clear
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                .MoveNext
            Loop
        End With
        If Option1.Value = True Then
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%'"
        Else
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN2 WHERE Descripcion LIKE '%" & Text1.Text & "%'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                .MoveNext
            Loop
        End With
    End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    If Combo1.Text <> "" Then
        If KeyAscii = 13 Then
            Dim sBuscar As String
            Dim tRs As ADODB.Recordset
            Dim tLi As ListItem
            If Option3.Value = True Then
                sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM VSEXISALMA3 WHERE ID_PRODUCTO LIKE '%" & Text5.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
            Else
                sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM VSEXISALMA3 WHERE Descripcion LIKE '%" & Text5.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
            End If
            Set tRs = cnn.Execute(sBuscar)
            With tRs
                ListView3.ListItems.Clear
                Do While Not .EOF
                    Set tLi = ListView3.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                    If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = .Fields("CANTIDAD") & ""
                    .MoveNext
                Loop
            End With
            If Option3.Value = True Then
                sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM VSEXISALMA2 WHERE ID_PRODUCTO LIKE '%" & Text5.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
            Else
                sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM VSEXISALMA2 WHERE Descripcion LIKE '%" & Text5.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
            End If
            Set tRs = cnn.Execute(sBuscar)
            With tRs
                Do While Not .EOF
                    Set tLi = ListView3.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                    If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = .Fields("CANTIDAD") & ""
                    .MoveNext
                Loop
            End With
        End If
    Else
        MsgBox "Seleccione una sucursal", vbExclamation, "SACC"
    End If
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
