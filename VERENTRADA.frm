VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form VERENTRADA 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Entrada"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "VERENTRADA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CommonDialog1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3120
         TabIndex        =   0
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   960
         Width           =   4455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4935
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8705
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7800
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         PrinterDefault  =   0   'False
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Productor registrados en la entrada No. "
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9360
      TabIndex        =   5
      Top             =   5520
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "VERENTRADA.frx":001C
         MousePointer    =   99  'Custom
         Picture         =   "VERENTRADA.frx":0326
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label34 
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
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9360
      TabIndex        =   3
      Top             =   4320
      Width           =   975
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Command2 
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         MouseIcon       =   "VERENTRADA.frx":2408
         MousePointer    =   99  'Custom
         Picture         =   "VERENTRADA.frx":2712
         Top             =   240
         Width           =   720
      End
   End
End
Attribute VB_Name = "VERENTRADA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim FechaGurdado As String
Dim NombreUsuario As String
Private Sub Command2_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tLi As ListItem
    Dim fecha As Date
    sBuscar = "SELECT * FROM ENTRADAS  where ID_ENTRADA = ' " & Text2.Text & " '"
    Set tRs = cnn.Execute(sBuscar)
    fecha = tRs.Fields("FECHA")
    sBuscar = "SELECT * FROM VSENTRADAS2  where ID_ENTRADA = ' " & Text2.Text & " '"
    Set tRs3 = cnn.Execute(sBuscar)
    Dim Total As Double
    Total = 0
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
    Printer.Print "             FECHA : " & fecha
    Printer.Print "             SUCURSAL : BODEGA"
    Printer.Print "             IMPRESO POR : " & NombreUsuario
    Printer.Print "             FOLIO: " & Text2.Text
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO")) / 2
    Printer.Print "COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO"
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("NOMBRE DEL PROVEEDOR:  " & Text1.Text)) / 2
    Printer.Print "NOMBRE DEL PROVEEDOR:  " & Text1.Text
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Dim NRegistros As Integer
    NRegistros = ListView1.ListItems.Count
    Dim Con As Integer
    Dim POSY As Integer
    POSY = 3800
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Clave del Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3500
    Printer.Print "Cantidad Registrada"
    Printer.CurrentY = POSY
    Printer.CurrentX = 6500
    Printer.Print "Precio"
    Printer.CurrentY = POSY
    Printer.CurrentX = 7500
    Printer.Print "Sucursal"
    Printer.CurrentY = POSY
    Printer.CurrentX = 8800
    Printer.Print "Num Orden"
    Printer.CurrentY = POSY
    Printer.CurrentX = 10000
    Printer.Print "Factura"
    POSY = POSY + 200
    For Con = 1 To NRegistros
        POSY = POSY + 200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print ListView1.ListItems(Con).Text
        Printer.CurrentY = POSY
        Printer.CurrentX = 4000
        Printer.Print ListView1.ListItems(Con).SubItems(1)
        Printer.CurrentY = POSY
        Printer.CurrentX = 6500
        Printer.Print ListView1.ListItems(Con).SubItems(3)
        Printer.CurrentY = POSY
        Printer.CurrentX = 8800
        Printer.Print tRs3.Fields("NUM_ORDEN")
        Printer.CurrentY = POSY
        Printer.CurrentX = 10000
        Printer.Print tRs3.Fields("FACT_PROV")
        'Total = tRs.Fields("TOTAL")
        Total = Total + (Val(Replace(ListView1.ListItems(Con).SubItems(3), ",", "")) * Val(Replace(ListView1.ListItems(Con).SubItems(1), ",", "")))
        Printer.CurrentY = POSY
        Printer.CurrentX = 7500
        Printer.Print ListView1.ListItems(Con).SubItems(2)
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
            Printer.Print "             FECHA : " & FechaGurdado
            Printer.Print "             SUCURSAL : " & VarMen.Text4(0).Text
            Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
            Printer.Print "             FOLIO: " & Text2.Text
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
            Printer.Print "Cantidad Registrada"
            Printer.CurrentY = POSY
            Printer.CurrentX = 6500
            Printer.Print "Precio"
            Printer.CurrentY = POSY
            Printer.CurrentX = 7500
            Printer.Print "Sucursal"
        End If
    Next Con
    Printer.Print ""
    Printer.Print "             Total = " & Total
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.EndDoc
    CommonDialog1.Copies = 1
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    If EntradaProd.Text4.Text <> "" Then
        Text2.Text = EntradaProd.Text4.Text
    Else
        Text2.Text = InputBox("Numero de Entrada a ver:", "")
        If Text2.Text = "" Then
            Do While Text2.Text = ""
                Text2.Text = InputBox("Numero de Entrada a ver:", "")
            Loop
        End If
    End If
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
        .ColumnHeaders.Add , , "Clave del Producto", 3700
        .ColumnHeaders.Add , , "CANTIDAD", 2700
        .ColumnHeaders.Add , , "SUCURSAL", 3500
        .ColumnHeaders.Add , , "PRECIO", 2500
    End With
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT * FROM ENTRADA_PRODUCTO AS E JOIN SUCURSALES AS S ON E.ID_SUCURSAL = S.ID_SUCURSAL WHERE ID_ENTRADA =" & CDbl(Text2.Text)
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        ListView1.ListItems.Clear
        sBuscar = "SELECT ID_PROVEEDOR, FECHA, ID_USUARIO FROM ENTRADAS WHERE ID_ENTRADA = " & CDbl(Text2.Text)
        Set tRs1 = cnn.Execute(sBuscar)
        If Not (tRs1.EOF And tRs1.BOF) Then
            FechaGurdado = tRs1.Fields("FECHA")
            sBuscar = "SELECT NOMBRE FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
            Set tRs2 = cnn.Execute(sBuscar)
            If Not (tRs2.EOF And tRs2.BOF) Then
                Text1.Text = tRs2.Fields("NOMBRE")
            End If
            sBuscar = "SELECT NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs3 = cnn.Execute(sBuscar)
            If Not (tRs3.EOF And tRs3.BOF) Then
                NombreUsuario = tRs3.Fields("NOMBRE")
            End If
        End If
        Do While Not .EOF
            Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                tLi.SubItems(2) = .Fields("NOMBRE") & ""
                tLi.SubItems(3) = .Fields("PRECIO") & ""
            .MoveNext
        Loop
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Text1_Change()
    Command2.Enabled = False
    If Text1.Text <> "" And Text2.Text <> "" Then
        Command2.Enabled = True
    End If
End Sub
Private Sub Text2_Change()
    Command2.Enabled = False
    If Text1.Text <> "" And Text2.Text <> "" Then
        Command2.Enabled = True
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Busca Entrada
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        sBuscar = "SELECT * FROM ENTRADA_PRODUCTO AS E JOIN SUCURSALES AS S ON E.ID_SUCURSAL = S.ID_SUCURSAL WHERE ID_ENTRADA =" & CDbl(Text2.Text)
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            ListView1.ListItems.Clear
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                    tLi.SubItems(2) = .Fields("NOMBRE") & ""
                    tLi.SubItems(3) = .Fields("PRECIO") & ""
                .MoveNext
            Loop
        End With
    End If
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
