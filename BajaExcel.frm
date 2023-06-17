VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form BajaExcel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pasar listados a EXCEL"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3360
      TabIndex        =   10
      Top             =   2760
      Width           =   975
      Begin VB.Label Label26 
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
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "BajaExcel.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "BajaExcel.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Listado a Exportar"
      TabPicture(0)   =   "BajaExcel.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CommonDialog1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ProgressBar1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.CommandButton Command1 
         Caption         =   "EXCEL"
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
         Left            =   960
         Picture         =   "BajaExcel.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Productos de Almacen 1"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Productos de Almacen 2"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Productos de Almacen 3"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1200
         Width           =   2295
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Juegos de Reparacion"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   2175
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Clientes"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1920
         Width           =   2175
      End
      Begin VB.OptionButton Option6 
         Caption         =   " Proveedores"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   2280
         Width           =   1935
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2880
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2400
         Top             =   2400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Copiando..."
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   975
      End
   End
End
Attribute VB_Name = "BajaExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim FILE As String
Private Sub Command1_Click()
On Error GoTo ManejaError
    CommonDialog1.DialogTitle = "Guardar Como"
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "[Archivo Excel (*.xls)] |*.xls|"
    CommonDialog1.ShowOpen
    FILE = CommonDialog1.FileName
    Me.Label1.Visible = True
    Me.ProgressBar1.Visible = True
    If FILE <> "" Then
        If Option1.Value Then
            FunExAl1
        End If
        If Option2.Value Then
            FunExAl2
        End If
        If Option3.Value Then
            FunExAl3
        End If
        If Option4.Value Then
            FunExJueRep
        End If
        If Option5.Value Then
            FunExCli
        End If
        If Option6.Value Then
            FunExProv
        End If
        ProgressBar1.Value = 0
    End If
    Me.Label1.Visible = False
    Me.ProgressBar1.Visible = False
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Me.Label1.Visible = False
    Me.ProgressBar1.Visible = False
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub FunExCli()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim RecSe As ADODB.Recordset
    Dim Conta As Integer
    sBuscar = "SELECT COUNT(ID_CLIENTE) AS CUANTOS FROM CLIENTE"
    Set RecSe = cnn.Execute(sBuscar)
    ProgressBar1.Min = 0
    ProgressBar1.Max = RecSe.Fields("CUANTOS")
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT * FROM CLIENTE ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            MsgBox "No se han encontrado los datos buscados"
        Else
            .MoveFirst
            Dim ApExcel As Excel.Application
            Set ApExcel = CreateObject("Excel.application")
            ApExcel.Workbooks.Add
            Dim Cont As Integer
            Cont = 2
            ApExcel.Cells(1, 1) = "ID CLIENTE"
            ApExcel.Cells(1, 2) = "NOMBRE"
            ApExcel.Cells(1, 3) = "NOMBRE COMERCIAL"
            ApExcel.Cells(1, 4) = "RFC"
            ApExcel.Cells(1, 5) = "TELEFONO DE CASA"
            ApExcel.Cells(1, 6) = "TELEFONO DE OFICINA"
            ApExcel.Cells(1, 7) = "PAIS"
            ApExcel.Cells(1, 8) = "DIAS DE CREDITO"
            ApExcel.Cells(1, 9) = "FAX"
            ApExcel.Cells(1, 10) = "CONTACTO"
            ApExcel.Cells(1, 11) = "DIRECCION"
            ApExcel.Cells(1, 12) = "COMENTARIOS"
            ApExcel.Cells(1, 13) = "CIUDAD"
            ApExcel.Cells(1, 14) = "COLONIA"
            ApExcel.Cells(1, 15) = "DESCUENTO"
            ApExcel.Cells(1, 16) = "NUMERO EXTERIOR"
            ApExcel.Cells(1, 17) = "NUMERO INTERIOR"
            ApExcel.Cells(1, 18) = "CURP"
            ApExcel.Cells(1, 19) = "CP"
            ApExcel.Cells(1, 20) = "E-MAIL"
            ApExcel.Cells(1, 21) = "WEB PASSWORD"
            ApExcel.Cells(1, 22) = "ESTADO"
            ApExcel.Cells(1, 23) = "FECHA DE ALTA"
            ApExcel.Cells(1, 24) = "LIMITE DE CREDITO"
            ApExcel.Cells(1, 25) = "DIRECCION DE ENVIO"
            Do While Not .EOF
                Conta = Conta + 1
                ProgressBar1.Value = Conta
                ApExcel.Cells(Cont, 1) = .Fields("ID_CLIENTE") & ""
                ApExcel.Cells(Cont, 2) = .Fields("NOMBRE") & ""
                ApExcel.Cells(Cont, 3) = .Fields("NOMBRE_COMERCIAL") & ""
                ApExcel.Cells(Cont, 4) = .Fields("RFC") & ""
                ApExcel.Cells(Cont, 5) = .Fields("TELEFONO_CASA") & ""
                ApExcel.Cells(Cont, 6) = .Fields("TELEFONO_TRABAJO") & ""
                ApExcel.Cells(Cont, 7) = .Fields("PAIS") & ""
                ApExcel.Cells(Cont, 8) = .Fields("DIAS_CREDITO") & ""
                ApExcel.Cells(Cont, 9) = .Fields("FAX") & ""
                ApExcel.Cells(Cont, 10) = .Fields("CONTACTO") & ""
                ApExcel.Cells(Cont, 11) = .Fields("DIRECCION") & ""
                ApExcel.Cells(Cont, 12) = .Fields("COMENTARIOS") & ""
                ApExcel.Cells(Cont, 13) = .Fields("CIUDAD") & ""
                ApExcel.Cells(Cont, 14) = .Fields("COLONIA") & ""
                ApExcel.Cells(Cont, 15) = .Fields("DESCUENTO") & ""
                ApExcel.Cells(Cont, 16) = .Fields("NUMERO_EXTERIOR") & ""
                ApExcel.Cells(Cont, 17) = .Fields("NUMERO_INTERIOR") & ""
                ApExcel.Cells(Cont, 18) = .Fields("CURP") & ""
                ApExcel.Cells(Cont, 19) = .Fields("CP") & ""
                ApExcel.Cells(Cont, 20) = .Fields("EMAIL") & ""
                ApExcel.Cells(Cont, 21) = .Fields("WEB_PASSWORD") & ""
                ApExcel.Cells(Cont, 22) = .Fields("ESTADO") & ""
                ApExcel.Cells(Cont, 23) = .Fields("FECHA_ALTA") & ""
                ApExcel.Cells(Cont, 24) = .Fields("LIMITE_CREDITO") & ""
                ApExcel.Cells(Cont, 25) = .Fields("DIRECCION_ENVIO") & ""
                Cont = Cont + 1
                .MoveNext
            Loop
        End If
    End With
    ApExcel.AlertBeforeOverwriting = False
    ApExcel.ActiveWorkbook.SaveAs "" & FILE
    ApExcel.Visible = True
    Set ApExcel = Nothing
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub FunExAl3()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim RecSe As ADODB.Recordset
    Dim Conta As Integer
    sBuscar = "SELECT COUNT(ID_PRODUCTO) AS CUANTOS FROM ALMACEN3"
    Set RecSe = cnn.Execute(sBuscar)
    ProgressBar1.Min = 0
    ProgressBar1.Max = RecSe.Fields("CUANTOS")
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT * FROM ALMACEN3 ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            MsgBox "No se han encontrado los datos buscados"
        Else
            .MoveFirst
            Dim ApExcel As Excel.Application
            Set ApExcel = CreateObject("Excel.application")
            ApExcel.Workbooks.Add
            Dim Cont As Integer
            Cont = 2
            ApExcel.Cells(1, 1) = "ID PRODUCTO"
            ApExcel.Cells(1, 2) = "DESCRIPCION"
            ApExcel.Cells(1, 3) = "CANTIDAD MINIMA"
            ApExcel.Cells(1, 4) = "CANTIDAD MAXIMA"
            ApExcel.Cells(1, 5) = "TIPO"
            ApExcel.Cells(1, 6) = "VENTA WEB"
            ApExcel.Cells(1, 7) = "MARCA"
            ApExcel.Cells(1, 8) = "MATERIAL"
            ApExcel.Cells(1, 9) = "COLOR"
            ApExcel.Cells(1, 10) = "FOTO FRENTE"
            ApExcel.Cells(1, 11) = "FOTO LADO"
            ApExcel.Cells(1, 12) = "GANANCIA"
            ApExcel.Cells(1, 13) = "PRECIO DE COSTO"
            Do While Not .EOF
                Conta = Conta + 1
                ProgressBar1.Value = Conta
                ApExcel.Cells(Cont, 1) = .Fields("ID_PRODUCTO") & ""
                ApExcel.Cells(Cont, 2) = .Fields("DESCRIPCION") & ""
                ApExcel.Cells(Cont, 3) = .Fields("C_MINIMA") & ""
                ApExcel.Cells(Cont, 4) = .Fields("C_MAXIMA") & ""
                ApExcel.Cells(Cont, 5) = .Fields("TIPO") & ""
                ApExcel.Cells(Cont, 6) = .Fields("VENTA_WEB") & ""
                ApExcel.Cells(Cont, 7) = .Fields("MARCA") & ""
                ApExcel.Cells(Cont, 8) = .Fields("MATERIAL") & ""
                ApExcel.Cells(Cont, 9) = .Fields("COLOR") & ""
                ApExcel.Cells(Cont, 10) = .Fields("FOTO_FRENTE") & ""
                ApExcel.Cells(Cont, 11) = .Fields("FOTO_LADO") & ""
                ApExcel.Cells(Cont, 12) = .Fields("GANANCIA") & ""
                ApExcel.Cells(Cont, 13) = .Fields("PRECIO_COSTO") & ""
                Cont = Cont + 1
                .MoveNext
            Loop
        End If
    End With
    ApExcel.AlertBeforeOverwriting = False
    ApExcel.ActiveWorkbook.SaveAs "" & FILE
    ApExcel.Visible = True
    Set ApExcel = Nothing
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub FunExAl2()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim RecSe As ADODB.Recordset
    Dim Conta As Integer
    sBuscar = "SELECT COUNT(ID_PRODUCTO) AS CUANTOS FROM ALMACEN2"
    Set RecSe = cnn.Execute(sBuscar)
    ProgressBar1.Min = 0
    ProgressBar1.Max = RecSe.Fields("CUANTOS")
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT * FROM ALMACEN2 ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            MsgBox "No se han encontrado los datos buscados"
        Else
            .MoveFirst
            Dim ApExcel As Excel.Application
            Set ApExcel = CreateObject("Excel.application")
            ApExcel.Workbooks.Add
            Dim Cont As Integer
            Cont = 2
            ApExcel.Cells(1, 1) = "ID PRODUCTO"
            ApExcel.Cells(1, 2) = "DESCRIPCION"
            ApExcel.Cells(1, 3) = "CANTIDAD MINIMA"
            ApExcel.Cells(1, 4) = "CANTIDAD MAXIMA"
            ApExcel.Cells(1, 5) = "TIPO"
            ApExcel.Cells(1, 6) = "VENTA WEB"
            ApExcel.Cells(1, 7) = "MARCA"
            ApExcel.Cells(1, 8) = "MATERIAL"
            ApExcel.Cells(1, 9) = "COLOR"
            ApExcel.Cells(1, 10) = "FOTO FRENTE"
            ApExcel.Cells(1, 11) = "FOTO LADO"
            Do While Not .EOF
                Conta = Conta + 1
                ProgressBar1.Value = Conta
                ApExcel.Cells(Cont, 1) = .Fields("ID_PRODUCTO") & ""
                ApExcel.Cells(Cont, 2) = .Fields("DESCRIPCION") & ""
                ApExcel.Cells(Cont, 3) = .Fields("C_MINIMA") & ""
                ApExcel.Cells(Cont, 4) = .Fields("C_MAXIMA") & ""
                ApExcel.Cells(Cont, 5) = .Fields("TIPO") & ""
                ApExcel.Cells(Cont, 6) = .Fields("VENTA_WEB") & ""
                ApExcel.Cells(Cont, 7) = .Fields("MARCA") & ""
                ApExcel.Cells(Cont, 8) = .Fields("MATERIAL") & ""
                ApExcel.Cells(Cont, 9) = .Fields("COLOR") & ""
                ApExcel.Cells(Cont, 10) = .Fields("FOTO_FRENTE") & ""
                ApExcel.Cells(Cont, 11) = .Fields("FOTO_LADO") & ""
                Cont = Cont + 1
                .MoveNext
            Loop
        End If
    End With
    ApExcel.AlertBeforeOverwriting = False
    ApExcel.ActiveWorkbook.SaveAs "" & FILE
    ApExcel.Visible = True
    Set ApExcel = Nothing
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub FunExAl1()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim RecSe As ADODB.Recordset
    Dim Conta As Integer
    sBuscar = "SELECT COUNT(ID_PRODUCTO) AS CUANTOS FROM ALMACEN1"
    Set RecSe = cnn.Execute(sBuscar)
    ProgressBar1.Min = 0
    ProgressBar1.Max = RecSe.Fields("CUANTOS")
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT * FROM ALMACEN1 ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            MsgBox "No se han encontrado los datos buscados"
        Else
            .MoveFirst
            Dim ApExcel As Excel.Application
            Set ApExcel = CreateObject("Excel.application")
            ApExcel.Workbooks.Add
            Dim Cont As Integer
            Cont = 2
            ApExcel.Cells(1, 1) = "ID PRODUCTO"
            ApExcel.Cells(1, 2) = "DESCRIPCION"
            ApExcel.Cells(1, 3) = "CANTIDAD MINIMA"
            ApExcel.Cells(1, 4) = "CANTIDAD MAXIMA"
            ApExcel.Cells(1, 5) = "TIPO"
            ApExcel.Cells(1, 6) = "VENTA WEB"
            ApExcel.Cells(1, 7) = "MARCA"
            ApExcel.Cells(1, 8) = "MATERIAL"
            ApExcel.Cells(1, 9) = "COLOR"
            ApExcel.Cells(1, 10) = "FOTO FRENTE"
            ApExcel.Cells(1, 11) = "FOTO LADO"
            Do While Not .EOF
                Conta = Conta + 1
                ProgressBar1.Value = Conta
                ApExcel.Cells(Cont, 1) = .Fields("ID_PRODUCTO") & ""
                ApExcel.Cells(Cont, 2) = .Fields("DESCRIPCION") & ""
                ApExcel.Cells(Cont, 3) = .Fields("C_MINIMA") & ""
                ApExcel.Cells(Cont, 4) = .Fields("C_MAXIMA") & ""
                ApExcel.Cells(Cont, 5) = .Fields("TIPO") & ""
                ApExcel.Cells(Cont, 6) = .Fields("VENTA_WEB") & ""
                ApExcel.Cells(Cont, 7) = .Fields("MARCA") & ""
                ApExcel.Cells(Cont, 8) = .Fields("MATERIAL") & ""
                ApExcel.Cells(Cont, 9) = .Fields("COLOR") & ""
                ApExcel.Cells(Cont, 10) = .Fields("FOTO_FRENTE") & ""
                ApExcel.Cells(Cont, 11) = .Fields("FOTO_LADO") & ""
                Cont = Cont + 1
                .MoveNext
            Loop
        End If
    End With
    ApExcel.AlertBeforeOverwriting = False
    ApExcel.ActiveWorkbook.SaveAs "" & FILE
    ApExcel.Visible = True
    Set ApExcel = Nothing
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub FunExProv()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim RecSe As ADODB.Recordset
    Dim Conta As Integer
    sBuscar = "SELECT COUNT(ID_PROVEEDOR) AS CUANTOS FROM PROVEEDOR"
    Set RecSe = cnn.Execute(sBuscar)
    ProgressBar1.Min = 0
    ProgressBar1.Max = RecSe.Fields("CUANTOS")
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT * FROM PROVEEDOR ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            MsgBox "No se han encontrado los datos buscados"
        Else
            .MoveFirst
            Dim ApExcel As Excel.Application
            Set ApExcel = CreateObject("Excel.application")
            ApExcel.Workbooks.Add
            Dim Cont As Integer
            Cont = 2
            ApExcel.Cells(1, 1) = "ID PREOVEEDOR"
            ApExcel.Cells(1, 2) = "NOMBRE"
            ApExcel.Cells(1, 3) = "DIRECCION"
            ApExcel.Cells(1, 4) = "COLONIA"
            ApExcel.Cells(1, 5) = "CIUDAD"
            ApExcel.Cells(1, 6) = "CP"
            ApExcel.Cells(1, 7) = "RFC"
            ApExcel.Cells(1, 8) = "TELEFONO1"
            ApExcel.Cells(1, 9) = "TELEFONO2"
            ApExcel.Cells(1, 10) = "TELEFONO3"
            ApExcel.Cells(1, 11) = "NOTAS"
            ApExcel.Cells(1, 12) = "ESTADO"
            ApExcel.Cells(1, 13) = "PAIS"
            ApExcel.Cells(1, 14) = "TRANS_BANCO"
            ApExcel.Cells(1, 15) = "TRANS_DIRECCION"
            ApExcel.Cells(1, 16) = "TRANS_CIUDAD"
            ApExcel.Cells(1, 17) = "TRANS_ROUTING"
            ApExcel.Cells(1, 18) = "TRANS_CUENTA"
            ApExcel.Cells(1, 19) = "TRANS_CLAVE_SWIFT"
            Do While Not .EOF
                Conta = Conta + 1
                ProgressBar1.Value = Conta
                ApExcel.Cells(Cont, 1) = .Fields("ID_PROVEEDOR") & ""
                ApExcel.Cells(Cont, 2) = .Fields("NOMBRE") & ""
                ApExcel.Cells(Cont, 3) = .Fields("DIRECCION") & ""
                ApExcel.Cells(Cont, 4) = .Fields("COLONIA") & ""
                ApExcel.Cells(Cont, 5) = .Fields("CIUDAD") & ""
                ApExcel.Cells(Cont, 6) = .Fields("CP") & ""
                ApExcel.Cells(Cont, 7) = .Fields("RFC") & ""
                ApExcel.Cells(Cont, 8) = .Fields("TELEFONO1") & ""
                ApExcel.Cells(Cont, 9) = .Fields("TELEFONO2") & ""
                ApExcel.Cells(Cont, 10) = .Fields("TELEFONO3") & ""
                ApExcel.Cells(Cont, 11) = .Fields("NOTAS") & ""
                ApExcel.Cells(Cont, 12) = .Fields("ESTADO") & ""
                ApExcel.Cells(Cont, 13) = .Fields("PAIS") & ""
                ApExcel.Cells(Cont, 14) = .Fields("TRANS_BANCO") & ""
                ApExcel.Cells(Cont, 15) = .Fields("TRANS_DIRECCION") & ""
                ApExcel.Cells(Cont, 16) = .Fields("TRANS_CIUDAD") & ""
                ApExcel.Cells(Cont, 17) = .Fields("TRANS_ROUTING") & ""
                ApExcel.Cells(Cont, 18) = .Fields("TRANS_CUENTA") & ""
                ApExcel.Cells(Cont, 19) = .Fields("TRANS_CLAVE_SWIFT") & ""
                Cont = Cont + 1
                .MoveNext
            Loop
        End If
    End With
    ApExcel.AlertBeforeOverwriting = False
    ApExcel.ActiveWorkbook.SaveAs "" & FILE
    ApExcel.Visible = True
    Set ApExcel = Nothing
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub FunExJueRep()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim RecSe As ADODB.Recordset
    Dim Conta As Integer
    Dim Cont As Integer
    Dim tRs As ADODB.Recordset
    Dim ComparaProducto As String
    Dim ApExcel As Excel.Application
    sBuscar = "SELECT COUNT(ID_PRODUCTO) AS CUANTOS FROM JUEGO_REPARACION"
    Set RecSe = cnn.Execute(sBuscar)
    ProgressBar1.Min = 0
    ProgressBar1.Max = RecSe.Fields("CUANTOS")
    sBuscar = "SELECT * FROM JUEGO_REPARACION ORDER BY ID_REPARACION"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            MsgBox "No se han encontrado los datos buscados"
        Else
            .MoveFirst
            Set ApExcel = CreateObject("Excel.application")
            ApExcel.Workbooks.Add
            Cont = 2
            ApExcel.Cells(1, 1).Columns.ColumnWidth = 29
            ApExcel.Cells(1, 2).Columns.ColumnWidth = 29
            ApExcel.Cells(1, 1) = "PRODUCTO"
            ApExcel.Cells(1, 2) = "CONSUMIBLES"
            ApExcel.Cells(1, 3) = "CANTIDAD"
            Do While Not .EOF
                Conta = Conta + 1
                ProgressBar1.Value = Conta
                If ComparaProducto <> .Fields("ID_REPARACION") Then
                    Cont = Cont + 1
                    ApExcel.Cells(Cont, 1).Font.Size = 12
                    ApExcel.Cells(Cont, 1) = .Fields("ID_REPARACION") & ""
                    ComparaProducto = .Fields("ID_REPARACION")
                    ApExcel.Range("A" & Cont & ":C" & Cont).Borders.Color = RGB(0, 0, 255)
                Else
                    ApExcel.Range("A" & Cont & ":C" & Cont).Borders.Color = RGB(0, 0, 0)
                End If
                ApExcel.Cells(Cont, 2) = .Fields("ID_PRODUCTO") & ""
                ApExcel.Cells(Cont, 3) = .Fields("CANTIDAD") & ""
                Cont = Cont + 1
                .MoveNext
            Loop
        End If
    End With
    ApExcel.AlertBeforeOverwriting = False
    ApExcel.ActiveWorkbook.SaveAs "" & FILE
    ApExcel.Visible = True
    Set ApExcel = Nothing
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Option1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 14 Then
        Option1.Value = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Option2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 14 Then
        Option2.Value = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Option3_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 14 Then
        Option3.Value = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Option4_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 14 Then
        Option4.Value = True
    End If
Exit Sub
ManejaError:
     MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Option5_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 14 Then
        Option5.Value = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Option6_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 14 Then
        Option6.Value = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
