VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmImpCotiza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Cotizacion a Cliente"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6495
      Left            =   9480
      ScaleHeight     =   6435
      ScaleWidth      =   1395
      TabIndex        =   8
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
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
         Left            =   120
         Picture         =   "FrmImpCotiza.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
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
         Left            =   120
         Picture         =   "FrmImpCotiza.frx":29D2
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
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
         Height          =   375
         Left            =   120
         Picture         =   "FrmImpCotiza.frx":53A4
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
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
         Left            =   120
         Picture         =   "FrmImpCotiza.frx":7D76
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton Cerrar 
         Caption         =   "Cerrar"
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
         Left            =   120
         Picture         =   "FrmImpCotiza.frx":A748
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   240
         TabIndex        =   9
         Top             =   4680
         Width           =   975
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmImpCotiza.frx":D11A
            MousePointer    =   99  'Custom
            Picture         =   "FrmImpCotiza.frx":D424
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
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
            TabIndex        =   10
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView4 
      Height          =   2415
      Left            =   4800
      TabIndex        =   7
      Top             =   3480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4260
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
   Begin MSComctlLib.ListView ListView3 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4260
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
      Height          =   2535
      Left            =   4800
      TabIndex        =   3
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4471
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4471
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Listado 
      Caption         =   "Precio para el cliente"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Cotizado con"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   9360
      X2              =   240
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label2 
      Caption         =   "Articulos"
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Cotizaciones Pendientes"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "FrmImpCotiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim it0 As String
Dim it1 As String
Dim it2 As String
Dim it3 As String
Dim it4 As String
Dim it5 As String
Dim it6 As String
Dim INDI As Integer
Dim IdCotiza As String
Dim Nomb As String
Dim Dire As String
Dim Colo As String
Dim Ciud As String
Dim Tele As String
Dim Come As String
Private Sub Cerrar_Click()
On Error GoTo ManejaError
    Operacion
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim tLi As ListItem
    Set tLi = ListView4.ListItems.Add(, , it0 & "")
    tLi.SubItems(1) = it1 & ""
    tLi.SubItems(2) = it2 & ""
    tLi.SubItems(3) = it3 & ""
    tLi.SubItems(4) = it4 & ""
    tLi.SubItems(5) = it5 & ""
    tLi.SubItems(6) = it6 & ""
    Me.Command1.Enabled = False
    Me.Command3.Enabled = True
    Me.Command4.Enabled = True
    Me.Cerrar.Enabled = True
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cmdCancelar_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    Dim Archi As String
    CommonDialog1.DialogTitle = "Buscando"
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "[Archivo Excel (*.xls)] |*.xls|"
    CommonDialog1.ShowOpen
    Archi = CommonDialog1.FileName
    Dim ApExcel As Excel.Application
    Set ApExcel = CreateObject("Excel.application")
    ApExcel.Workbooks.Add
    ApExcel.Cells(1, 1) = "Calve del Producto"
    ApExcel.Cells(1, 2) = "Descripción"
    ApExcel.Cells(1, 3) = "Cantidad"
    ApExcel.Cells(1, 4) = "Precio"
    Dim NumeroRegistros As Integer
    NumeroRegistros = ListView4.ListItems.Count
    Dim Conta As Integer
    For Conta = 1 To NumeroRegistros
        ApExcel.Cells(1, 1) = ListView4.ListItems(Conta).SubItems(3)
        ApExcel.Cells(1, 2) = ListView4.ListItems(Conta).SubItems(4)
        ApExcel.Cells(1, 3) = ListView4.ListItems(Conta).SubItems(5)
        ApExcel.Cells(1, 4) = ListView4.ListItems(Conta).SubItems(6)
    Next Conta
    ApExcel.AlertBeforeOverwriting = False
    ApExcel.ActiveWorkbook.SaveAs "" & Archi
    ApExcel.Visible = True
    Set ApExcel = Nothing
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Command4_Click()
On Error GoTo ManejaError
    CommonDialog1.Flags = 64
    CommonDialog1.ShowPrinter
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "                                                                                       ACTITUD POSITIVA EN TONER S DE RL MI"
    Printer.Print "                                                                                                        R.F.C APT-040201-KA5"
    Printer.Print "                                                 ORTIZ DE CAMPOS No. 1308 COL. SAN FELIPE CHIHUAHUA, CHIHUAHUA C.P. 31203"
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "Fecha : " & Date
    Printer.CurrentY = 1100
    Printer.CurrentX = 100
    Printer.Print "Clave del Prod."
    Printer.CurrentY = 1100
    Printer.CurrentX = 2800
    Printer.Print "Cantidad"
    Printer.CurrentY = 1100
    Printer.CurrentX = 5000
    Printer.Print "Precio"
    Dim NumeroRegistros As Integer
    NumeroRegistros = ListView4.ListItems.Count
    Dim Conta As Integer
    Dim X As Integer
    X = 1300
    Dim CoN As Double
    CoN = 0
    For Conta = 1 To NumeroRegistros
        Printer.CurrentY = X
        Printer.CurrentX = 100
        Printer.Print ListView4.ListItems(Conta).SubItems(3)
        Printer.CurrentY = X
        Printer.CurrentX = 5000
        Printer.Print ListView4.ListItems(Conta).SubItems(5)
        Printer.CurrentY = X
        Printer.CurrentX = 8000
        Printer.Print ListView4.ListItems(Conta).SubItems(6)
        X = X + 200
        CoN = CoN + (CDbl(ListView4.ListItems(Conta).SubItems(5)) * CDbl(ListView4.ListItems(Conta).SubItems(6)))
        If X >= 14400 Then
            Printer.EndDoc
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print "                                                                                       ACTITUD POSITIVA EN TONER S DE RL MI"
            Printer.Print "                                                                                                        R.F.C APT-040201-KA5"
            Printer.Print "                                                 ORTIZ DE CAMPOS No. 1308 COL. SAN FELIPE CHIHUAHUA, CHIHUAHUA C.P. 31203"
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print "Fecha : " & Date
            Printer.CurrentY = 1100
            Printer.CurrentX = 100
            Printer.Print "Clave del Prod."
            Printer.CurrentY = 1100
            Printer.CurrentX = 5000
            Printer.Print "Cantidad"
            Printer.CurrentY = 1100
            Printer.CurrentX = 8000
            Printer.Print "Precio"
            X = 1300
        End If
    Next Conta
    Printer.CurrentY = X
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.CurrentY = X + 200
    Printer.CurrentX = 7600
    Printer.Print "Total : $ " & CoN
    Printer.EndDoc
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Command5_Click()
On Error GoTo ManejaError
    ListView4.ListItems.Remove (INDI)
    Me.Command5.Enabled = False
    If ListView4.ListItems.Count = 0 Then
        Me.Command3.Enabled = False
        Me.Command4.Enabled = False
        Me.Cerrar.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Form_Load()
On Error GoTo ManejaError
    Me.Command1.Enabled = False
    Me.Command3.Enabled = False
    Me.Command4.Enabled = False
    Me.Cerrar.Enabled = False
    Me.Command5.Enabled = False
    Const sPathBase As String = "LINUX"
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "No. COTIZACION", 0
        .ColumnHeaders.Add , , "No. CLIENTE", 0
        .ColumnHeaders.Add , , "NOMBRE", 3700
        .ColumnHeaders.Add , , "DIRECCION", 1500
        .ColumnHeaders.Add , , "COLONIA", 1500
        .ColumnHeaders.Add , , "CIUDAD", 1500
        .ColumnHeaders.Add , , "TELEFONO", 1000
        .ColumnHeaders.Add , , "COMENTARIOS", 1000
    End With
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "No. COTIZACION", 0
        .ColumnHeaders.Add , , "CLAVE", 2000
        .ColumnHeaders.Add , , "DESCRIPCION", 3700
        .ColumnHeaders.Add , , "CANTIDAD", 1500
    End With
        With ListView3
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "No. COTIZACION", 0
        .ColumnHeaders.Add , , "No. PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE", 3700
        .ColumnHeaders.Add , , "CLAVE", 1500
        .ColumnHeaders.Add , , "DESCRIPCION", 1500
        .ColumnHeaders.Add , , "CANTIDAD", 1500
        .ColumnHeaders.Add , , "PRECIO", 1000
    End With
    With ListView4
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "No. COTIZACION", 0
        .ColumnHeaders.Add , , "No. PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE", 3700
        .ColumnHeaders.Add , , "CLAVE", 1500
        .ColumnHeaders.Add , , "DESCRIPCION", 1500
        .ColumnHeaders.Add , , "CANTIDAD", 1500
        .ColumnHeaders.Add , , "PRECIO", 1000
    End With
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_COTIZACION, ID_CLIENTE, NOMBRE, DIRECCION, COLONIA, CIUDAD, TELEFONO, COMENTARIOS FROM COTIZACION WHERE PENDIENTE = 'S'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        ListView1.ListItems.Clear
        tRs.MoveFirst
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_COTIZACION") & "")
            tLi.SubItems(1) = tRs.Fields("ID_CLIENTE") & ""
            tLi.SubItems(2) = tRs.Fields("NOMBRE") & ""
            tLi.SubItems(3) = tRs.Fields("DIRECCION") & ""
            tLi.SubItems(4) = tRs.Fields("COLONIA") & ""
            tLi.SubItems(5) = tRs.Fields("CIUDAD") & ""
            tLi.SubItems(6) = tRs.Fields("TELEFONO") & ""
            tLi.SubItems(7) = tRs.Fields("COMENTARIOS") & ""
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
    Nomb = Item.SubItems(2)
    Dire = Item.SubItems(3)
    Colo = Item.SubItems(4)
    Ciud = Item.SubItems(5)
    Tele = Item.SubItems(6)
    Come = Item.SubItems(7)
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_COTIZACION, ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM COTIZACION_DETALLE WHERE ID_COTIZACION = " & Item
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        ListView2.ListItems.Clear
        tRs.MoveFirst
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_COTIZACION") & "")
            tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO") & ""
            tLi.SubItems(2) = tRs.Fields("DESCRIPCION") & ""
            tLi.SubItems(3) = tRs.Fields("CANTIDAD") & ""
            tRs.MoveNext
        Loop
    End If
    Me.Command5.Enabled = False
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    ListView3.ListItems.Clear
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tRs1 As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, PRECIO FROM COTIZACION_PROV WHERE ID_PRODUCTO = '" & Item.SubItems(1) & "' AND DESCRIPCION = '" & Item.SubItems(2) & "' AND ID_COTIZACION = '" & Item & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        ListView3.ListItems.Clear
        tRs.MoveFirst
        Do While Not tRs.EOF
            sBuscar = "SELECT NOMBRE FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs.Fields("ID_PROVEEDOR")
            Set tRs1 = cnn.Execute(sBuscar)
            Set tLi = ListView3.ListItems.Add(, , Item & "")
            tLi.SubItems(1) = tRs.Fields("ID_PROVEEDOR") & ""
            tLi.SubItems(2) = tRs1.Fields("NOMBRE") & ""
            tLi.SubItems(3) = tRs.Fields("ID_PRODUCTO") & ""
            tLi.SubItems(4) = tRs.Fields("DESCRIPCION") & ""
            tLi.SubItems(5) = Item.SubItems(3) & ""
            tLi.SubItems(6) = tRs.Fields("PRECIO") & ""
            tRs.MoveNext
        Loop
    Else
        ListView3.ListItems.Clear
    End If
    Me.Command5.Enabled = False
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    it0 = Item
    it1 = Item.SubItems(1)
    it2 = Item.SubItems(2)
    it3 = Item.SubItems(3)
    it4 = Item.SubItems(4)
    it5 = Item.SubItems(5)
    it6 = Item.SubItems(6)
    Me.Command1.Enabled = True
    Me.Command5.Enabled = False
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub ListView4_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    INDI = Item.Index
    Me.Command5.Enabled = True
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Operacion()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tRs2 As Recordset
    Dim tRs3 As Recordset
    Dim Descu As String
    Dim chec As Integer
    Dim Canti As String
    Dim Presc As String
    chec = 0
    Dim NumeroRegistros As Integer
    NumeroRegistros = ListView4.ListItems.Count
    Dim Conta As Integer
    For Conta = 1 To NumeroRegistros
        sBuscar = "SELECT ID_CLIENTE, NOMBRE, DIRECCION, COLONIA, CIUDAD, TELEFONO, COMENTARIOS FROM COTIZACION WHERE ID_COTIZACION = " & ListView4.ListItems(Conta)
        Set tRs = cnn.Execute(sBuscar)
        If tRs.Fields("ID_CLIENTE") <> 0 Then
            sBuscar = "SELECT DESCUENTO FROM CLIENTE WHERE ID_CLIENTE = " & tRs.Fields("ID_CLIENTE")
            Set tRs2 = cnn.Execute(sBuscar)
            If Not IsNull(tRs2.Fields("DESCUENTO")) Then
                Descu = tRs2.Fields("DESCUENTO")
            Else
                Descu = "0"
            End If
            Nomb = tRs.Fields("NOMBRE")
            Dire = tRs.Fields("DIRECCION")
            Colo = tRs.Fields("COLONIA")
            Ciud = tRs.Fields("CIUDAD")
            Tele = tRs.Fields("TELEFONO")
            Come = tRs.Fields("COMENTARIOS")
        Else
            Descu = "0"
            Come = tRs.Fields("COMENTARIOS")
        End If
        If chec = 0 Then
            sBuscar = "INSERT INTO COTIZACION_FINAL (FECHA, NOMBRE_CLIENTE, DIRECCION, COLONIA, CIUDAD, TELEFONO, DESCUENTO, COMENTARIOS) VALUES ('" & Date & "', '" & Nomb & "', '" & Dire & "', '" & Colo & "', '" & Ciud & "', '" & Tele & "', " & Descu & ", '" & Come & "');"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT ID_COTIZA FROM COTIZACION_FINAL ORDER BY ID_COTIZA DESC"
            Set tRs3 = cnn.Execute(sBuscar)
            chec = 1
            sBuscar = "UPDATE COTIZACION SET PENDIENTE = 'N' WHERE ID_COTIZACION = " & ListView4.ListItems(Conta)
            Set tRs = cnn.Execute(sBuscar)
        End If
        Canti = Replace(ListView4.ListItems(Conta).SubItems(5), ",", ".")
        Presc = Replace(ListView4.ListItems(Conta).SubItems(6), ",", ".")
        sBuscar = "INSERT INTO COTIZACION_FINAL_DETALLE (ID_COTIZA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO) VALUES (" & tRs3.Fields("ID_COTIZA") & ", '" & ListView4.ListItems(Conta).SubItems(3) & "', '" & ListView4.ListItems(Conta).SubItems(4) & "', " & Canti & ", " & Presc & ");"
        cnn.Execute (sBuscar)
    Next Conta
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub

Sub Reconexion()
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Const sPathBase As String = "LINUX"
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
        MsgBox "LA CONEXION SE RESABLECIO CON EXITO. PUEDE CONTINUAR CON SU TRABAJO.", vbInformation, "MENSAJE DEL SISTEMA"
    End With
Exit Sub
ManejaError:
    MsgBox Err.Number & Err.Description
    Err.Clear
    If MsgBox("NO PUDIMOS RESTABLECER LA CONEXIÓN, ¿DESEA REINTENTARLO?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
End Sub


