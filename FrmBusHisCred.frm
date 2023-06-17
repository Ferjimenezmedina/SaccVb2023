VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmBusHisCred 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Historial de Credito Cliente"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   8040
      ScaleHeight     =   6315
      ScaleWidth      =   1755
      TabIndex        =   9
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton Command1 
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
         Left            =   120
         Picture         =   "FrmBusHisCred.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdVer 
         Caption         =   "Ver"
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
         Picture         =   "FrmBusHisCred.frx":29D2
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton CmdImprimir 
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
         Picture         =   "FrmBusHisCred.frx":53A4
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton CmdExcel 
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
         Picture         =   "FrmBusHisCred.frx":7D76
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   240
         TabIndex        =   10
         Top             =   4920
         Width           =   975
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmBusHisCred.frx":A748
            MousePointer    =   99  'Custom
            Picture         =   "FrmBusHisCred.frx":AA52
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
            TabIndex        =   11
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   7815
   End
   Begin MSComctlLib.ListView LVDeuda 
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   7815
      _ExtentX        =   13785
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTFechaAl 
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   95354881
      CurrentDate     =   38828
   End
   Begin MSComCtl2.DTPicker DTFechade 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   95354881
      CurrentDate     =   38828
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3625
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   6975
   End
   Begin VB.Label Label2 
      Caption         =   "A la Fecha :"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label De 
      Caption         =   "De la Fecha :"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmBusHisCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim NOM As String
Dim CLVCLIEN As Integer
Dim LimCred As String

Private Sub CmdExcel_Click()
On Error GoTo ManejaError
    Dim FILE As String
    CommonDialog1.DialogTitle = "GUARDAR COMO"
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "[Archivo Excel (*.xls)] |*.xls|"
    FILE = CommonDialog1.FileName
    CommonDialog1.ShowOpen
    Dim ApExcel As Excel.Application
    Set ApExcel = CreateObject("Excel.application")
    ApExcel.Workbooks.Add
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim Y As Integer
    Dim fech As String
    Dim ACUMABONO As Double
    Dim ACUMDEUDA As Double
    Dim LIM As Integer
    ACUMABONO = 0
    fech = DTFechade.Value
    Y = 11
    LIM = 0
    sBuscar = "SELECT CANT_ABONO, FECHA FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & CLVCLIEN & " ORDER BY ID_ABONO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        tRs.MoveFirst
        Do While LIM = 0
            If tRs.Fields("FECHA") = DTFechade.Value Then
                LIM = 1
            Else
                ACUMABONO = CDbl(ACUMABONO) + CDbl(tRs.Fields("CANT_ABONO"))
            End If
            tRs.MoveNext
            If (tRs.BOF Or tRs.EOF) Then
                LIM = 1
            End If
        Loop
    End If
    LIM = 0
    sBuscar = "SELECT TOTAL_COMPRA, FECHA FROM CUENTAS WHERE ID_CLIENTE = " & CLVCLIEN & " ORDER BY ID_CUENTA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
    tRs.MoveFirst
        Do While LIM = 0
            If tRs.Fields("FECHA") = DTFechade.Value Then
                LIM = 1
            Else
                ACUMDEUDA = CDbl(ACUMDEUDA) + CDbl(tRs.Fields("TOTAL_COMPRA"))
            End If
            tRs.MoveNext
            If (tRs.BOF Or tRs.EOF) Then
                LIM = 1
            End If
        Loop
    End If
    ACUMDEUDA = ACUMDEUDA - ACUMABONO
    ApExcel.Cells(1, 1) = "ACTITUD POSITIVA EN TONER S DE RL MI"
    ApExcel.Cells(2, 1) = "R.F.C. APT- 040201-KA5"
    ApExcel.Cells(3, 1) = "ORTIZ DE CAMPOS No. 1308 COL. SAN FELIPE"
    ApExcel.Cells(4, 1) = "CHIHUAHUA, CHIHUAHUA C.P. 31203"
    ApExcel.Cells(5, 1) = "ESTADO DE CUENTA"
    ApExcel.Cells(6, 1) = "FECHA : " & Date
    ApExcel.Cells(7, 1) = "SUCURSAL : " '& Menu.Text4(0).Text
    ApExcel.Cells(8, 1) = "IMPRESO POR : " ' & Menu.Text1(1).Text & " " & Menu.Text1(2).Text
    ApExcel.Cells(9, 1) = "CLIENTE : " & NOM
    ApExcel.Cells(10, 1) = "CONCEPTO"
    ApExcel.Cells(10, 2) = "FECHA"
    ApExcel.Cells(10, 3) = "IMPORTE"
    ApExcel.Cells(10, 4) = "PENDIENTE"
    Do While fech <> (DTFechaAl.Value + 1)
        sBuscar = "SELECT FECHA, CANT_ABONO FROM ABONOS_CUENTA WHERE FECHA = '" & fech & "' AND ID_CLIENTE = " & CLVCLIEN
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.BOF And tRs.EOF) Then
            Do While Not tRs.EOF
                ApExcel.Cells(Y, 1) = "ABONO"
                ApExcel.Cells(Y, 2) = fech
                ApExcel.Cells(Y, 3) = tRs.Fields("CANT_ABONO")
                ACUMDEUDA = Format(CDbl(ACUMDEUDA) - CDbl(tRs.Fields("CANT_ABONO")), "0.00")
                ApExcel.Cells(Y, 4) = CDbl(ACUMDEUDA)
                Y = Y + 1
                tRs.MoveNext
            Loop
        End If
        sBuscar = "SELECT FECHA, TOTAL_COMPRA FROM CUENTAS WHERE FECHA = '" & fech & "' AND ID_CLIENTE = " & CLVCLIEN
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.BOF And tRs.EOF) Then
            Do While Not tRs.EOF
                ApExcel.Cells(Y, 1) = "COMPRA"
                ApExcel.Cells(Y, 2) = fech
                ApExcel.Cells(Y, 3) = tRs.Fields("TOTAL_COMPRA")
                ACUMDEUDA = Format(CDbl(ACUMDEUDA) + CDbl(tRs.Fields("TOTAL_COMPRA")), "0.00")
                ApExcel.Cells(Y, 4) = CDbl(ACUMDEUDA)
                Y = Y + 1
                tRs.MoveNext
            Loop
        End If
    Loop
    Y = Y + 1
    ApExcel.Cells(Y, 4) = "TOTAL DE CREDITO   : $ " & ACUMDEUDA
    Y = Y + 1
    ApExcel.Cells(Y, 4) = "LIMITE DE CREDITO  : $ " & LimCred
    Y = Y + 1
    Dim TOTCREDDIS As Double
    TOTCREDDIS = CDbl(LimCred) - CDbl(ACUMDEUDA)
    Y = Y + 1
    ApExcel.Cells(Y, 4) = "CREDITO DISPONIBLE : $ " & TOTCREDDIS
    fech = DTFechade.Value + 1
    ApExcel.AlertBeforeOverwriting = False
    ApExcel.ActiveWorkbook.SaveAs "" & FILE
    ApExcel.Visible = True
    Set ApExcel = Nothing
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
Private Sub CmdImprimir_Click()
On Error GoTo ManejaError
    CommonDialog1.Flags = 64
    CommonDialog1.ShowPrinter
    ImprimeDeuda
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
Private Sub CmdVer_Click()
On Error GoTo ManejaError
    FrmBusHisCred.Height = 8550
    Dim vuelta As Integer
    Dim tLi As ListItem
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim fech As String
    Dim ACUMABONO As Double
    Dim ACUMDEUDA As Double
    Dim LIM As Integer
    ACUMABONO = 0
    fech = DTFechade.Value
    LIM = 0
    LVDeuda.ListItems.Clear
    sBuscar = "SELECT CANT_ABONO, FECHA FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & CLVCLIEN & " ORDER BY ID_ABONO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        tRs.MoveFirst
        Do While LIM = 0
            If tRs.Fields("FECHA") = DTFechade.Value Then
                LIM = 1
            Else
                ACUMABONO = CDbl(ACUMABONO) + CDbl(tRs.Fields("CANT_ABONO"))
            End If
            tRs.MoveNext
            If (tRs.BOF Or tRs.EOF) Then
                LIM = 1
            End If
        Loop
    End If
    LIM = 0
    sBuscar = "SELECT TOTAL_COMPRA, FECHA FROM CUENTAS WHERE ID_CLIENTE = " & CLVCLIEN & " ORDER BY ID_CUENTA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        tRs.MoveFirst
        Do While LIM = 0
            If tRs.Fields("FECHA") = DTFechade.Value Then
                LIM = 1
            Else
                If tRs.Fields("TOTAL_COMPRA") <> Null Then
                    ACUMDEUDA = CDbl(ACUMDEUDA) + CDbl(tRs.Fields("TOTAL_COMPRA"))
                End If
            End If
            tRs.MoveNext
            If (tRs.BOF Or tRs.EOF) Then
                LIM = 1
            End If
        Loop
    End If
    ACUMDEUDA = ACUMDEUDA - ACUMABONO
    Do While fech <> (DTFechaAl.Value + 1)
        sBuscar = "SELECT FECHA, CANT_ABONO FROM ABONOS_CUENTA WHERE FECHA = '" & fech & "' AND ID_CLIENTE = " & CLVCLIEN
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.BOF And tRs.EOF) Then
            Do While Not tRs.EOF
                Set tLi = LVDeuda.ListItems.Add(, , "ABONO" & "")
                tLi.SubItems(1) = fech & ""
                tLi.SubItems(2) = tRs.Fields("CANT_ABONO") & ""
                ACUMDEUDA = Format(CDbl(ACUMDEUDA) - CDbl(tRs.Fields("CANT_ABONO")), "0.00")
                tLi.SubItems(3) = CDbl(ACUMDEUDA)
                tRs.MoveNext
            Loop
        End If
        sBuscar = "SELECT FECHA, TOTAL_COMPRA FROM CUENTAS WHERE FECHA = '" & fech & "' AND ID_CLIENTE = " & CLVCLIEN
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.BOF And tRs.EOF) Then
            Do While Not tRs.EOF
                Set tLi = LVDeuda.ListItems.Add(, , "COMPRA" & "")
                tLi.SubItems(1) = fech & ""
                If Not IsNull(tRs.Fields("TOTAL_COMPRA")) Then
                    tLi.SubItems(2) = tRs.Fields("TOTAL_COMPRA") & ""
                Else
                    tLi.SubItems(2) = "0.00"
                End If
                If Not IsNull(tRs.Fields("TOTAL_COMPRA")) Then
                    ACUMDEUDA = Format(CDbl(ACUMDEUDA) + CDbl(tRs.Fields("TOTAL_COMPRA")), "0.00")
                End If
                tLi.SubItems(3) = CDbl(ACUMDEUDA)
                tRs.MoveNext
            Loop
        End If
        vuelta = vuelta + 1
        fech = DTFechade.Value + vuelta
    Loop
    Dim TOTCREDDIS As Double
    TOTCREDDIS = CDbl(LimCred) - CDbl(ACUMDEUDA)
    Text2.Text = "TOTAL DE CREDITO   : $ " & ACUMDEUDA & "    LIMITE DE CREDITO  : $ " & LimCred & "    CREDITO DISPONIBLE  : $ " & TOTCREDDIS
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
'Private Sub Form_Load()
'On Error GoTo ManejaError
'    FrmBusHisCred.Height = 3945
'    DTFechade.Value = Date
'    DTFechade.Value = DTFechade.Value - 30
'    DTFechaAl.Value = Date
'    Me.CmdVer.Enabled = False
'    Me.CmdImprimir.Enabled = False
'    Me.CmdExcel.Enabled = False
'    'Me.Command1.Enabled = False
'    Const sPathBase As String = "LINUX"
'    Set cnn = New ADODB.Connection
'    Set rst = New ADODB.Recordset
'    With cnn
'        .ConnectionString = _
 '           "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
 '           "Data Source=" & sPathBase & ";"
 '       .Open
'    End With
    'With ListView1
    '    .View = lvwReport
    '    .Gridlines = True
    '    .LabelEdit = lvwManual
    '    .ColumnHeaders.Add , , "Clave del Cliente", 1800
    '    .ColumnHeaders.Add , , "Nombre", 7450
    '    .ColumnHeaders.Add , , "RFC", 2450
    '    .ColumnHeaders.Add , , "Limite de credito", 2450
    'End With
'    With LVDeuda
'        .View = lvwReport
'        .Gridlines = True
'        .LabelEdit = lvwManual
'        .ColumnHeaders.Add , , "CONCEPTO", 1800
'        .ColumnHeaders.Add , , "FECHA", 2500
'        .ColumnHeaders.Add , , "IMPORTE", 2450
'        .ColumnHeaders.Add , , "PENDIENTE", 2450
'    End With
'End Sub
'Private Sub Buscar()
'On Error GoTo ManejaError
'    Dim sBuscar As String
'    Dim tRs As Recordset
'    Dim tLi As ListItem
'    sBuscar = "SELECT ID_CLIENTE, NOMBRE, RFC, LIMITE_CREDITO FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND LIMITE_CREDITO <> 0"
'    Set tRs = cnn.Execute(sBuscar)
'    With tRs
'        If (.BOF And .EOF) Then
'            Text1.Text = ""
'            MsgBox "No se encontro cliente con credito registrado a ese nombre"
'        Else
'            ListView1.ListItems.Clear
'            .MoveFirst
'            Do While Not .EOF
'                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
'                tLi.SubItems(1) = .Fields("NOMBRE") & ""
'                tLi.SubItems(2) = .Fields("RFC") & ""
'                tLi.SubItems(3) = .Fields("LIMITE_CREDITO") & ""
'                .MoveNext
'            Loop
'        End If
'    End With
'End Sub
'Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo ManejaError
'    Text1.Text = Item.SubItems(1)
'    NOM = Item.SubItems(1)
'    CLVCLIEN = Item
'    LimCred = Item.SubItems(3)
'    Me.CmdVer.Enabled = True
'    Me.CmdImprimir.Enabled = True
'    Me.CmdExcel.Enabled = True
'End Sub
'Private Sub Text1_Change()
'On Error GoTo ManejaError
'    If Text1.Text <> "" Then
'        Me.Command1.Enabled = True
'    Else
'        Me.CmdVer.Enabled = False
'        Me.CmdImprimir.Enabled = False
'        Me.CmdExcel.Enabled = False
'        Me.Command1.Enabled = False
'    End If
'End Sub
'Private Sub Text1_KeyPress(KeyAscii As Integer)
'On Error GoTo ManejaError
'    If Text1.Text <> "" Then
'        If KeyAscii = 13 Then
 '           Buscar
 ''           ListView1.SetFocus
 '       End If
 '   End If
 '   Dim Valido As String
 '   Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
 '   KeyAscii = Asc(UCase(Chr(KeyAscii)))
 '   If KeyAscii > 26 Then
 '       If InStr(Valido, Chr(KeyAscii)) = 0 Then
  '          KeyAscii = 0
  '      End If
  '  End If
'End Sub
Private Sub ImprimeDeuda()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim X As Integer
    Dim Y As Integer
    Dim fech As String
    Dim ACUMABONO As Double
    Dim ACUMDEUDA As Double
    Dim LIM As Integer
    ACUMABONO = 0
    fech = DTFechade.Value
    X = 3000
    Y = 20
    LIM = 0
    sBuscar = "SELECT CANT_ABONO, FECHA FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & CLVCLIEN & " ORDER BY ID_ABONO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        tRs.MoveFirst
        Do While LIM = 0
            If tRs.Fields("FECHA") = DTFechade.Value Then
                LIM = 1
            Else
                ACUMABONO = CDbl(ACUMABONO) + CDbl(tRs.Fields("CANT_ABONO"))
            End If
            tRs.MoveNext
            If (tRs.BOF Or tRs.EOF) Then
                LIM = 1
            End If
        Loop
    End If
    LIM = 0
    sBuscar = "SELECT TOTAL_COMPRA, FECHA FROM CUENTAS WHERE ID_CLIENTE = " & CLVCLIEN & " ORDER BY ID_CUENTA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        tRs.MoveFirst
        Do While LIM = 0
            If tRs.Fields("FECHA") = DTFechade.Value Then
                LIM = 1
            Else
                If tRs.Fields("TOTAL_COMPRA") <> Null Then
                    ACUMDEUDA = CDbl(ACUMDEUDA) + CDbl(tRs.Fields("TOTAL_COMPRA"))
                End If
            End If
            tRs.MoveNext
            If (tRs.BOF Or tRs.EOF) Then
                LIM = 1
            End If
        Loop
    End If
    ACUMDEUDA = ACUMDEUDA - ACUMABONO
    Enca
    Printer.CurrentY = 2800
    Printer.CurrentX = 100
    Printer.Print "CONCEPTO"
    Printer.CurrentY = 2800
    Printer.CurrentX = 3000
    Printer.Print "FECHA"
    Printer.CurrentY = 2800
    Printer.CurrentX = 6000
    Printer.Print "IMPORTE"
    Printer.CurrentY = 2800
    Printer.CurrentX = 9000
    Printer.Print "PENDIENTE"
    Do While fech <> (DTFechaAl.Value + 1)
        sBuscar = "SELECT FECHA, CANT_ABONO FROM ABONOS_CUENTA WHERE FECHA = '" & fech & "' AND ID_CLIENTE = " & CLVCLIEN
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.BOF And tRs.EOF) Then
            Do While Not tRs.EOF
                Printer.CurrentY = X
                Printer.CurrentX = 100
                Printer.Print "ABONO"
                Printer.CurrentY = X
                Printer.CurrentX = 3000
                Printer.Print fech
                Printer.CurrentY = X
                Printer.CurrentX = 6000
                Printer.Print tRs.Fields("CANT_ABONO")
                Printer.CurrentY = X
                Printer.CurrentX = 9000
                ACUMDEUDA = Format(CDbl(ACUMDEUDA) - CDbl(tRs.Fields("CANT_ABONO")), "0.00")
                Printer.Print CDbl(ACUMDEUDA)
                X = X + 200
                Y = Y + 1
                tRs.MoveNext
                If Y = 73 Then
                    Printer.EndDoc
                    Enca
                    X = 200
                    Y = 20
                End If
            Loop
        End If
        sBuscar = "SELECT FECHA, TOTAL_COMPRA FROM CUENTAS WHERE FECHA = '" & fech & "' AND ID_CLIENTE = " & CLVCLIEN
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.BOF And tRs.EOF) Then
            Do While Not tRs.EOF
                Printer.CurrentY = X
                Printer.CurrentX = 100
                Printer.Print "COMPRA"
                Printer.CurrentY = X
                Printer.CurrentX = 3000
                Printer.Print fech
                Printer.CurrentY = X
                Printer.CurrentX = 6000
                Printer.Print tRs.Fields("TOTAL_COMPRA")
                Printer.CurrentY = X
                Printer.CurrentX = 9000
                ACUMDEUDA = Format(CDbl(ACUMDEUDA) + CDbl(tRs.Fields("TOTAL_COMPRA")), "0.00")
                Printer.Print CDbl(ACUMDEUDA)
                X = X + 200
                Y = Y + 1
                tRs.MoveNext
                If Y = 73 Then
                    Printer.EndDoc
                    Enca
                    X = 200
                    Y = 20
                End If
            Loop
        End If
    Loop
    Printer.CurrentY = X
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.CurrentY = X + 200
    Printer.CurrentX = 6500
    Printer.Print "TOTAL DE CREDITO   : $ " & ACUMDEUDA
    Printer.CurrentY = X + 400
    Printer.CurrentX = 6500
    Printer.Print "LIMITE DE CREDITO  : $ " & LimCred
    Printer.CurrentY = X + 600
    Printer.CurrentX = 6500
    Dim TOTCREDDIS As Double
    TOTCREDDIS = CDbl(LimCred) - CDbl(ACUMDEUDA)
    Printer.Print "CREDITO DISPONIBLE : $ " & TOTCREDDIS
    fech = DTFechade.Value + 1
    Printer.EndDoc
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
Private Sub Enca()
On Error GoTo ManejaError
    Printer.Print ""
    Printer.Print ""
    Printer.Print "                                                                                   ACTITUD POSITIVA EN TONER S DE RL MI"
    Printer.Print "                                                                                                R.F.C. APT- 040201-KA5"
    Printer.Print "                                                                           ORTIZ DE CAMPOS No. 1308 COL. SAN FELIPE"
    Printer.Print "                                                                                    CHIHUAHUA, CHIHUAHUA C.P. 31203"
    Printer.Print "             ESTADO DE CUENTA"
    Printer.Print "             FECHA : " & Date
    Printer.Print "             SUCURSAL : " & Menu.Text4(0).Text
    Printer.Print "             IMPRESO POR : " & Menu.Text1(1).Text & " " & Menu.Text1(2).Text
    Printer.Print "             CLIENTE : " & NOM
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "                                                                          COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO"
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.CurrentY = 2800
    Printer.CurrentX = 100
    Printer.Print "CONCEPTO"
    Printer.CurrentY = 2800
    Printer.CurrentX = 3000
    Printer.Print "FECHA"
    Printer.CurrentY = 2800
    Printer.CurrentX = 6000
    Printer.Print "IMPORTE"
    Printer.CurrentY = 2800
    Printer.CurrentX = 9000
    Printer.Print "PENDIENTE"
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
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

Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_LostFocus()
      Text1.BackColor = &H80000005
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_LostFocus()
      Text2.BackColor = &H80000005
End Sub
