VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form CorteCaja 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CORTE DE CAJA"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   10350
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9240
      TabIndex        =   19
      Top             =   4200
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
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Facturar.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Facturar.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9240
      TabIndex        =   17
      Top             =   3000
      Width           =   975
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   780
         Left            =   120
         MouseIcon       =   "Facturar.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "Facturar.frx":26F6
         Top             =   120
         Width           =   705
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Corte"
      TabPicture(0)   =   "Facturar.frx":4478
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.ListView ListView2 
         Height          =   3255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   8775
         _ExtentX        =   15478
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
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ver Corte"
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
      Picture         =   "Facturar.frx":4494
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   5160
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   21299201
      CurrentDate     =   38701
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ciudad"
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tarjeta"
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contado"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cheques"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chihuahua"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
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
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sucursal"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Agente"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha"
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu MenTipoddeCortes 
      Caption         =   "Tipos de Cortes"
      Begin VB.Menu SubMenCortedeCredito 
         Caption         =   "Corte de Credito"
      End
      Begin VB.Menu SubMenCorteporPeriodos 
         Caption         =   "Corte por Periodos"
      End
   End
End
Attribute VB_Name = "CorteCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Attribute cnn.VB_VarHelpID = -1
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim FechImp As String
Private Sub Command1_Click()
    CorteHoy
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    CorteCaja.Caption = "Corte de Caja Sucursal " & VarMen.Text4(0).Text
    Label5.Caption = VarMen.Text1(1).Text
    Label7.Caption = VarMen.Text4(0).Text
    Label9.Caption = VarMen.Text4(3).Text
    DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .CheckBoxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. Venta", 1000
        .ColumnHeaders.Add , , "Factura", 1000, lvwColumnCenter
        .ColumnHeaders.Add , , "Cliente", 6000, lvwColumnCenter
        .ColumnHeaders.Add , , "Total", 1000, lvwColumnCenter
        .ColumnHeaders.Add , , "Tipo", 1000, lvwColumnCenter
    End With
    CorteHoy
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub

Sub Reconexion()
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
        MsgBox "LA CONEXION SE RESABLECIO CON EXITO. PUEDE CONTINUAR CON SU TRABAJO.", vbInformation, "SACC"
    End With
Exit Sub
ManejaError:
    MsgBox Err.Number & Err.Description
    Err.Clear
    If MsgBox("NO PUDIMOS RESTABLECER LA CONEXIÓN, ¿DESEA REINTENTARLO?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
End Sub
Private Sub CorteHoy()
    Dim tRs As Recordset
    Dim tLi As ListItem
    Dim sBuscar As String
    Dim TotCont As String
    Dim TotChek As String
    Dim TotTarj As String
    FechImp = Format(DTPicker1.Value, "dd/mm/yyyy")
    TotCont = "0.00"
    TotChek = "0.00"
    TotTarj = "0.00"
    sBuscar = "SELECT ID_VENTA, TIPO_PAGO, NOMBRE, TOTAL, FOLIO FROM VENTAS WHERE SUCURSAL = '" & Label7.Caption & "' AND FECHA = '" & DTPicker1.Value & "' AND FACTURADO = 0"
    Set tRs = cnn.Execute(sBuscar)
    ListView2.ListItems.Clear
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_VENTA") & "")
                If Not IsNull(.Fields("FOLIO")) Then tLi.SubItems(1) = .Fields("FOLIO") & ""
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(2) = .Fields("NOMBRE") & ""
                If Not IsNull(.Fields("TOTAL")) Then tLi.SubItems(3) = .Fields("TOTAL")
                If tRs.Fields("TIPO_PAGO") = "C" Then
                    tLi.SubItems(4) = "CONTADO"
                    tLi.ForeColor = &H80000015
                    TotCont = Format(CDbl(TotCont) + CDbl(.Fields("TOTAL")), "0.00")
                Else
                    If .Fields("TIPO_PAGO") = "H" Then
                        tLi.SubItems(4) = "CHEQUE"
                        tLi.ForeColor = &H8000000D
                        TotChek = Format(CDbl(TotChek) + CDbl(.Fields("TOTAL")), "0.00")
                    Else
                        If tRs.Fields("TIPO_PAGO") = "T" Then
                            tLi.SubItems(4) = "TARJETA"
                            tLi.ForeColor = &H80000007
                            TotTarj = Format(CDbl(TotTarj) + CDbl(.Fields("TOTAL")), "0.00")
                        Else
                            tLi.SubItems(4) = "CONTADO"
                            tLi.ForeColor = &H80000015
                            TotCont = Format(CDbl(TotCont) + CDbl(.Fields("TOTAL")), "0.00")
                        End If
                    End If
                End If
                .MoveNext
            Loop
            Text1.ForeColor = &H80000015
            Text2.ForeColor = &H8000000D
            Text3.ForeColor = &H80000007
            Text1.Text = TotCont
            Text2.Text = TotChek
            Text3.Text = TotTarj
        End If
    End With
End Sub

Private Sub Image3_Click()
    On Error GoTo ManejaError
    Printer.Print ""
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
    Printer.Print "         Corte de Caja de Suc. " & Label7.Caption
    Printer.Print "         Ciudad : " & Label9.Caption
    Printer.Print "         Fecha : " & FechImp
    Dim POSY As Integer
    Dim Acum As String
    Acum = "0"
    Dim ACUM1 As String
    ACUM1 = "0"
    Dim ACUM2 As String
    ACUM2 = "0"
    POSY = 2600
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Nota"
    Printer.CurrentY = POSY
    Printer.CurrentX = 1100
    Printer.Print "Factura"
    Printer.CurrentY = POSY
    Printer.CurrentX = 2000
    Printer.Print "Cliente."
    Printer.CurrentY = POSY
    Printer.CurrentX = 7500
    Printer.Print "Efectivo"
    Printer.CurrentY = POSY
    Printer.CurrentX = 8500
    Printer.Print "Cheque"
    Printer.CurrentY = POSY
    Printer.CurrentX = 9500
    Printer.Print "T. Credito"
    POSY = POSY + 200
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Dim NumeroRegistros As Integer
    NumeroRegistros = ListView2.ListItems.Count
    Dim Conta As Integer
    POSY = POSY + 200
    For Conta = 1 To NumeroRegistros
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print ListView2.ListItems.Item(Conta)
        Printer.CurrentY = POSY
        Printer.CurrentX = 1100
        Printer.Print ListView2.ListItems.Item(Conta).SubItems(1)
        Printer.CurrentY = POSY
        Printer.CurrentX = 2000
        Printer.Print ListView2.ListItems.Item(Conta).SubItems(2)
        If ListView2.ListItems.Item(Conta).SubItems(4) = "CONTADO" Or ListView2.ListItems.Item(Conta).SubItems(4) = "" Then
            Printer.CurrentY = POSY
            Printer.CurrentX = 7500
            Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
            Acum = CDbl(Format(Acum, "0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "0.00"))
        End If
        If ListView2.ListItems.Item(Conta).SubItems(4) = "CHEQUE" Then
            Printer.CurrentY = POSY
            Printer.CurrentX = 8500
            Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
            ACUM1 = CDbl(Format(ACUM1, "0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "0.00"))
        End If
        If ListView2.ListItems.Item(Conta).SubItems(4) = "TARJETA" Then
            Printer.CurrentY = POSY
            Printer.CurrentX = 9500
            Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
            ACUM2 = CDbl(Format(ACUM2, "0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "0.00"))
        End If
        POSY = POSY + 200
    Next Conta
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    POSY = POSY + 400
    Printer.CurrentY = POSY
    Printer.CurrentX = 5000
    Printer.Print "TOTAL : "
    Printer.CurrentY = POSY
    Printer.CurrentX = 7500
    Printer.Print Acum
    Printer.CurrentY = POSY
    Printer.CurrentX = 8500
    Printer.Print ACUM1
    Printer.CurrentY = POSY
    Printer.CurrentX = 9500
    Printer.Print ACUM2
    Printer.EndDoc
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub

Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub SubMenCortedeCredito_Click()
    FrmCorteCredito.Show vbModal
End Sub
Private Sub SubMenCorteporPeriodos_Click()
    FrmCorteSemana.Show vbModal
End Sub
