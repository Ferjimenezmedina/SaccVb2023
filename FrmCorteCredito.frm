VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCorteCredito 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Corte de Ventas de Credito"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12330
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11280
      TabIndex        =   24
      Top             =   4680
      Width           =   975
      Begin VB.Image Image2 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FrmCorteCredito.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCorteCredito.frx":030A
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label15 
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
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10320
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11280
      TabIndex        =   22
      Top             =   3480
      Width           =   975
      Begin VB.Image Image1 
         Height          =   825
         Left            =   120
         MouseIcon       =   "FrmCorteCredito.frx":1EDC
         MousePointer    =   99  'Custom
         Picture         =   "FrmCorteCredito.frx":21E6
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo"
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
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11280
      TabIndex        =   20
      Top             =   5880
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
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCorteCredito.frx":43AC
         MousePointer    =   99  'Custom
         Picture         =   "FrmCorteCredito.frx":46B6
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   16
      Top             =   480
      Width           =   2175
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
      Left            =   8640
      Picture         =   "FrmCorteCredito.frx":6798
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Corte Credito"
      TabPicture(0)   =   "FrmCorteCredito.frx":916A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Corte Contado"
      TabPicture(1)   =   "FrmCorteCredito.frx":9186
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text8"
      Tab(1).Control(1)=   "Text7"
      Tab(1).Control(2)=   "Text6"
      Tab(1).Control(3)=   "Check1"
      Tab(1).Control(4)=   "Text1"
      Tab(1).Control(5)=   "Text2"
      Tab(1).Control(6)=   "Text3"
      Tab(1).Control(7)=   "ListView2"
      Tab(1).Control(8)=   "Label18"
      Tab(1).Control(9)=   "Label17"
      Tab(1).Control(10)=   "Label16"
      Tab(1).Control(11)=   "Label1"
      Tab(1).Control(12)=   "Label2"
      Tab(1).Control(13)=   "Label10"
      Tab(1).ControlCount=   14
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -65040
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   5640
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -68640
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   5640
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -66480
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   5640
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Seleccionar Todo"
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   5280
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9600
         TabIndex        =   18
         Top             =   5520
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4695
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8281
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -74040
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   5640
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   5640
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -70440
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   5640
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   7
         Top             =   600
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8070
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label18 
         Caption         =   "NA"
         Height          =   255
         Left            =   -65400
         TabIndex        =   32
         Top             =   5640
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "T. Debito"
         Height          =   255
         Left            =   -69360
         TabIndex        =   30
         Top             =   5640
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "T. Electrónica"
         Height          =   255
         Left            =   -67560
         TabIndex        =   28
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "TOTAL :"
         Height          =   255
         Left            =   8880
         TabIndex        =   19
         Top             =   5520
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cheques"
         Height          =   255
         Left            =   -72960
         TabIndex        =   6
         Top             =   5640
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Contado"
         Height          =   255
         Left            =   -74760
         TabIndex        =   5
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "T. Credito"
         Height          =   255
         Left            =   -71160
         TabIndex        =   4
         Top             =   5640
         Width           =   735
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16252929
      CurrentDate     =   38701
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha"
      Height          =   255
      Left            =   6600
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Agente"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   615
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
      TabIndex        =   13
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sucursal"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   600
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
      Left            =   7320
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ciudad"
      Height          =   255
      Left            =   6600
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "FrmCorteCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim FechImp As String
Dim PresShift As Boolean
Dim sFechaCorte As String
Private Sub Check1_Click()
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBuscar As String
    Dim Cont As Integer
    Dim NoRe As Integer
    NoRe = Me.ListView2.ListItems.Count
    For Cont = 1 To NoRe
        ListView2.ListItems.Item(Cont).Checked = True
    Next Cont
End Sub
Private Sub Combo1_DropDown()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Combo1.Clear
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command1_Click()
    CorteHoy
    CorteCredito
    sFechaCorte = DTPicker1.Value
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 16 Then
        PresShift = True
    Else
        PresShift = False
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    PresShift = False
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    PresShift = False
    Label5.Caption = VarMen.Text1(1).Text
    Label9.Caption = VarMen.Text4(3).Text
    Combo1.Text = VarMen.Text4(0).Text
    DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    sFechaCorte = Format(Date, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
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
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. Venta", 1000
        .ColumnHeaders.Add , , "Factura", 1000, lvwColumnCenter
        .ColumnHeaders.Add , , "Cliente", 6000, lvwColumnCenter
        .ColumnHeaders.Add , , "Total", 1000, lvwColumnCenter
        .ColumnHeaders.Add , , "Fecha de Vencimiento", 1000, lvwColumnCenter
    End With
    CorteHoy
    CorteCredito
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub CorteHoy()
    If Combo1.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        Dim sBuscar As String
        Dim TotCont As String
        Dim TotChek As String
        Dim TotTarj As String
        Dim TraElec As String
        Dim TotTarD As String
        Dim TotNa As String
        FechImp = Format(DTPicker1.Value, "dd/mm/yyyy")
        TotCont = "0.00"
        TotChek = "0.00"
        TotTarj = "0.00"
        TraElec = "0.00"
        TotTarD = "0.00"
        TotNa = "0.00"
        sBuscar = "SELECT ID_VENTA, TIPO_PAGO, NOMBRE, TOTAL, FOLIO FROM VENTAS WHERE SUCURSAL = '" & Combo1.Text & "' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker1.Value & " 23:59:59.997' AND UNA_EXIBICION = 'S' AND FACTURADO < 2"
        Set tRs = cnn.Execute(sBuscar)
        ListView2.ListItems.Clear
        With tRs
            If Not (.EOF And .BOF) Then
                Do While Not .EOF
                    Set tLi = ListView2.ListItems.Add(, , .Fields("ID_VENTA") & "")
                    If Not IsNull(.Fields("FOLIO")) Then tLi.SubItems(1) = .Fields("FOLIO") & ""
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(2) = .Fields("NOMBRE") & ""
                    If Not IsNull(.Fields("TOTAL")) Then tLi.SubItems(3) = CDbl(.Fields("TOTAL"))
                    If tRs.Fields("TIPO_PAGO") = "C" Then
                        tLi.SubItems(4) = "CONTADO"
                        tLi.ForeColor = &H80000015
                        TotCont = Format(Val(Replace(TotCont, ",", "")) + CDbl(.Fields("TOTAL")), "###,###,##0.00")
                    Else
                        If .Fields("TIPO_PAGO") = "H" Then
                            tLi.SubItems(4) = "CHEQUE"
                            tLi.ForeColor = &H8000000D
                            TotChek = Format(Val(Replace(TotChek, ",", "")) + CDbl(.Fields("TOTAL")), "###,###,##0.00")
                        Else
                            If tRs.Fields("TIPO_PAGO") = "T" Then
                                tLi.SubItems(4) = "TARJETA DE CREDTO"
                                tLi.ForeColor = &H80000007
                                TotTarj = Format(Val(Replace(TotTarj, ",", "")) + CDbl(.Fields("TOTAL")), "###,###,##0.00")
                            Else
                                If tRs.Fields("TIPO_PAGO") = "E" Then
                                   tLi.SubItems(4) = "TRA. ELECTRONICA"
                                   tLi.ForeColor = &H80000003
                                   TraElec = Format(Val(Replace(TraElec, ",", "")) + CDbl(.Fields("TOTAL")), "###,###,##0.00")
                                Else
                                    If tRs.Fields("TIPO_PAGO") = "D" Then
                                        tLi.SubItems(4) = "TARJETA DE DEBITO"
                                        tLi.ForeColor = &H80000007
                                        TotTarD = Format(Val(Replace(TotTarD, ",", "")) + CDbl(.Fields("TOTAL")), "###,###,##0.00")
                                    Else
                                        If tRs.Fields("TIPO_PAGO") = "N" Then
                                            tLi.SubItems(4) = "NO APLICA"
                                            tLi.ForeColor = &H80000007
                                            TotNa = Format(Val(Replace(TotNa, ",", "")) + CDbl(.Fields("TOTAL")), "###,###,##0.00")
                                        Else
                                            tLi.SubItems(4) = "CONTADO"
                                            tLi.ForeColor = &H80000015
                                            TotCont = Format(Val(Replace(TotCont, ",", "")) + CDbl(.Fields("TOTAL")), "###,###,##0.00")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    .MoveNext
                Loop
                Text1.ForeColor = &H80000015
                Text2.ForeColor = &H8000000D
                Text3.ForeColor = &H80000007
                Text6.ForeColor = &H80000003
                Text7.ForeColor = &H80000007
                Text1.Text = TotCont
                Text2.Text = TotChek
                Text3.Text = TotTarj
                Text6.Text = TraElec
                Text7.Text = TotTarD
                Text8.Text = TotNa
            End If
        End With
    End If
End Sub
Private Sub Image1_Click()
    FrmCorteSemana.Show vbModal
End Sub
Private Sub Image2_Click()
    Command1.Value = True
    ImpCorte
    ImpCredito
    ImpFacturadoHoy
    'If PresShift = True Then
    ImpCorteNoFacturado
    'End If
End Sub
Private Sub CorteCredito()
    If Combo1.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        Dim sBuscar As String
        Dim TotCont As String
        Dim TotChek As String
        Dim TotTarj As String
        FechImp = Format(DTPicker1.Value, "dd/mm/yyyy")
        TotCont = "0.00"
        TotChek = "0.00"
        TotTarj = "0.00"
        sBuscar = "SELECT ID_VENTA, FECHA_VENCE, NOMBRE, TOTAL, FOLIO FROM VENTAS WHERE SUCURSAL = '" & Combo1.Text & "' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker1.Value & " 23:59:59.997' AND UNA_EXIBICION = 'N' AND FACTURADO < 2"
        Set tRs = cnn.Execute(sBuscar)
        ListView1.ListItems.Clear
        With tRs
            If Not (.EOF And .BOF) Then
                Do While Not .EOF
                    Set tLi = ListView1.ListItems.Add(, , .Fields("ID_VENTA") & "")
                    If Not IsNull(.Fields("FOLIO")) Then tLi.SubItems(1) = .Fields("FOLIO") & ""
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(2) = .Fields("NOMBRE") & ""
                    If Not IsNull(.Fields("TOTAL")) Then tLi.SubItems(3) = CDbl(.Fields("TOTAL"))
                    If Not IsNull(.Fields("FECHA_VENCE")) Then tLi.SubItems(4) = .Fields("FECHA_VENCE")
                    TotCont = Format(Val(Replace(TotCont, ",", "")) + CDbl(.Fields("TOTAL")), "###,###,##0.00")
                    .MoveNext
                Loop
                Text4.Text = TotCont
            End If
        End With
    End If
End Sub
Private Sub ImpCorte()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
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
    Printer.Print "         Corte de Caja de Suc. " & Combo1.Text
    Printer.Print "         Ciudad : " & Label9.Caption
    Printer.Print "         Fecha : " & FechImp
    Printer.Print "-----------------------------------------------------------------------CORTE  DE CONTADO------------------------------------------------------------------------------------------------------------------------------------"
    If Format(sFechaCorte, "dd/mm/yyyy") = Format(Date, "dd/mm/yyyy") Then
        Printer.Print "---------------------------------------------------------CORTE  NO VALIDO POR NO SER AL CIERRE DEL DIA-----------------------------------------------------------------------------------------------------------------------"
    Else
        sBuscar = "SELECT ( DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE()))) AS FECHA"
        Set tRs = cnn.Execute(sBuscar)
        If Format(tRs.Fields("FECHA"), "dd/mm/yyyy") = Format(sFechaCorte, "dd/mm/yyyy") Then
            Printer.Print "---------------------------------------------------------CORTE  NO VALIDO POR NO SER AL CIERRE DEL DIA-----------------------------------------------------------------------------------------------------------------------"
        End If
    End If
    Dim POSY As Integer
    Dim Acum As String
    Acum = "0"
    Dim ACUM1 As String
    ACUM1 = "0"
    Dim ACUM2 As String
    ACUM2 = "0"
    Dim ACUM3 As String
    ACUM3 = "0"
    Dim ACUM4 As String
    ACUM4 = "0"
    Dim ACUM5 As String
    ACUM5 = "0"
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
    Printer.CurrentX = 7300
    Printer.Print "Efectivo"
    Printer.CurrentY = POSY
    Printer.CurrentX = 8100
    Printer.Print "Cheque"
    Printer.CurrentY = POSY
    Printer.CurrentX = 8900
    Printer.Print "T. Credito"
    Printer.CurrentY = POSY
    Printer.CurrentX = 9700
    Printer.Print "T. Debito"
    Printer.CurrentY = POSY
    Printer.CurrentX = 10500
    Printer.Print "Tra. Elec."
    POSY = POSY + 200
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Dim NumeroRegistros As Integer
    NumeroRegistros = ListView2.ListItems.Count
    Dim Conta As Integer
    POSY = POSY + 200
    For Conta = 1 To NumeroRegistros
        If ListView2.ListItems.Item(Conta).SubItems(1) <> "" Then
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print ListView2.ListItems.Item(Conta)
            Printer.CurrentY = POSY
            Printer.CurrentX = 1100
            Printer.Print Mid(ListView2.ListItems.Item(Conta).SubItems(1), 1, 60)
            Printer.CurrentY = POSY
            Printer.CurrentX = 2000
            Printer.Print ListView2.ListItems.Item(Conta).SubItems(2)
            If ListView2.ListItems.Item(Conta).SubItems(4) = "CONTADO" Or ListView2.ListItems.Item(Conta).SubItems(4) = "" Then
                Printer.CurrentY = POSY
                Printer.CurrentX = 7300
                Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
                Acum = CDbl(Format(Acum, "###,###,##0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "###,###,##0.00"))
            End If
            If ListView2.ListItems.Item(Conta).SubItems(4) = "CHEQUE" Then
                Printer.CurrentY = POSY
                Printer.CurrentX = 8100
                Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
                ACUM1 = CDbl(Format(ACUM1, "###,###,##0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "###,###,##0.00"))
            End If
            If ListView2.ListItems.Item(Conta).SubItems(4) = "TARJETA DE CREDTO" Then
                Printer.CurrentY = POSY
                Printer.CurrentX = 8900
                Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
                ACUM2 = CDbl(Format(ACUM2, "###,###,##0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "###,###,##0.00"))
            End If
            If ListView2.ListItems.Item(Conta).SubItems(4) = "TARJETA DE DEBITO" Then
                Printer.CurrentY = POSY
                Printer.CurrentX = 9700
                Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
                ACUM3 = CDbl(Format(ACUM3, "###,###,##0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "###,###,##0.00"))
            End If
            If ListView2.ListItems.Item(Conta).SubItems(4) = "TRA. ELECTRONICA" Then
                Printer.CurrentY = POSY
                Printer.CurrentX = 10500
                Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
                ACUM4 = CDbl(Format(ACUM4, "###,###,##0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "###,###,##0.00"))
            End If
            If ListView2.ListItems.Item(Conta).SubItems(4) = "NO APLICA" Then
                Printer.CurrentY = POSY
                Printer.CurrentX = 10500
                Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
                ACUM5 = CDbl(Format(ACUM5, "###,###,##0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "###,###,##0.00"))
            End If
            POSY = POSY + 200
            If POSY >= 14200 Then
                Printer.NewPage
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
                Printer.Print "         Corte de Caja de Suc. " & Combo1.Text
                Printer.Print "         Ciudad : " & Label9.Caption
                Printer.Print "         Fecha : " & FechImp
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
                Printer.CurrentX = 7300
                Printer.Print "Efectivo"
                Printer.CurrentY = POSY
                Printer.CurrentX = 8100
                Printer.Print "Cheque"
                Printer.CurrentY = POSY
                Printer.CurrentX = 8900
                Printer.Print "T. Credito"
                Printer.CurrentY = POSY
                Printer.CurrentX = 9700
                Printer.Print "T. Debito"
                Printer.CurrentY = POSY
                Printer.CurrentX = 10500
                Printer.Print "Tra. Elec."
                POSY = POSY + 200
                Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            End If
        End If
    Next Conta
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    POSY = POSY + 400
    Printer.CurrentY = POSY
    Printer.CurrentX = 5000
    Printer.Print "TOTAL : "
    Printer.CurrentY = POSY
    Printer.CurrentX = 7300
    Printer.Print Acum
    Printer.CurrentY = POSY
    Printer.CurrentX = 8100
    Printer.Print ACUM1
    Printer.CurrentY = POSY
    Printer.CurrentX = 8900
    Printer.Print ACUM2
    Printer.CurrentY = POSY
    Printer.CurrentX = 9700
    Printer.Print ACUM3
    Printer.CurrentY = POSY
    Printer.CurrentX = 10500
    Printer.Print ACUM4
    Printer.CurrentY = POSY
    Printer.CurrentX = 10500
    Printer.Print ACUM5
    Printer.Print "                                                                                    __________________________                                                                                                                    "
    Printer.Print "                                                                                      FIRMA DEL RESPONSABLE                                                                                                                 "
    Printer.EndDoc
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ImpCredito()
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
    Printer.Print "         Corte de Creditos de Suc. " & Combo1.Text
    Printer.Print "         Ciudad : " & Label9.Caption
    Printer.Print "         Fecha : " & FechImp
    Printer.Print "-----------------------------------------------------------------------CORTE  DE CREDITO------------------------------------------------------------------------------------------------------------------------------------"
    If sFechaCorte = Date Then
        Printer.Print "---------------------------------------------------------CORTE  NO VALIDO POR NO SER AL CIERRE DEL DIA-----------------------------------------------------------------------------------------------------------------------"
    End If
    Dim POSY As Integer
    Dim Acum As String
    Acum = "0"
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
    Printer.CurrentX = 8200
    Printer.Print "Vence"
    Printer.CurrentY = POSY
    Printer.CurrentX = 9500
    Printer.Print "Importe"
    POSY = POSY + 200
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Dim NumeroRegistros As Integer
    NumeroRegistros = ListView1.ListItems.Count
    Dim Conta As Integer
    POSY = POSY + 200
    For Conta = 1 To NumeroRegistros
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print ListView1.ListItems.Item(Conta)
        Printer.CurrentY = POSY
        Printer.CurrentX = 1100
        Printer.Print ListView1.ListItems.Item(Conta).SubItems(1)
        Printer.CurrentY = POSY
        Printer.CurrentX = 2000
        Printer.Print Mid(ListView1.ListItems.Item(Conta).SubItems(2), 1, 60)
        Printer.CurrentY = POSY
        Printer.CurrentX = 8200
        Printer.Print Format(ListView1.ListItems.Item(Conta).SubItems(4), "###,###,##0.00")
        Printer.CurrentY = POSY
        Printer.CurrentX = 9500
        Printer.Print ListView1.ListItems.Item(Conta).SubItems(3)
        Acum = CDbl(Format(Acum, "###,###,##0.00")) + CDbl(Format(ListView1.ListItems.Item(Conta).SubItems(3), "###,###,##0.00"))
        POSY = POSY + 200
        If POSY >= 14200 Then
            Printer.NewPage
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
            Printer.Print "         Corte de Creditos de Suc. " & Combo1.Text
            Printer.Print "         Ciudad : " & Label9.Caption
            Printer.Print "         Fecha : " & FechImp
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
            Printer.CurrentX = 8200
            Printer.Print "Vence"
            Printer.CurrentY = POSY
            Printer.CurrentX = 9500
            Printer.Print "Importe"
            POSY = POSY + 200
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        End If
    Next Conta
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    POSY = POSY + 400
    Printer.CurrentY = POSY
    Printer.CurrentX = 8300
    Printer.Print "TOTAL : $ "
    Printer.CurrentY = POSY
    Printer.CurrentX = 9500
    Printer.Print Format(Acum, "###,###,##0.00")
    Printer.Print "                                                                                    ___________________________                                                                                                                  "
    Printer.Print "                                                                                      FIRMA DEL RESPONSABLE                                                                                                                 "
    Printer.EndDoc
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ImpCorteNoFacturado()
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
    Printer.Print "         Corte de Caja de Suc. " & Combo1.Text
    Printer.Print "         Ciudad : " & Label9.Caption
    Printer.Print "         Fecha : " & FechImp
    Printer.Print "-----------------------------------------------------------------------CORTE  DE NOTAS NO FACTURADAS------------------------------------------------------------------------------------------------------------------------------------"
    If sFechaCorte = Date Then
        Printer.Print "---------------------------------------------------------CORTE  NO VALIDO POR NO SER AL CIERRE DEL DIA-----------------------------------------------------------------------------------------------------------------------"
    End If
    Dim POSY As Integer
    Dim Acum As String
    Acum = "0"
    Dim ACUM1 As String
    ACUM1 = "0"
    Dim ACUM2 As String
    ACUM2 = "0"
    Dim ACUM3 As String
    ACUM3 = "0"
    Dim ACUM4 As String
    ACUM4 = "0"
    Dim ACUM5 As String
    ACUM5 = "0"
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
    Printer.CurrentX = 7300
    Printer.Print "Efectivo"
    Printer.CurrentY = POSY
    Printer.CurrentX = 8100
    Printer.Print "Cheque"
    Printer.CurrentY = POSY
    Printer.CurrentX = 8900
    Printer.Print "T. Credito"
    Printer.CurrentY = POSY
    Printer.CurrentX = 9700
    Printer.Print "T. Debito"
    Printer.CurrentY = POSY
    Printer.CurrentX = 10500
    Printer.Print "Tra. Elec."
    POSY = POSY + 200
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Dim NumeroRegistros As Integer
    NumeroRegistros = ListView2.ListItems.Count
    Dim Conta As Integer
    POSY = POSY + 200
    For Conta = 1 To NumeroRegistros
        If ListView2.ListItems.Item(Conta).SubItems(1) = "" Then
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print ListView2.ListItems.Item(Conta)
            Printer.CurrentY = POSY
            Printer.CurrentX = 1100
            Printer.Print Mid(ListView2.ListItems.Item(Conta).SubItems(1), 1, 60)
            Printer.CurrentY = POSY
            Printer.CurrentX = 2000
            Printer.Print ListView2.ListItems.Item(Conta).SubItems(2)
            If ListView2.ListItems.Item(Conta).SubItems(4) = "CONTADO" Or ListView2.ListItems.Item(Conta).SubItems(4) = "" Then
                Printer.CurrentY = POSY
                Printer.CurrentX = 7300
                Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
                Acum = CDbl(Format(Acum, "###,###,##0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "###,###,##0.00"))
            End If
            If ListView2.ListItems.Item(Conta).SubItems(4) = "CHEQUE" Then
                Printer.CurrentY = POSY
                Printer.CurrentX = 8100
                Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
                ACUM1 = CDbl(Format(ACUM1, "###,###,##0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "###,###,##0.00"))
            End If
            If ListView2.ListItems.Item(Conta).SubItems(4) = "TARJETA DE CREDTO" Then
                Printer.CurrentY = POSY
                Printer.CurrentX = 8900
                Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
                ACUM2 = CDbl(Format(ACUM2, "###,###,##0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "###,###,##0.00"))
            End If
            If ListView2.ListItems.Item(Conta).SubItems(4) = "TARJETA DE DEBITO" Then
                Printer.CurrentY = POSY
                Printer.CurrentX = 9700
                Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
                ACUM3 = CDbl(Format(ACUM3, "###,###,##0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "###,###,##0.00"))
            End If
            'TRA. ELECTRONICA
            If ListView2.ListItems.Item(Conta).SubItems(4) = "TRA. ELECTRONICA" Then
                Printer.CurrentY = POSY
                Printer.CurrentX = 10500
                Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
                ACUM4 = CDbl(Format(ACUM4, "###,###,##0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "###,###,##0.00"))
            End If
            If ListView2.ListItems.Item(Conta).SubItems(4) = "NO APLICA" Then
                Printer.CurrentY = POSY
                Printer.CurrentX = 10500
                Printer.Print ListView2.ListItems.Item(Conta).SubItems(3)
                ACUM5 = CDbl(Format(ACUM5, "###,###,##0.00")) + CDbl(Format(ListView2.ListItems.Item(Conta).SubItems(3), "###,###,##0.00"))
            End If
            POSY = POSY + 200
            If POSY >= 14200 Then
                Printer.NewPage
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
                Printer.Print "         Corte de Caja de Suc. " & Combo1.Text
                Printer.Print "         Ciudad : " & Label9.Caption
                Printer.Print "         Fecha : " & FechImp
                POSY = 2600
                Printer.CurrentY = POSY
                Printer.CurrentX = 1100
                Printer.Print "Factura"
                Printer.CurrentY = POSY
                Printer.CurrentX = 2000
                Printer.Print "Cliente."
                Printer.CurrentY = POSY
                Printer.CurrentX = 7300
                Printer.Print "Efectivo"
                Printer.CurrentY = POSY
                Printer.CurrentX = 8100
                Printer.Print "Cheque"
                Printer.CurrentY = POSY
                Printer.CurrentX = 8900
                Printer.Print "T. Credito"
                Printer.CurrentY = POSY
                Printer.CurrentX = 9700
                Printer.Print "T. Debito"
                Printer.CurrentY = POSY
                Printer.CurrentX = 10500
                Printer.Print "Tra. Elec."
                POSY = POSY + 200
                Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            End If
        End If
    Next Conta
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    POSY = POSY + 400
    Printer.CurrentY = POSY
    Printer.CurrentX = 5000
    Printer.Print "TOTAL : "
    Printer.CurrentY = POSY
    Printer.CurrentX = 7300
    Printer.Print Format(Acum, "###,###,##0.00")
    Printer.CurrentY = POSY
    Printer.CurrentX = 8100
    Printer.Print Format(ACUM1, "###,###,##0.00")
    Printer.CurrentY = POSY
    Printer.CurrentX = 8900
    Printer.Print Format(ACUM2, "###,###,##0.00")
    Printer.CurrentY = POSY
    Printer.CurrentX = 9700
    Printer.Print Format(ACUM3, "###,###,##0.00")
    Printer.CurrentY = POSY
    Printer.CurrentX = 10500
    Printer.Print Format(ACUM4, "###,###,##0.00")
    Printer.CurrentY = POSY
    Printer.CurrentX = 11300
    Printer.Print Format(ACUM5, "###,###,##0.00")
    Printer.EndDoc
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ImpFacturadoHoy()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim POSY As Integer
    Dim VarTotal As String
    VarTotal = "0"
    sBuscar = "SELECT ID_VENTA, FOLIO, NOMBRE, TOTAL, FECHA FROM VENTAS WHERE SUCURSAL= '" & Combo1.Text & "' AND FECHA_FACTURA >= '" & DTPicker1.Value & "' AND FECHA_FACTURA < '" & DTPicker1.Value + 1 & "' AND FACTURADO = '1'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
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
        Printer.Print "         Corte de Caja de Suc. " & Combo1.Text
        Printer.Print "         Ciudad : " & Label9.Caption
        Printer.Print "         Fecha : " & FechImp
        Printer.Print "-----------------------------------------------------------------------CORTE  DE NOTAS FACTURADAS HOY------------------------------------------------------------------------------------------------------------------------------------"
        If sFechaCorte = Date Then
            Printer.Print "---------------------------------------------------------CORTE  NO VALIDO POR NO SER AL CIERRE DEL DIA-----------------------------------------------------------------------------------------------------------------------"
        End If
        POSY = 2800
        Printer.Print "         FACTURADAS EL DIA " & DTPicker1.Value
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
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
        Printer.Print "Total"
        POSY = POSY + 300
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Do While Not (tRs.EOF)
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print tRs.Fields("ID_VENTA")
            Printer.CurrentY = POSY
            Printer.CurrentX = 1100
            Printer.Print tRs.Fields("FOLIO")
            Printer.CurrentY = POSY
            Printer.CurrentX = 2000
            Printer.Print tRs.Fields("NOMBRE")
            Printer.CurrentY = POSY
            Printer.CurrentX = 7500
            Printer.Print Format(tRs.Fields("TOTAL"), "###,###,##0.00")
            VarTotal = CDbl(VarTotal) + tRs.Fields("TOTAL")
            Printer.CurrentY = POSY
            Printer.CurrentX = 8800
            Printer.Print tRs.Fields("FECHA")
            POSY = POSY + 200
            If POSY >= 14200 Then
                Printer.NewPage
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
                Printer.Print "         Corte de Caja de Suc. " & Combo1.Text
                Printer.Print "         Ciudad : " & Label9.Caption
                Printer.Print "         Fecha : " & FechImp
                POSY = 2800
                Printer.Print "         FACTURADAS EL DIA DE HOY"
                Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
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
                Printer.Print "Total"
                POSY = POSY + 300
                Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            End If
            tRs.MoveNext
        Loop
        Printer.CurrentY = POSY
        Printer.CurrentX = 2000
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        POSY = POSY + 200
        Printer.CurrentY = POSY
        Printer.CurrentX = 7500
        Printer.Print VarTotal
        Printer.Print "                                                                                    _________________________                                                                                                                    "
        Printer.Print "                                                                                      FIRMA DEL RESPONSABLE                                                                                                                 "
        Printer.EndDoc
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
