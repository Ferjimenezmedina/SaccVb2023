VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFacturaCredito 
   BackColor       =   &H00FFFFFF&
   Caption         =   "FACTURAS"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7200
      TabIndex        =   32
      Top             =   5760
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmFacturaCredito.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmFacturaCredito.frx":030A
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
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmFacturaCredito.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblFP"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblN"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblD"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblT"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblC"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "DTPicker2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "DTPicker1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "CommonDialog1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtCom"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtOC"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtFolio"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdFacturar"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdChange"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtTL"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Frame1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtMH"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtAH"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtDP"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtMP"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtAP"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtDH"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      Begin VB.TextBox txtDH 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   6240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAP 
         Height          =   285
         Left            =   3480
         TabIndex        =   15
         Top             =   6240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtMP 
         Height          =   285
         Left            =   2880
         TabIndex        =   14
         Top             =   6240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDP 
         Height          =   285
         Left            =   2280
         TabIndex        =   13
         Text            =   " "
         Top             =   6240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAH 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   6240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtMH 
         Height          =   285
         Left            =   840
         TabIndex        =   11
         Top             =   6240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   6495
         Begin VB.CommandButton cmdOk 
            Caption         =   "OK"
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
            Left            =   5280
            Picture         =   "frmFacturaCredito.frx":2408
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtVenta 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   4080
            TabIndex        =   0
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "ESCRIBA AQUÍ EL NUMERO DE LA VENTA QUE DESEA FACTURAR"
            Height          =   495
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.TextBox txtTL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3960
         Width           =   6015
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "..."
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
         Left            =   6360
         Picture         =   "frmFacturaCredito.frx":4DDA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton cmdFacturar 
         Caption         =   "Facturar"
         Enabled         =   0   'False
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
         Left            =   5520
         Picture         =   "frmFacturaCredito.frx":77AC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6000
         Width           =   1095
      End
      Begin VB.TextBox txtFolio 
         Height          =   285
         Left            =   4440
         TabIndex        =   5
         Top             =   5640
         Width           =   1815
      End
      Begin VB.TextBox txtOC 
         Height          =   285
         Left            =   4440
         TabIndex        =   4
         Top             =   5280
         Width           =   1815
      End
      Begin VB.TextBox txtCom 
         Height          =   285
         Left            =   4440
         TabIndex        =   3
         Top             =   4920
         Width           =   1815
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4920
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         PrinterDefault  =   0   'False
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   5760
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   21168129
         CurrentDate     =   38728
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   5760
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   21168129
         CurrentDate     =   38728
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "CLIENTE:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "DIAS CREDITO:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "NOMBRE:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label lblC 
         Caption         =   "..."
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
         Left            =   1920
         TabIndex        =   27
         Top             =   2400
         Width           =   4695
      End
      Begin VB.Label lblT 
         Caption         =   "..."
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
         Left            =   1920
         TabIndex        =   26
         Top             =   3480
         Width           =   4695
      End
      Begin VB.Label lblD 
         Caption         =   "..."
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
         Left            =   1920
         TabIndex        =   25
         Top             =   3120
         Width           =   4695
      End
      Begin VB.Label lblN 
         Caption         =   "..."
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
         Left            =   1920
         TabIndex        =   24
         Top             =   2760
         Width           =   4695
      End
      Begin VB.Line Line1 
         X1              =   1080
         X2              =   2520
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line2 
         X1              =   1080
         X2              =   2520
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line3 
         X1              =   600
         X2              =   2520
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line4 
         X1              =   1200
         X2              =   2520
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "FORMA PAGO:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label lblFP 
         Caption         =   "..."
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
         Left            =   1920
         TabIndex        =   22
         Top             =   4440
         Width           =   4695
      End
      Begin VB.Line Line5 
         X1              =   1800
         X2              =   1800
         Y1              =   2160
         Y2              =   4680
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "FOLIO DE FACTURA:"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "ORDEN DE COMPRA:"
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "SOLICITO:"
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   4920
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmFacturaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DIaS As Integer
Dim TOTAL_LETRA As Double
Dim DECIMALES As Double
Private cnn As ADODB.Connection


Private Sub cmdChange_Click()
On Error GoTo ManejaError
    Me.txtTL.Locked = False
    Me.txtTL.SetFocus
    Me.txtTL.SelStart = 0
    Me.txtTL.SelLength = Len(Me.txtTL.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub

Private Sub cmdFacturar_Click()
On Error GoTo ManejaError
    Dim sBusca As String
    Dim Retry As Integer
    Dim tRs As Recordset
    Dim tRs2 As Recordset
    Dim IDV As Integer
    Dim Path As String
    Retry = 1
    Path = App.Path
    If Dir(Path & "\REPORTES\Factura2.rpt") <> "" Then
        sBusca = "Select folio from Ventas Where folio = '" & Distin & txtFolio.Text & "'"
        Set tRs = cnn.Execute(sBusca)
        sBusca = "Select folio from FACTCAN Where folio = '" & Distin & txtFolio.Text & "'"
        Set tRs2 = cnn.Execute(sBusca)
        If (tRs.BOF And tRs.EOF) And (tRs2.BOF And tRs2.EOF) Then
            sBusca = "Select ID_Venta, ID_Cliente from Ventas Where ID_Venta in (" & txtVenta.Text & ") and Facturado < 1"
            Set tRs = cnn.Execute(sBusca)
            If Not (tRs.BOF And tRs.EOF) Then
                Do While Not (tRs.EOF)
                    IDV = tRs.Fields("ID_VENTA")
                    'deAPTONER.INSERTAR_DATOS_FACTURA_CREDITO IDV, Me.DTPicker1.Value, Me.DTPicker2.Value, Me.txtDH.Text, Me.txtMH.Text, Me.txtAH.Text, Me.txtDP.Text, Me.txtMP.Text, Me.txtAP.Text, Me.txtTL.Text, Me.lblFP.Caption, Distin & txtFolio.Text, txtOC.Text, txtCom.Text
                    tRs.MoveNext
                Loop
                Do While Retry = 1
                    Set crReport = crApplication.OpenReport(Path & "\REPORTES\Factura2.rpt")
                    crReport.Database.LogOnServer "p2ssql.dll", GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "1"), GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER"), GetSetting("APTONER", "ConfigSACC", "USUARIO", "1"), GetSetting("APTONER", "ConfigSACC", "PASSWORD", "1")
                    crReport.ParameterFields.Item(1).ClearCurrentValueAndRange
                    crReport.ParameterFields.Item(1).AddCurrentValue Trim(Distin & Me.txtFolio.Text)
                    crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    frmRep.Show vbModal, Me
                    If MsgBox("Se imprimio correctamente", vbYesNo, "SACC") = vbNo Then
                        If MsgBox("Desea Reintentarlo", vbYesNo, "SACC") = vbNo Then
                            Retry = 0
                        End If
                    Else
                        Retry = 2
                    End If
                Loop
                
                If Retry = 0 Then
                        If MsgBox("Desea Cancelar el Folio?", vbYesNo, "SACC") = vbNo Then
                            Folio = ""
                        Else
                            Folio = Distin & txtFolio.Text
                        End If
                        sBusca = "Select ID_Venta, ID_Cliente from Ventas Where ID_Venta in (" & txtVenta.Text & ") and Facturado = 1"
                        Set tRs = cnn.Execute(sBusca)
                        tRs.MoveFirst
                        Do While Not (tRs.EOF)
                            IDV = tRs.Fields("ID_VENTA")
                            sBusca = "UPDATE VENTAS SET FACTURADO = '0', FOLIO = '" & Folio & "' WHERE ID_VENTA = '" & IDV & "'"
                            cnn.Execute (sBusca)
                            If Folio <> "" Then
                                sBusca = "INSERT INTO FACTCAN (ID_VENTA, FOLIO, FECHA) VALUES('" & IDV & "', '" & Folio & "', '" & Format(Date, "dd/mm/yyyy") & "');"
                                cnn.Execute (sBusca)
                            End If
                            tRs.MoveNext
                        Loop
                        
                    End If
                'crReport.PrintOut , 3
            End If
        Else ' Checar centrado
            MsgBox "EL FOLIO DE LA FACTURA YA SE USO" & Chr(13) & "                VERIFIQUE", vbInformation, "SACC"
        End If
    Else
        MsgBox "FALTAN ARCHIVOS DEL SISTEMA, LLAME A SOPORTE", vbInformation, "SACC"
    End If
    Me.txtFolio.Text = ""
    Me.txtVenta.Text = ""
    Me.cmdFacturar.Enabled = False
    Me.cmdOk.Enabled = False
    Me.lblC.Caption = ""
    Me.lblD.Caption = ""
    DIaS = 0
    Me.lblN.Caption = ""
    Me.lblT.Caption = ""
    Me.txtTL.Text = ""
    Me.txtVenta.SetFocus
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

Private Sub cmdOk_Click()
On Error GoTo ManejaError
    Dim sBusca As String
    Dim tRs As Recordset
    Dim IDC As Integer
    Dim TotalAcum As Double
    
    Borrar_Campos
    DIaS = 0
    TOTAL_LETRA = 0
    DECIMALES = 0
    
    sBusca = "Select ID_Venta, ID_Cliente from Ventas Where ID_Venta in (" & txtVenta.Text & ")"
    Set tRs = cnn.Execute(sBusca)
    
    If Not (tRs.BOF And tRs.EOF) Then
        IDC = tRs.Fields("ID_CLIENTE")
        Do While Not (tRs.EOF)
            If IDC <> tRs.Fields("ID_CLIENTE") Then
                MsgBox "El Folio #" & tRs.Fields("ID_Venta") & " de las ventas no corresponde al mismo cliente"
                Exit Sub
            Else
                tRs.MoveNext
            End If
        Loop
    
        tRs.MoveFirst
        TotalAcum = 0
        Do While Not (tRs.EOF)
            IDC = tRs.Fields("ID_VENTA")
            If Puede_Facturar Then
                deAPTONER.DATOS_FACTURAS (IDC)
                With deAPTONER.rsDATOS_FACTURAS
                    If Not .EOF Then
                        If Not IsNull(!ID_CLIENTE) Then Me.lblC.Caption = Trim(!ID_CLIENTE)
                        If Not IsNull(!DIAS_CREDITO) Then Me.lblD.Caption = Trim(!DIAS_CREDITO)
                        If Not IsNull(!DIAS_CREDITO) Then DIaS = Trim(!DIAS_CREDITO)
                        If Not IsNull(!Nombre) Then Me.lblN.Caption = Trim(!Nombre)
                        If Not IsNull(!Total) Then TotalAcum = TotalAcum + Trim(!Total) 'Me.lblT.Caption = Trim(!TOTAL)
                        If Not IsNull(!Total) Then TOTAL_LETRA = TOTAL_LETRA + Val(!Total)
                        If Not IsNull(!Total) Then DECIMALES = DECIMALES + (!Total - Int(!Total)) * 100
                        'Me.txtTL.Text = Letra(TOTAL_LETRA) & " " & FormatNumber(DECIMALES, 0) & "/100"
                        
                        'If Val(DIaS) = 0 Then
                            'Me.lblFP.Caption = "PAGO EN UNA SOLA EXIBICION"
                        'Else
                            Me.lblFP.Caption = "PAGO EN PARCIALIDADES"
                        'End If
            
                    End If
                    .Close
                End With
                Me.cmdFacturar.Enabled = False
            End If
            tRs.MoveNext
            
            Me.DTPicker2.Value = Format(Date, "dd/mm/yyyy") + Val(DIaS)
            Me.txtDP.Text = Me.DTPicker2.Day
            Me.txtMP.Text = MESES(Me.DTPicker2.Month)
            Me.txtAP.Text = Me.DTPicker2.Year
            Me.txtFolio.SetFocus
        Loop
        Me.txtTL.Text = Letra(TOTAL_LETRA) & " PESOS " & FormatNumber(DECIMALES, 0) & "/100 M.N."
        Me.lblT.Caption = TotalAcum
    Else
        MsgBox "Las Ventas Pedidas no Existen"
    End If
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

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Dim sBusca As String
    Dim tRs As Recordset

    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    Me.DTPicker1.Value = Format(Date, "dd/mm/yyyy")

    sBusca = "Select DISTINTIVO from Sucursales Where Nombre = '" & VarMen.Text4(0).Text & "'"
    Set tRs = cnn.Execute(sBusca)
    If Not (tRs.BOF And tRs.EOF) Then
        Distin = tRs.Fields("DISTINTIVO")
    End If
    Me.txtDH.Text = Me.DTPicker1.Day
    Me.txtMH.Text = MESES(Me.DTPicker1.Month)
    Me.txtAH.Text = Me.DTPicker1.Year
    
    Me.lblT.Caption = Ventas.Text5.Text
    Me.txtVenta.Text = Ventas.IdBenta.Text
    Me.cmdOk.Value = True
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub

Private Sub Image9_Click()
On Error GoTo ManejaError
    Unload Me
    Unload FrmRegVentCred
    Unload FrmVentaCredito
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub

Private Sub txtCom_GotFocus()
    txtCom.BackColor = &HFFE1E1
End Sub

Private Sub txtCom_LostFocus()
    txtCom.BackColor = &H80000005
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If (KeyAscii = 13) And (txtFolio.Text <> "") Then
        Me.cmdFacturar.Value = True
    Else
        If txtFolio.Text <> "" Then
            Me.cmdFacturar.Enabled = True
        Else
            Me.cmdFacturar.Enabled = False
        End If
        Dim Valido As String
        Valido = "1234567890"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
            End If
        End If
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub txtFolio_GotFocus()
    txtFolio.BackColor = &HFFE1E1
End Sub
Private Sub txtFolio_LostFocus()
      txtFolio.BackColor = &H80000005
End Sub

Private Sub txtOC_GotFocus()
    txtOC.BackColor = &HFFE1E1
End Sub

Private Sub txtOC_LostFocus()
    txtOC.BackColor = &H80000005
End Sub

Private Sub txtVenta_Change()
On Error GoTo ManejaError
    Me.cmdFacturar.Enabled = False
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub

Private Sub txtVenta_GotFocus()
On Error GoTo ManejaError
    txtVenta.BackColor = &HFFE1E1
    txtVenta.SelStart = 0
    txtVenta.SelLength = Len(txtVenta.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub txtVenta_LostFocus()
      txtVenta.BackColor = &H80000005
End Sub

Private Sub txtVenta_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.cmdOk.Value = True
        'Me.cmdFacturar.SetFocus
    Else
        Dim Valido As String
        Valido = "1234567890."
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii > 26 Then
                If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
                End If
            End If
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub

Function Letra(Numero)
Dim Texto
Dim Millones
Dim Miles
Dim Cientos
Dim DECIMALES
Dim CADENA
Dim CadMillones
Dim CadMiles
Dim CadCientos
Dim caddecimales
Texto = Round(Val(Numero), 2)
Texto = FormatNumber(Texto, 2)
Texto = Right(Space(14) & Texto, 14)
Millones = Mid(Texto, 1, 3)
Miles = Mid(Texto, 5, 3)
Cientos = Mid(Texto, 9, 3)
DECIMALES = Mid(Texto, 13, 2)
CadMillones = ConvierteCifra(Millones, False)
CadMiles = ConvierteCifra(Miles, False)
CadCientos = ConvierteCifra(Cientos, True)
caddecimales = ConvierteDecimal(DECIMALES)
If Trim(CadMillones) > "" Then
If Trim(CadMillones) = "UN" Then
CADENA = CadMillones & " MILLON"
Else
CADENA = CadMillones & " MILLONES"
End If
End If
If Trim(CadMiles) > "" Then
If Trim(CadMiles) = "UN" Then
CadMiles = ""
CADENA = CADENA & "" & CadMiles & "MIL"
CadMiles = "UN"
Else
CADENA = CADENA & " " & CadMiles & " MIL"
End If
End If

If DECIMALES = "00" Then
If Trim(CadMillones & CadMiles & CadCientos & caddecimales) = "UN" Then
CADENA = CADENA & "UNO "
Else
If Miles & Cientos = "000000" Then
CADENA = CADENA & " " & Trim(CadCientos)
Else
CADENA = CADENA & " " & Trim(CadCientos)
End If
Letra = Trim(CADENA)
End If
Else
If Trim(CadMillones & CadMiles & CadCientos & caddecimales) = "UN" Then
CADENA = CADENA & "UNO " & Trim(caddecimales)
Else
If Miles & Cientos = "000000" Then
CADENA = CADENA & " " & Trim(CadCientos) & Trim(caddecimales)
Else
CADENA = CADENA & " " & Trim(CadCientos) & Trim(caddecimales)
End If
Letra = Trim(CADENA)
End If
End If

End Function

Function ConvierteCifra(Texto, IsCientos As Boolean)
Dim Centena
Dim Decena
Dim UNIDAD
Dim txtCentena
Dim txtDecena
Dim txtUnidad
Centena = Mid(Texto, 1, 1)
Decena = Mid(Texto, 2, 1)
UNIDAD = Mid(Texto, 3, 1)
Select Case Centena
Case "1"
txtCentena = "CIEN"
If Decena & UNIDAD <> "00" Then
txtCentena = "CIENTO"
End If
Case "2"
txtCentena = "DOSCIENTOS"
Case "3"
txtCentena = "TRESCIENTOS"
Case "4"
txtCentena = "CUATROCIENTOS"
Case "5"
txtCentena = "QUINIENTOS"
Case "6"
txtCentena = "SEISCIENTOS"
Case "7"
txtCentena = "SETECIENTOS"
Case "8"
txtCentena = "OCHOCIENTOS"
Case "9"
txtCentena = "NOVECIENTOS"
End Select

Select Case Decena
Case "1"
txtDecena = "DIEZ"
Select Case UNIDAD
Case "1"
txtDecena = "ONCE"
Case "2"
txtDecena = "DOCE"
Case "3"
txtDecena = "TRECE"
Case "4"
txtDecena = "CATORCE"
Case "5"
txtDecena = "QUINCE"
Case "6"
txtDecena = "DIECISEIS"
Case "7"
txtDecena = "DIECISIETE"
Case "8"
txtDecena = "DIECIOCHO"
Case "9"
txtDecena = "DIECINUEVE"
End Select
Case "2"
txtDecena = "VEINTE"
If UNIDAD <> "0" Then
txtDecena = "VEINTI"
End If
Case "3"
txtDecena = "TREINTA"
If UNIDAD <> "0" Then
txtDecena = "TREINTA Y "
End If
Case "4"
txtDecena = "CUARENTA"
If UNIDAD <> "0" Then
txtDecena = "CUARENTA Y "
End If
Case "5"
txtDecena = "CINCUENTA"
If UNIDAD <> "0" Then
txtDecena = "CINCUENTA Y "
End If
Case "6"
txtDecena = "SESENTA"

If UNIDAD <> "0" Then
txtDecena = "SESENTA Y "
End If
Case "7"
txtDecena = "SETENTA"
If UNIDAD <> "0" Then
txtDecena = "SETENTA Y "
End If
Case "8"
txtDecena = "OCHENTA"
If UNIDAD <> "0" Then
txtDecena = "OCHENTA Y "
End If
Case "9"
txtDecena = "NOVENTA"
If UNIDAD <> "0" Then
txtDecena = "NOVENTA Y "
End If
End Select

If Decena <> "1" Then
Select Case UNIDAD
Case "1"
If IsCientos = False Then
txtUnidad = "UN"
Else
txtUnidad = "UNO"
End If
Case "2"
txtUnidad = "DOS"
Case "3"
txtUnidad = "TRES"
Case "4"
txtUnidad = "CUATRO"
Case "5"
txtUnidad = "CINCO"
Case "6"
txtUnidad = "SEIS"
Case "7"
txtUnidad = "SIETE"
Case "8"
txtUnidad = "OCHO"
Case "9"
txtUnidad = "NUEVE"
End Select
End If
ConvierteCifra = txtCentena & " " & txtDecena & txtUnidad
End Function

Function ConvierteDecimal(Texto)
Dim Decenadecimal
Dim Unidaddecimal
Dim txtDecenadecimal
Dim txtUnidaddecimal
Decenadecimal = Mid(Texto, 1, 1)
Unidaddecimal = Mid(Texto, 2, 1)

Select Case Decenadecimal
Case "1"
txtDecenadecimal = ""
Select Case Unidaddecimal
Case "1"
txtDecenadecimal = ""
Case "2"
txtDecenadecimal = ""
Case "3"
txtDecenadecimal = ""
Case "4"
txtDecenadecimal = ""
Case "5"
txtDecenadecimal = ""
Case "6"
txtDecenadecimal = ""
Case "7"
txtDecenadecimal = ""
Case "8"
txtDecenadecimal = ""
Case "9"
txtDecenadecimal = ""
End Select
Case "2"
txtDecenadecimal = ""
If Unidaddecimal <> "0" Then
txtDecenadecimal = ""
End If
Case "3"
txtDecenadecimal = ""
If Unidaddecimal <> "0" Then
txtDecenadecimal = ""
End If
Case "4"
txtDecenadecimal = ""
If Unidaddecimal <> "0" Then
txtDecenadecimal = ""
End If
Case "5"
txtDecenadecimal = ""
If Unidaddecimal <> "0" Then
txtDecenadecimal = ""
End If
Case "6"
txtDecenadecimal = ""

If Unidaddecimal <> "0" Then
txtDecenadecimal = ""
End If
Case "7"
txtDecenadecimal = ""
If Unidaddecimal <> "0" Then
txtDecenadecimal = ""
End If
Case "8"
txtDecenadecimal = ""
If Unidaddecimal <> "0" Then
txtDecenadecimal = ""
End If
Case "9"
txtDecenadecimal = ""
If Unidaddecimal <> "0" Then
txtDecenadecimal = ""
End If
End Select

If Decenadecimal <> "1" Then
Select Case Unidaddecimal
Case "1"
txtUnidaddecimal = ""
Case "2"
txtUnidaddecimal = ""
Case "3"
txtUnidaddecimal = ""
Case "4"
txtUnidaddecimal = ""
Case "5"
txtUnidaddecimal = ""
Case "6"
txtUnidaddecimal = ""
Case "7"
txtUnidaddecimal = ""
Case "8"
txtUnidaddecimal = ""
Case "9"
txtUnidaddecimal = ""
End Select
End If
If Decenadecimal = 0 And Unidaddecimal = 0 Then
ConvierteDecimal = ""
Else
ConvierteDecimal = txtDecenadecimal & txtUnidaddecimal
End If
End Function

Function Puede_Facturar() As Boolean
On Error GoTo ManejaError
    If Trim(Me.txtVenta.Text) = "" Then
        MsgBox "ESCRIBA EL MUNARO DE VENTA PARA FACTURAR", vbInformation, "SACC"
        Me.txtVenta.SetFocus
        Puede_Facturar = False
        'Borrar_Campos
        Exit Function
    End If
    
    deAPTONER.Puede_Facturar Val(Trim(Me.txtVenta.Text))
    With deAPTONER.rsPUEDE_FACTURAR
        If Not (.BOF Or .EOF) Then
            If (!FACTURADO) = 0 Then
                Puede_Facturar = True
            Else
                MsgBox "ESTA VENTA YA FUE FACTURADA" & Chr(13) & "         FACTURA #" & !Folio, vbInformation, "SACC"
                Puede_Facturar = False
            End If
        Else
            MsgBox "NUMERO DE VENTA INCORRECTO", vbInformation, "SACC"
            Me.txtVenta.SetFocus
            Puede_Facturar = False
            'Borrar_Campos
        End If
        .Close
    End With
Exit Function
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Function

Sub Borrar_Campos()
On Error GoTo ManejaError
    Me.lblC.Caption = ""
    Me.lblD.Caption = ""
    Me.lblN.Caption = ""
    Me.lblT.Caption = ""
    Me.lblFP.Caption = ""
    Me.txtTL.Text = ""
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
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


