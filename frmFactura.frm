VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmFactura 
   BackColor       =   &H00FFFFFF&
   Caption         =   "FACTURAS"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7200
      TabIndex        =   32
      Top             =   5280
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmFactura.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmFactura.frx":030A
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7440
      Top             =   1320
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Factura"
      TabPicture(0)   =   "frmFactura.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFP"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblN"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblD"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblT"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblC"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label5"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label7"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label8"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label9"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label11"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label12"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label13"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdChange"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtTL"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdFacturar"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdOk"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtVenta"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtFolio"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtOC"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtCom"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtFolioCorrecto"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TxtPedimento"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "DTPicker3"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "TxtAduana"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Buscar"
      TabPicture(1)   =   "frmFactura.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(1)=   "Text1"
      Tab(1).Control(2)=   "Option1"
      Tab(1).Control(3)=   "Option2"
      Tab(1).Control(4)=   "ListView1"
      Tab(1).Control(5)=   "Command1"
      Tab(1).ControlCount=   6
      Begin VB.TextBox TxtAduana 
         Height          =   285
         Left            =   3720
         TabIndex        =   45
         Top             =   5880
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   285
         Left            =   3720
         TabIndex        =   42
         Top             =   5520
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Format          =   51576833
         CurrentDate     =   40498
      End
      Begin VB.TextBox TxtPedimento 
         Height          =   285
         Left            =   3720
         TabIndex        =   40
         Top             =   5160
         Width           =   1935
      End
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
         Left            =   -70560
         Picture         =   "frmFactura.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   720
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   38
         Top             =   1200
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   8705
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
         Caption         =   "Nota de Venta"
         Height          =   195
         Left            =   -72240
         TabIndex        =   37
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Factura"
         Height          =   195
         Left            =   -72240
         TabIndex        =   36
         Top             =   600
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -73920
         TabIndex        =   35
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtFolioCorrecto 
         Height          =   285
         Left            =   6240
         TabIndex        =   31
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtCom 
         Height          =   285
         Left            =   3720
         TabIndex        =   2
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox txtOC 
         Height          =   285
         Left            =   3720
         TabIndex        =   3
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox txtFolio 
         Height          =   285
         Left            =   3720
         TabIndex        =   4
         Top             =   4800
         Width           =   1935
      End
      Begin VB.TextBox txtVenta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5520
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
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
         Picture         =   "frmFactura.frx":4DF6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
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
         Height          =   405
         Left            =   5760
         Picture         =   "frmFactura.frx":77C8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5760
         Width           =   1095
      End
      Begin VB.TextBox txtTL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3240
         Width           =   6015
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "..."
         Height          =   255
         Left            =   6360
         TabIndex        =   15
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "NOMBRE ADUANA:"
         Height          =   255
         Left            =   1800
         TabIndex        =   44
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "FECHA PEDIMENTO:"
         Height          =   255
         Left            =   1800
         TabIndex        =   43
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "PEDIMENTO ADUANAL:"
         Height          =   255
         Left            =   1800
         TabIndex        =   41
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   34
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "SOLICITO:"
         Height          =   255
         Left            =   1800
         TabIndex        =   30
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "ORDEN DE COMPRA:"
         Height          =   255
         Left            =   1800
         TabIndex        =   29
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "* FOLIO DE FACTURA:"
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
         Left            =   1440
         TabIndex        =   28
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "NOTA DE VENTA:"
         Height          =   255
         Left            =   3960
         TabIndex        =   27
         Top             =   720
         Width           =   1455
      End
      Begin VB.Line Line4 
         X1              =   1200
         X2              =   2520
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line3 
         X1              =   600
         X2              =   2520
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line2 
         X1              =   1080
         X2              =   2520
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         X1              =   1080
         X2              =   2520
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "CLIENTE:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "DIAS CREDITO:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "NOMBRE:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2040
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
         TabIndex        =   22
         Top             =   1680
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
         TabIndex        =   21
         Top             =   2760
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
         TabIndex        =   20
         Top             =   2400
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
         TabIndex        =   19
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "FORMA PAGO:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3720
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
         TabIndex        =   17
         Top             =   3720
         Width           =   4695
      End
      Begin VB.Line Line5 
         X1              =   1800
         X2              =   1800
         Y1              =   1440
         Y2              =   3960
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.TextBox txtMH 
      Height          =   285
      Left            =   7440
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtAH 
      Height          =   285
      Left            =   7200
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtDP 
      Height          =   285
      Left            =   7680
      TabIndex        =   11
      Text            =   " "
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Left            =   7440
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtAP 
      Height          =   285
      Left            =   7200
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtDH 
      Height          =   285
      Left            =   7680
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51576833
      CurrentDate     =   38728
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51576833
      CurrentDate     =   38728
   End
End
Attribute VB_Name = "frmFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DIaS As Integer
Dim TOTAL_LETRA As Double
Dim DECIMALES As Double
Dim Distin As String
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
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
    Dim tRs1 As ADODB.Recordset
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim IDV As Integer
    Dim Path As String
    Dim Retry As Integer
    Dim Folio As String
    Dim Resp As Integer
    Dim folioant As Boolean
    Dim ClvCliente As Integer
    Dim Conta As Integer
    Dim StrFact As String
    Retry = 1
    Path = App.Path
    folioant = True
    If txtFolioCorrecto.Text = "NULL" Then
        folioant = False
        txtFolioCorrecto.Text = 1
    End If
    If CDbl(txtFolio.Text) >= CDbl(txtFolioCorrecto.Text) Then
        sBusca = "Select folio from Ventas Where folio = '" & Distin & txtFolio.Text & "'"
        Set tRs = cnn.Execute(sBusca)
        sBusca = "Select folio from FACTCAN Where folio = '" & Distin & txtFolio.Text & "'"
        Set tRs2 = cnn.Execute(sBusca)
        If (tRs.BOF And tRs.EOF) And (tRs2.BOF And tRs2.EOF) Then
            If CDbl(txtFolio.Text) <> CDbl(txtFolioCorrecto.Text) Then
                Resp = MsgBox("                              EL FOLIO DE LA FACTURA ES MAYOR AL CONSECUTIVO," & Chr(13) _
                            & "DE CONTINUAR USANDO ESE FOLIO LOS FOLIOS ANTERIORES NO USADOS SE CANCELARAN" & Chr(13) & _
                              "                             DESEA QUE SE CAMBIE EL FOLIO POR EL CONSECUTIVO " & txtFolioCorrecto.Text, vbYesNoCancel, "SACC")
                If Resp = 6 Then
                    txtFolio.Text = txtFolioCorrecto.Text
                ElseIf Resp = 2 Then
                    txtFolio.Text = ""
                End If
            End If
            If txtFolio.Text <> "" Then
                sBusca = "Select ID_Venta, ID_Cliente from Ventas Where ID_Venta in (" & txtVenta.Text & ") and Facturado < 1"
                Set tRs = cnn.Execute(sBusca)
                If Not (tRs.BOF And tRs.EOF) Then
                    Do While Not (tRs.EOF)
                        IDV = tRs.Fields("ID_VENTA")
                        sBusca = "UPDATE  VENTAS  SET FACTURADO = 1, FECHA_PAGARE = '" & Me.DTPicker1.value & "', FECHA_HOY = '" & Me.DTPicker2.value & "', DIA_HOY = '" & Me.txtDH.Text & "', MES_HOY = '" & Me.txtMH.Text & "', AÑO_HOY = '" & Me.txtAH.Text & "', DIA_PAGARE = '" & Me.txtDP.Text & "', MES_PAGARE = '" & Me.txtMP.Text & "', AÑO_PAGARE = '" & Me.txtAP.Text & "', TOTAL_LETRA = '" & Me.txtTL.Text & "', FOLIO = '" & Distin & txtFolio.Text & "', COMENTARIO = '" & txtCom.Text & "', NOOC = '" & txtOC.Text & "', FECHA_FACTURA =  DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), EFECTO_COMPROBANTE = 'I', PEDIMENTO_ADUANA = '" & TxtPedimento.Text & "', FECHA_PEDIMENTO = '" & DTPicker3.value & "', NOMBRE_ADUANA = '" & TxtAduana.Text & "' Where ID_VENTA = " & IDV
                        cnn.Execute (sBusca)
                        ClvCliente = tRs.Fields("ID_CLIENTE")
                        tRs.MoveNext
                    Loop
                    Do While Retry = 1
                        sBusca = "SELECT LEYENDAS FROM CLIENTE WHERE ID_CLIENTE = " & ClvCliente
                        Set tRs1 = cnn.Execute(sBusca)
                        If Not (tRs1.EOF And tRs1.BOF) Then
                            If tRs1.Fields("LEYENDAS") = "N" Then
                                FunFactura
                            Else
                                FunFactura
                            End If
                        Else
                            FunFactura
                        End If
                        If MsgBox("Se imprimio correctamente", vbYesNo, "SACC") = vbNo Then
                            If MsgBox("Desea Reintentarlo", vbYesNo, "SACC") = vbNo Then
                                Retry = 0
                            End If
                        Else
                            Retry = 2
                        End If
                    Loop
                    Folio = Distin & txtFolio.Text
                    If Retry = 0 Then
                        If MsgBox("Desea Cancelar el Folio?", vbYesNo, "SACC") = vbNo Then
                            Folio = ""
                        End If
                        sBusca = "Select ID_Venta, ID_Cliente from Ventas Where ID_Venta in (" & txtVenta.Text & ") and Facturado = 1"
                        Set tRs = cnn.Execute(sBusca)
                        tRs.MoveFirst
                        Do While Not (tRs.EOF)
                            IDV = tRs.Fields("ID_VENTA")
                            sBusca = "UPDATE VENTAS SET FACTURADO = '0', FOLIO = '" & Folio & "' WHERE ID_VENTA = " & IDV
                            cnn.Execute (sBusca)
                            If Folio <> "" Then
                                sBusca = "INSERT INTO FACTCAN (ID_VENTA, FOLIO, FECHA) VALUES('" & IDV & "', '" & Folio & "', '" & Format(Date, "dd/mm/yyyy") & "');"
                                cnn.Execute (sBusca)
                            End If
                            tRs.MoveNext
                        Loop
                    End If
                    If Folio <> "" Then
                        For Conta = 1 To Len(txtFolio.Text)
                            If IsNumeric(Mid(txtFolio.Text, Conta, 1)) Then
                                StrFact = StrFact & Mid(txtFolio.Text, Conta, 1)
                            End If
                        Next
                        If folioant Then
                            sBusca = "UPDATE FOLIOSUC SET FOLIO = " & StrFact & " WHERE SUCURSAL = '" & VarMen.Text4(0).Text & "'"
                        Else
                            sBusca = "INSERT INTO FOLIOSUC (FOLIO, SUCURSAL) VALUES(" & StrFact & ", '" & VarMen.Text4(0).Text & "');"
                        End If
                        cnn.Execute (sBusca)
                    End If
                End If
            End If
        Else ' Checar centrado
            MsgBox "EL FOLIO DE LA FACTURA YA SE USO" & Chr(13) & "             VERIFIQUE", vbInformation, "SACC"
        End If
    Else
        MsgBox "EL FOLIO DE LA FACTURA NO PUEDE SER MENOR AL ULTIMO USADO" & Chr(13) & "                                      VERIFIQUE", vbInformation, "SACC"
    End If
    txtVenta.Enabled = True
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
    CommonDialog1.Copies = 1
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdOk_Click()
On Error GoTo ManejaError
    Dim sBusca As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim Anterior As String
    Dim IDC As Integer
    Dim TotalAcum As Double
    Borrar_Campos
    DIaS = 0
    TOTAL_LETRA = 0
    DECIMALES = 0
    sBusca = "SELECT ID_CLIENTE FROM VENTAS WHERE ID_VENTA IN (" & txtVenta.Text & ") GROUP BY ID_CLIENTE"
    Set tRs = cnn.Execute(sBusca)
    If tRs.Fields.COUNT <= 1 Then
        sBusca = "SELECT ID_CLIENTE, SUM(TOTAL) AS TOTAL FROM VENTAS WHERE ID_VENTA IN (" & txtVenta.Text & ") GROUP BY ID_CLIENTE"
        Set tRs = cnn.Execute(sBusca)
        TOTAL_LETRA = tRs.Fields("TOTAL")
        Me.lblT.Caption = Format(tRs.Fields("TOTAL"), "###,###,##0.00")
        sBusca = "select id_venta from ventas where datediff (month, FECHA,  DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE()))) = 0 and id_venta in (" & txtVenta.Text & ")"
        Set tRs = cnn.Execute(sBusca)
        If (tRs.EOF And tRs.BOF) Then
            sBusca = "select id_venta from ventas where UNA_EXIBICION = 'N' and id_venta in (" & txtVenta.Text & ")"
            Set tRs = cnn.Execute(sBusca)
            If (tRs.EOF And tRs.BOF) Then
                sBusca = "select id_venta from ventas where facturado = 1 and id_venta in (" & txtVenta.Text & ")"
                Set tRs = cnn.Execute(sBusca)
                If (tRs.EOF And tRs.BOF) Then
                    MsgBox "IMPOSIBLE FACTURAR UNA VENTA SI EL MES ES DIFERENTE", vbInformation, "SACC"
                    Exit Sub
                End If
            End If
        End If
        sBusca = "Select ID_Venta, ID_Cliente, Una_Exibicion, isNull(COMENTARIO, '') as COMENTARIO, isNull(NOOC, '') as NOOC from Ventas Where ID_Venta in (" & txtVenta.Text & ") and Sucursal = '" & VarMen.Text4(0).Text & "'"
        Set tRs = cnn.Execute(sBusca)
        If Not (tRs.BOF And tRs.EOF) Then
            IDC = tRs.Fields("ID_CLIENTE")
            If tRs.Fields("Una_Exibicion") = "S" Or IsNull(tRs.Fields("Una_Exibicion")) Then
                Anterior = "S"
            Else
                Anterior = "N"
            End If
            Do While Not (tRs.EOF)
                If (IDC <> tRs.Fields("ID_CLIENTE")) Or (Anterior <> tRs.Fields("Una_Exibicion")) Then
                    If (IDC <> tRs.Fields("ID_CLIENTE")) Then
                        MsgBox "El Folio #" & tRs.Fields("ID_Venta") & " de las ventas no corresponde al mismo cliente"
                    Else
                        MsgBox "El Folio #" & tRs.Fields("ID_Venta") & " de las ventas no corresponde al mismo tipo de pago"
                    End If
                    Exit Sub
                Else
                    tRs.MoveNext
                End If
            Loop
            TotalAcum = 0
            DECIMALES = (CDbl(TOTAL_LETRA) - Int(CDbl(TOTAL_LETRA))) * 100
            Me.txtTL.Text = Letra(TOTAL_LETRA) & " PESOS " & FormatNumber(DECIMALES, 0) & "/100 M.N."
            tRs.MoveFirst
            Do While Not (tRs.EOF)
                IDC = tRs.Fields("ID_VENTA")
                'txtCom.Text = tRs.Fields("COMENTARIO")
                'txtOC.Text = tRs.Fields("NOOC")
                
                If Puede_Facturar(IDC) Then
                    sBuscar = "SELECT  V.ID_CLIENTE, V.TOTAL, C.NOMBRE, V.DIAS_CREDITO, V.UNA_EXIBICION FROM VENTAS AS  V JOIN CLIENTE AS C ON C.ID_CLIENTE = V.ID_CLIENTE WHERE V.ID_Venta = " & IDC
                    Set tRs2 = cnn.Execute(sBuscar)
                    If Not tRs.EOF Then
                        If Not IsNull(tRs2.Fields("ID_CLIENTE")) Then Me.lblC.Caption = Trim(tRs2.Fields("ID_CLIENTE"))
                        If Not IsNull(tRs2.Fields("DIAS_CREDITO")) Then Me.lblD.Caption = Trim(tRs2.Fields("DIAS_CREDITO"))
                        If Not IsNull(tRs2.Fields("DIAS_CREDITO")) Then DIaS = Trim(tRs2.Fields("DIAS_CREDITO"))
                        If Not IsNull(tRs2.Fields("Nombre")) Then Me.lblN.Caption = Trim(tRs2.Fields("Nombre"))
                        If Not IsNull(tRs2.Fields("Total")) Then TotalAcum = TotalAcum + Trim(tRs2.Fields("Total")) 'Me.lblT.Caption = Trim(tRs.Fields("Total"))
                        If Not IsNull(tRs2.Fields("Total")) Then DECIMALES = DECIMALES + (tRs2.Fields("Total") - Int(tRs2.Fields("Total"))) * 100
                        If IsNull(tRs2.Fields("UNA_EXIBICION")) Or tRs2.Fields("UNA_EXIBICION") = "S" Then
                            Me.lblFP.Caption = "PAGO EN UNA SOLA EXHIBICION"
                        Else
                            Me.lblFP.Caption = "PAGO EN PARCIALIDADES"
                        End If
                        sBuscar = "SELECT LEYENDAS FROM CLIENTE WHERE ID_CLIENTE = " & Trim(tRs.Fields("ID_CLIENTE"))
                        Set tRs1 = cnn.Execute(sBuscar)
                        If Not (tRs1.EOF And tRs1.BOF) Then
                            If tRs1.Fields("LEYENDAS") = "N" Then
                                Me.lblFP.Caption = ""
                                Me.lblD.Caption = ""
                                DIaS = "0"
                            End If
                        End If
                    End If
                    Me.cmdFacturar.Enabled = False
                End If
                tRs.MoveNext
                Me.DTPicker2.value = Format(Date + Val(DIaS), "dd/mm/yyyy")
                Me.txtDP.Text = Me.DTPicker2.Day
                Me.txtMP.Text = MESES(Me.DTPicker2.Month)
                Me.txtAP.Text = Me.DTPicker2.Year
                Me.txtFolio.SetFocus
            Loop
            sBusca = "Select top 1 FOLIO from FOLIOSUC Where SUCURSAL = '" & VarMen.Text4(0).Text & "'"
            Set tRs = cnn.Execute(sBusca)
            If Not (tRs.BOF And tRs.EOF) Then
                txtFolio.Text = CDbl(tRs.Fields("FOLIO")) + 1
                txtFolioCorrecto.Text = CDbl(tRs.Fields("FOLIO")) + 1
            Else
                txtFolio.Text = "1"
                txtFolioCorrecto.Text = "NULL"
            End If
        Else
            If MsgBox("ESTA VENTA ES DE OTRA SUCURSAL DESEA VER LA IMPRESION? ", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
                FunFactura
            End If
        End If
    Else
        MsgBox "Las notas dadas no corresponden a un solo cliente, favor de verificarlas!", vbExclamation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdSalir_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command1_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    If Option1.value Then
        sBuscar = "SELECT ID_VENTA, FOLIO, NOMBRE, FECHA, UNA_EXIBICION, SUBTOTAL, IVA, TOTAL, ID_CLIENTE FROM VENTAS WHERE FOLIO LIKE '%" & Text1.Text & "%'"
    Else
        sBuscar = "SELECT ID_VENTA, FOLIO, NOMBRE, FECHA, UNA_EXIBICION, SUBTOTAL, IVA, TOTAL, ID_CLIENTE FROM VENTAS WHERE ID_VENTA = " & Text1.Text
    End If
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_VENTA"))
            If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(1) = tRs.Fields("FOLIO")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(3) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("UNA_EXIBICION")) Then
                If tRs.Fields("UNA_EXIBICION") = "S" Then
                    tLi.SubItems(4) = "CONTADO"
                Else
                     tLi.SubItems(4) = "CREDITO"
                End If
                sBuscar = "SELECT LEYENDAS FROM CLIENTE WHERE ID_CLIENTE = " & tRs.Fields("ID_CLIENTE")
                Set tRs1 = cnn.Execute(sBuscar)
                If Not (tRs1.EOF And tRs1.BOF) Then
                    If tRs1.Fields("LEYENDAS") = "N" Then tLi.SubItems(4) = ""
                End If
            End If
            If Not IsNull(tRs.Fields("SUBTOTAL")) Then tLi.SubItems(5) = tRs.Fields("SUBTOTAL")
            If Not IsNull(tRs.Fields("IVA")) Then tLi.SubItems(6) = tRs.Fields("IVA")
            If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(7) = tRs.Fields("TOTAL")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Me.DTPicker1.value = Format(Date, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Dim sBusca As String
    Dim tRs As ADODB.Recordset
    sBusca = "SELECT DISTINTIVO FROM SUCURSALES WHERE NOMBRE = '" & VarMen.Text4(0).Text & "' AND ELIMINADO = 'N'"
    Set tRs = cnn.Execute(sBusca)
    If Not (tRs.BOF And tRs.EOF) Then
        Distin = tRs.Fields("DISTINTIVO")
    End If
    Me.txtDH.Text = Me.DTPicker1.Day
    Me.txtMH.Text = MESES(Me.DTPicker1.Month)
    Me.txtAH.Text = Me.DTPicker1.Year
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Venta", 1000
        .ColumnHeaders.Add , , "Folio", 1000
        .ColumnHeaders.Add , , "Cliente", 4500
        .ColumnHeaders.Add , , "Fecha", 1000
        .ColumnHeaders.Add , , "Tipo", 1000
        .ColumnHeaders.Add , , "Subtotal", 1000
        .ColumnHeaders.Add , , "IVA", 1000
        .ColumnHeaders.Add , , "Total", 1000
    End With
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
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtVenta = Item
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1.value = True
    End If
    Dim Valido As String
    If Option1.value = True Then
        Valido = "1234567890ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-"
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
Private Sub Timer1_Timer()
    cmdOk.value = True
    Timer1.Enabled = False
End Sub
Private Sub TxtAduana_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ &%$()/-_"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtCom_GotFocus()
    txtCom.BackColor = &HFFE1E1
End Sub
Private Sub txtCom_LostFocus()
    txtCom.BackColor = &H80000005
End Sub
Private Sub txtFolio_Change()
    If txtFolio.Text <> "" Then
        Me.cmdFacturar.Enabled = True
    Else
        Me.cmdFacturar.Enabled = False
    End If
End Sub
Private Sub txtFolio_GotFocus()
    txtFolio.BackColor = &HFFE1E1
End Sub
Private Sub txtFolio_LostFocus()
    txtFolio.BackColor = &H80000005
End Sub
Private Sub txtFolio_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If (KeyAscii = 13) And (txtFolio.Text <> "") Then
        Me.cmdFacturar.value = True
    Else
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
Private Sub txtOC_GotFocus()
    txtOC.BackColor = &HFFE1E1
End Sub
Private Sub txtOC_LostFocus()
    txtOC.BackColor = &H80000005
End Sub
Private Sub TxtPedimento_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ &%$()/-_"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtVenta_Change()
On Error GoTo ManejaError
    Me.cmdFacturar.Enabled = False
    If txtVenta.Text <> "" Then
        cmdOk.Enabled = True
    Else
        cmdOk.Enabled = False
    End If
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
        Me.cmdOk.value = True
    Else
        Dim Valido As String
        Valido = "1234567890,"
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
Function Puede_Facturar(NV As Integer) As Boolean
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If Trim(Me.txtVenta.Text) = "" Then
        MsgBox "ESCRIBA EL NUMERO DE VENTA PARA FACTURAR", vbInformation, "SACC"
        Me.txtVenta.SetFocus
        Puede_Facturar = False
        Exit Function
    End If
    sBuscar = "SELECT FACTURADO, SUCURSAL, FOLIO FROM VENTAS WHERE ID_VENTA = " & NV
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF Or tRs.EOF) Then
        If tRs.Fields("FACTURADO") = 0 Then
            If Trim(tRs.Fields("SUCURSAL")) = VarMen.Text4(0).Text Then
                Puede_Facturar = True
            Else
                MsgBox "VENTA DE OTRA SUCURSAL" & Chr(13) & "         VENTA #" & NV, vbInformation, "SACC"
            End If
        Else
            If MsgBox("ESTA VENTA YA FUE FACTURADA FACTURA #" & tRs.Fields("FOLIO") & Chr(13) & "                       DESEA REIMPRIMIRLA? ", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
                FunFactura
            End If
            Puede_Facturar = False
        End If
    Else
        MsgBox "NUMERO DE VENTA INCORRECTO", vbInformation, "SACC"
        Me.txtVenta.SetFocus
        Puede_Facturar = False
    End If
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
Private Sub FunFactura()
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
    Dim tRs4 As ADODB.Recordset
    Dim tRs5 As ADODB.Recordset
    Dim totprod As Double
    Dim Subtotal As Double
    Dim IVA As Double
    Dim Total As Double
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate(App.Path & "\Factura.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.NewPage A4_Vertical
    sBusca = "SELECT ID_CLIENTE FROM VENTAS WHERE ID_VENTA in (" & txtVenta.Text & ")"
    Set tRs = cnn.Execute(sBusca)
    sBusca = "Select * from Cliente Where ID_CLIENTE = " & tRs.Fields("ID_CLIENTE")
    Set tRs = cnn.Execute(sBusca)
    ' cuadros encabezado
    oDoc.WTextBox 100, 5, 20, 250, tRs.Fields("NOMBRE"), "F3", 8, hLeft
    oDoc.WTextBox 120, 5, 20, 250, tRs.Fields("DIRECCION"), "F3", 8, hLeft
    If tRs.Fields("NUMERO_EXTERIOR") <> 0 And Trim(tRs.Fields("NUMERO_EXTERIOR")) <> "" Then oDoc.WTextBox 130, 5, 20, 150, "NO. EXT" & tRs.Fields("NUMERO_EXTERIOR"), "F3", 8, hLeft
    If tRs.Fields("NUMERO_INTERIOR") <> 0 And Trim(tRs.Fields("NUMERO_INTERIOR")) <> "" Then oDoc.WTextBox 130, 112, 20, 150, "INT" & tRs.Fields("NUMERO_INTERIOR"), "F3", 8, hLeft
    oDoc.WTextBox 130, 148, 20, 250, "COLONIA " & tRs.Fields("COLONIA"), "F3", 8, hLeft
    If tRs.Fields("CP") <> 0 Then oDoc.WTextBox 140, 230, 20, 250, tRs.Fields("CP"), "F3", 8, hLeft
    If Not IsNull(tRs.Fields("CIUDAD")) Then oDoc.WTextBox 140, 5, 20, 130, tRs.Fields("CIUDAD"), "F3", 8, hLeft
    If Not IsNull(tRs.Fields("ESTADO")) Then oDoc.WTextBox 140, 180, 20, 100, tRs.Fields("ESTADO"), "F3", 8, hLeft
    oDoc.WTextBox 205, 30, 20, 80, tRs.Fields("RFC"), "F3", 8, hLeft
    ' cuadros encabezado 2
    sBusca = "SELECT * FROM VENTAS WHERE ID_VENTA IN (" & txtVenta.Text & ")"
    Set tRs1 = cnn.Execute(sBusca)
    oDoc.WTextBox 100, 355, 20, 250, "FOLIO :", "F3", 8, hLeft
    oDoc.WTextBox 100, 400, 20, 250, tRs1.Fields("FOLIO"), "F3", 8, hLeft
    oDoc.WTextBox 110, 355, 20, 250, "CLIENTE : ", "F3", 8, hLeft
    oDoc.WTextBox 110, 400, 20, 250, tRs1.Fields("ID_CLIENTE"), "F3", 8, hLeft
    oDoc.WTextBox 120, 355, 20, 280, "AGENTE : ", "F3", 8, hLeft
    txtFolio.Text = tRs1.Fields("FOLIO")
    ' pie de encabezado
    sBusca = "Select * from Ventas Where ID_Venta in (" & txtVenta.Text & ") "
    Set tRs3 = cnn.Execute(sBusca)
    If tRs3.Fields("UNA_EXIBICION") = "S" Then
        oDoc.WTextBox 250, 30, 20, 80, "CONTADO", "F3", 10, hLeft
    Else
        oDoc.WTextBox 250, 120, 20, 120, "CREDITO", "F3", 10, hLeft
        oDoc.WTextBox 250, 30, 20, 80, tRs("DIAS_CREDITO") & " DIAS", "F3", 10, hLeft
    End If
    oDoc.WTextBox 250, 260, 20, 120, VarMen.Text4(3).Text, "F3", 10, hLeft
    If Not IsNull(tRs1.Fields("COMENTARIO")) Then oDoc.WTextBox 205, 120, 20, 180, tRs1.Fields("COMENTARIO"), "F3", 10, hLeft
    If Not IsNull(tRs1.Fields("NOOC")) Then oDoc.WTextBox 205, 270, 20, 190, tRs1.Fields("NOOC"), "F3", 10, hLeft
    oDoc.WTextBox 250, 370, 35, 40, Date, "F3", 8, hLeft
    If Not IsNull(tRs1.Fields("FECHA_VENCE")) Then
        oDoc.WTextBox 205, 505, 35, 40, tRs1.Fields("FECHA_VENCE"), "F3", 8, hLeft
    Else
        If Not IsNull(tRs1.Fields("FECHA_HOY")) Then oDoc.WTextBox 205, 505, 35, 40, tRs1.Fields("FECHA_HOY"), "F3", 8, hLeft
    End If
    Posi = 295
    ' detalle Venta
    totprod = 0
    Subtotal = 0
    IVA = 0
    Total = 0
    Cont = 1
    If txtFolio.Text <> "" Then
        If totprod > 0 Then
            oDoc.WTextBox Posi, 480, 20, 50, totprod, "F3", 7, hRight
            Posi = Posi + 10
        End If
        sBusca = "Select * from Ventas_detalle Where ID_Venta in (" & txtVenta.Text & ")"
        Set tRs2 = cnn.Execute(sBusca)
        If Not (tRs2.BOF And tRs2.EOF) Then
           Do While Not (tRs2.EOF)
                oDoc.WTextBox Posi, 1, 20, 70, tRs2.Fields("ID_PRODUCTO"), "F3", 5, hLeft
                oDoc.WTextBox Posi, 72, 20, 25, tRs2.Fields("CANTIDAD"), "F3", 6, hRight
                oDoc.WTextBox Posi, 110, 20, 260, tRs2.Fields("Descripcion"), "F3", 6, hLeft
                oDoc.WTextBox Posi, 370, 35, 50, Format((tRs2.Fields("PRECIO_VENTA")), "###,###,##0.00"), "F3", 6, hRight
                totprod = Format(CDbl(tRs2.Fields("PRECIO_VENTA")) * CDbl(tRs2.Fields("CANTIDAD")), "###,###,##0.00")
                Subtotal = Format(CDbl(Subtotal) + CDbl(totprod), "###,###,##0.00")
                Posi = Posi + 10
                tRs2.MoveNext
                If totprod <> 0 Then
                    Posi = Posi - 10
                    oDoc.WTextBox Posi, 510, 35, 50, Format(totprod, "###,###,##0.00"), "F3", 6, hRight
                    totprod = 0
                    Posi = Posi + 8
                End If
                Cont = Cont + 1
                If Cont > 24 Then
                    Cont = 1
                    oDoc.NewPage A4_Vertical
                    sBusca = "SELECT ID_CLIENTE FROM VENTAS WHERE ID_VENTA in (" & txtVenta.Text & ")"
                    Set tRs = cnn.Execute(sBusca)
                    sBusca = "Select * from Cliente Where ID_CLIENTE = " & tRs.Fields("ID_CLIENTE")
                    Set tRs = cnn.Execute(sBusca)
                    ' cuadros encabezado
                    oDoc.WTextBox 100, 5, 20, 250, tRs.Fields("NOMBRE"), "F3", 8, hLeft
                    oDoc.WTextBox 120, 5, 20, 250, tRs.Fields("DIRECCION"), "F3", 8, hLeft
                    If tRs.Fields("NUMERO_EXTERIOR") <> 0 And Trim(tRs.Fields("NUMERO_EXTERIOR")) <> "" Then oDoc.WTextBox 130, 5, 20, 150, "NO. EXT" & tRs.Fields("NUMERO_EXTERIOR"), "F3", 8, hLeft
                    If tRs.Fields("NUMERO_INTERIOR") <> 0 And Trim(tRs.Fields("NUMERO_INTERIOR")) <> "" Then oDoc.WTextBox 130, 112, 20, 150, "INT" & tRs.Fields("NUMERO_INTERIOR"), "F3", 8, hLeft
                    oDoc.WTextBox 130, 148, 20, 250, "COLONIA " & tRs.Fields("COLONIA"), "F3", 8, hLeft
                    If tRs.Fields("CP") <> 0 Then oDoc.WTextBox 140, 230, 20, 250, tRs.Fields("CP"), "F3", 8, hLeft
                    If Not IsNull(tRs.Fields("CIUDAD")) Then oDoc.WTextBox 140, 5, 20, 130, tRs.Fields("CIUDAD"), "F3", 8, hLeft
                    If Not IsNull(tRs.Fields("ESTADO")) Then oDoc.WTextBox 140, 180, 20, 100, tRs.Fields("ESTADO"), "F3", 8, hLeft
                    oDoc.WTextBox 205, 30, 20, 80, tRs.Fields("RFC"), "F3", 8, hLeft
                    ' cuadros encabezado 2
                    sBusca = "SELECT * FROM VENTAS WHERE ID_VENTA IN (" & txtVenta.Text & ")"
                    Set tRs1 = cnn.Execute(sBusca)
                    oDoc.WTextBox 100, 355, 20, 250, "FOLIO :", "F3", 8, hLeft
                    oDoc.WTextBox 100, 400, 20, 250, tRs1.Fields("FOLIO"), "F3", 8, hLeft
                    oDoc.WTextBox 110, 355, 20, 250, "CLIENTE : ", "F3", 8, hLeft
                    oDoc.WTextBox 110, 400, 20, 250, tRs1.Fields("ID_CLIENTE"), "F3", 8, hLeft
                    oDoc.WTextBox 120, 355, 20, 280, "AGENTE : ", "F3", 8, hLeft
                    txtFolio.Text = tRs1.Fields("FOLIO")
                    ' pie de encabezado
                    sBusca = "Select * from Ventas Where ID_Venta in (" & txtVenta.Text & ") "
                    Set tRs3 = cnn.Execute(sBusca)
                    If tRs3.Fields("UNA_EXIBICION") = "S" Then
                        oDoc.WTextBox 250, 30, 20, 80, "CONTADO", "F3", 10, hLeft
                    Else
                        oDoc.WTextBox 250, 120, 20, 120, "CREDITO", "F3", 10, hLeft
                        oDoc.WTextBox 250, 30, 20, 80, tRs("DIAS_CREDITO") & " DIAS", "F3", 10, hLeft
                    End If
                    oDoc.WTextBox 250, 260, 20, 120, VarMen.Text4(3).Text, "F3", 10, hLeft
                    If Not IsNull(tRs1.Fields("COMENTARIO")) Then oDoc.WTextBox 205, 120, 20, 180, tRs1.Fields("COMENTARIO"), "F3", 10, hLeft
                    If Not IsNull(tRs1.Fields("NOOC")) Then oDoc.WTextBox 205, 270, 20, 190, tRs1.Fields("NOOC"), "F3", 10, hLeft
                    oDoc.WTextBox 250, 370, 35, 40, Date, "F3", 8, hLeft
                    If Not IsNull(tRs1.Fields("FECHA_VENCE")) Then
                        oDoc.WTextBox 205, 505, 35, 40, tRs1.Fields("FECHA_VENCE"), "F3", 8, hLeft
                    Else
                        oDoc.WTextBox 205, 505, 35, 40, tRs1.Fields("FECHA_HOY"), "F3", 8, hLeft
                    End If
                    Posi = 295
                End If
            Loop
            oDoc.WTextBox 508, 500, 35, 80, Format(Subtotal, "###,###,##0.00"), "F3", 8, hRight
            IVA = Format(CDbl(Subtotal) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
            oDoc.WTextBox 520, 500, 35, 80, Format((IVA), "###,###,##0.00"), "F3", 8, hRight
            Total = Format(CDbl(IVA) + CDbl(Subtotal), "###,###,##0.00")
            oDoc.WTextBox 532, 500, 20, 80, Format((Total), "###,###,##0.00"), "F3", 8, hRight
        End If
    End If
    ' pie de la factura
    oDoc.WTextBox 520, 35, 20, 330, txtTL.Text, "F3", 8, hLeft
    oDoc.WTextBox 560, 235, 20, 320, "LIDER EN RECARGA DE CARTUCHOS, TONER Y TINTA ASI COMO EN VENTA DE CONSUMIBLES ORIGINALES. GRACIAS POR SU PREFERENCIA", "F3", 11, hCenter
    If Not IsNull(tRs1.Fields("UNA_EXIBICION")) Then
        If tRs.Fields("LEYENDAS") = "S" Then
            If tRs1.Fields("UNA_EXIBICION") = "S" Then
                oDoc.WTextBox 610, 235, 20, 320, "PAGO EN UNA SOLA EXHIBICION", "F3", 11, hCenter
            Else
                oDoc.WTextBox 610, 235, 20, 320, "PAGO EN PARCIALIDADES", "F3", 11, hCenter
            End If
        End If
    End If
    If Not IsNull(tRs1.Fields("DIA_PAGARE")) Then oDoc.WTextBox 725, 45, 20, 20, tRs1.Fields("DIA_PAGARE"), "F3", 7, hLeft
    If Not IsNull(tRs1.Fields("MES_PAGARE")) Then oDoc.WTextBox 725, 75, 60, 60, tRs1.Fields("MES_PAGARE"), "F3", 7, hLeft
    If Not IsNull(tRs1.Fields("AÑO_PAGARE")) Then oDoc.WTextBox 725, 145, 20, 30, tRs1.Fields("AÑO_PAGARE"), "F3", 7, hLeft
    If Not IsNull(tRs3.Fields("DIA_HOY")) Then oDoc.WTextBox 760, 445, 20, 20, tRs3.Fields("DIA_HOY"), "F3", 7, hLeft
    If Not IsNull(tRs3.Fields("MES_HOY")) Then oDoc.WTextBox 760, 496, 60, 50, tRs3.Fields("MES_HOY"), "F3", 7, hLeft
    If Not IsNull(tRs3.Fields("AÑO_HOY")) Then oDoc.WTextBox 760, 560, 20, 30, tRs3.Fields("AÑO_HOY"), "F3", 7, hLeft
    If Not IsNull(tRs.Fields("NOMBRE_COMERCIAL")) Then oDoc.WTextBox 770, 85, 20, 280, tRs.Fields("NOMBRE_COMERCIAL"), "F3", 8, hLeft
    If Not IsNull(tRs.Fields("DIRECCION")) Then oDoc.WTextBox 780, 85, 20, 280, tRs.Fields("DIRECCION"), "F3", 8, hLeft
    If Not IsNull(tRs.Fields("CIUDAD")) Then oDoc.WTextBox 800, 85, 20, 280, tRs.Fields("CIUDAD"), "F3", 8, hLeft
    'cierre del reporte
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub FACT_ELEC1()
    Dim sBsucar As String
    Dim tRs As ADODB.Recordset
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim Ruta As String
    Dim sNameFile As String
    Dim sRFCEmpr As String
    ' Agregado de creacion y guardado de cadena de facturacion electronica F_ELECTRO
    ' Agregado el 16 de Nov de 2010 por Armando H Valdez Arras
    sBuscar = "SELECT RFC FROM EMPRESA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        sRFCEmpr = Replace(tRs.Fields("RFC"), "-", "")
        sRFCEmpr = Replace(sRFCEmpr, " ", "")
        sRFCEmpr = Replace(sRFCEmpr, "&", "")
        sRFCEmpr = Replace(sRFCEmpr, "_", "")
        sNameFile = "1" & sRFCEmpr & Format(Date, "mmyyyy")
        sBuscar = "SELECT * FROM F_ELECTRONICA1 WHERE FOLIO = '" & txtFolio.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        Me.CommonDialog1.FileName = sNameFile
        CommonDialog1.DialogTitle = "Guardar como"
        CommonDialog1.Filter = "Archivo de Texto (*.txt) |*.txt|"
        Me.CommonDialog1.ShowSave
        Ruta = Me.CommonDialog1.FileName
        If Not (tRs.EOF And tRs.BOF) Then
            If Ruta <> "" Then
                Do While Not tRs.EOF
                    StrCopi = StrCopi & "|" & tRs.Fields("RFC") & "|"
                    StrCopi = StrCopi & tRs.Fields("DISTINTIVO") & "|"
                    StrCopi = StrCopi & Replace(tRs.Fields("FOLIO"), tRs.Fields("DISTINTIVO"), "") & "|"
                    StrCopi = StrCopi & tRs.Fields("NO_APROBACION") & "|"
                    StrCopi = StrCopi & tRs.Fields("FECHA_FACTURA") & "|"
                    StrCopi = StrCopi & tRs.Fields("TOTAL") & "|"
                    StrCopi = StrCopi & tRs.Fields("IVA") & "|"
                    StrCopi = StrCopi & tRs.Fields("FACTURADO") & "|"
                    StrCopi = StrCopi & tRs.Fields("EFECTO_COMPROBANTE") & "|"
                    StrCopi = StrCopi & tRs.Fields("PEDIMENTO_ADUANA") & "|"
                    StrCopi = StrCopi & tRs.Fields("FECHA_PEDIMENTO") & "|"
                    StrCopi = StrCopi & tRs.Fields("NOMBRE_ADUANA") & "|" & Chr(13)
                    tRs.MoveNext
                Loop
                'archivo TXT
                Dim foo As Integer
                foo = FreeFile
                Open Ruta For Output As #foo
                Print #foo, StrCopi
                Close #foo
            End If
            ShellExecute Me.hWnd, "open", Ruta, "", "", 4
        End If
    End If
End Sub
