VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmShowPediC 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Surtir Venta Programada"
   ClientHeight    =   8055
   ClientLeft      =   495
   ClientTop       =   1230
   ClientWidth     =   10815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9720
      TabIndex        =   28
      Top             =   4320
      Width           =   975
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image14 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmShowPediC.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmShowPediC.frx":030A
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9720
      TabIndex        =   24
      Top             =   5520
      Width           =   975
      Begin VB.Image Command6 
         Height          =   735
         Left            =   120
         MouseIcon       =   "frmShowPediC.frx":1F0C
         MousePointer    =   99  'Custom
         Picture         =   "frmShowPediC.frx":2216
         Top             =   240
         Width           =   720
      End
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
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command7 
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
      Left            =   9720
      Picture         =   "frmShowPediC.frx":3DE8
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9720
      TabIndex        =   22
      Top             =   6720
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
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmShowPediC.frx":67BA
         MousePointer    =   99  'Custom
         Picture         =   "frmShowPediC.frx":6AC4
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmShowPediC.frx":8BA6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DTPicker1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CommonDialog1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtFechaCaptura"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtFechaEntrega"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtNoPed"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtUsuario"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtIDCliente"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command5"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command4"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Command3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Combo1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(5)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Command2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Command8"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text3"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Command9"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      Begin VB.CommandButton Command9 
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
         Left            =   8160
         Picture         =   "frmShowPediC.frx":8BC2
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   720
         TabIndex        =   36
         Top             =   360
         Width           =   7215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "-"
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
         Left            =   480
         Picture         =   "frmShowPediC.frx":B594
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   7320
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
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
         Picture         =   "frmShowPediC.frx":DF66
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   7320
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2040
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   7440
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6960
         TabIndex        =   31
         Top             =   840
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Automatico"
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
         Left            =   3360
         Picture         =   "frmShowPediC.frx":10938
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   7440
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   7440
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   7440
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Deshacer"
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
         Left            =   6360
         Picture         =   "frmShowPediC.frx":1330A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   7440
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   7440
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Enviar"
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
         Left            =   7920
         Picture         =   "frmShowPediC.frx":15CDC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7320
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Surtir Producto"
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
         Left            =   4800
         Picture         =   "frmShowPediC.frx":186AE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7320
         Width           =   1455
      End
      Begin VB.TextBox txtIDCliente 
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   7320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtUsuario 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   7320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtNoPed 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   7320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtFechaEntrega 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   7320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtFechaCaptura 
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Top             =   7320
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1440
         Top             =   7320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         PrinterDefault  =   0   'False
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   7320
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52101121
         CurrentDate     =   39127
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   4320
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5106
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   120
         TabIndex        =   0
         Top             =   1320
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5106
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
      Begin VB.Label Label5 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Sucursal de existencias"
         Height          =   255
         Left            =   4920
         TabIndex        =   30
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   3360
         Width           =   9255
      End
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   2415
      Left            =   9720
      TabIndex        =   8
      Top             =   1800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   10080
      TabIndex        =   27
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Clave"
      Height          =   255
      Left            =   9720
      TabIndex        =   26
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmShowPediC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim SelMod As String
Dim NoPed As Integer
Dim CantPed As Double
Dim CantSurt As Double
Dim NoOrdenCli As String
Dim sRastreaProd As String
Dim sRastreaNoPed As String
Private Sub Command1_Click()
    Dim Cont As Integer
    If Combo1.Text <> "" Then
        For Cont = 1 To ListView2.ListItems.Count
            If ListView2.ListItems.Item(Cont).Checked Then
                Text1(0).Text = ListView2.ListItems.Item(Cont)
                Text1(1).Text = ListView2.ListItems.Item(Cont).SubItems(4) 'CantPed
                Text1(2).Text = NoPed
                Text1(5).Text = ListView2.ListItems.Item(Cont).SubItems(1)
                FrmModSurt.Show vbModal
                Command4.Enabled = False
            End If
        Next Cont
        Actualizar
    Else
        MsgBox "Seleccione una sucursal", vbExclamation, "SACC"
    End If
    sRastreaNoPed = ""
    txtIDCliente.Text = ""
    TxtUsuario.Text = ""
    TxtNoPed.Text = ""
    txtFechaEntrega.Text = ""
    txtFechaCaptura.Text = ""
    NoOrdenCli = ""
    Label2.Caption = ""
End Sub
Private Sub Command2_Click()
    Dim Cont As Integer
    For Cont = 1 To ListView2.ListItems.Count
        ListView2.ListItems.Item(Cont).Checked = True
    Next Cont
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    Dim Pend As Double
    Dim NuevaExis As Double
    Dim sqlComanda As String
    Dim tRs As ADODB.Recordset
    Dim NoPed As Integer
    Dim NumeroRegistros As Integer
    NumeroRegistros = ListView1.ListItems.Count
    Dim Conta As Integer
    For Conta = 1 To NumeroRegistros
        NoPed = ListView1.ListItems(Conta)
        sqlComanda = "SELECT ID_PRODUCTO FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO = " & NoPed & " AND CANTIDAD_PENDIENTE > 0"
        Set tRs = cnn.Execute(sqlComanda)
        If (tRs.BOF And tRs.EOF) Then
            sqlComanda = "UPDATE PED_CLIEN SET ESTADO = 'C' WHERE NO_PEDIDO = " & NoPed
            Set tRs = cnn.Execute(sqlComanda)
            Actualizar
        Else
        ' AQUI VA EL MERO MOLE... LO SABROSO DE ESTO.... EL CODIGO QUE HACE TODO EL MOVIMIETO... WACHA!!!
            Dim NoReg As Integer
            NoReg = ListView2.ListItems.Count
            Dim Con As Integer
            Dim IDPro As String
            For Con = 1 To NoReg
                IDPro = ListView2.ListItems(Con)
                sqlComanda = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & IDPro & "' AND CANTIDAD >= " & ListView2.ListItems(Con).SubItems(3) & " AND SUCURSAL = '" & Combo1.Text & "'"
                Set tRs = cnn.Execute(sqlComanda)
                If Not (tRs.BOF And tRs.EOF) Then
                    sqlComanda = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs.Fields("CANTIDAD")) - CDbl(ListView2.ListItems(Con).SubItems(3)) & " WHERE ID_PRODUCTO = '" & ListView2.ListItems(Con) & "' AND SUCURSAL = '" & Combo1.Text & "'"
                    Set tRs = cnn.Execute(sqlComanda)
                    sqlComanda = "UPDATE PED_CLIEN_DETALLE SET CANTIDAD_PENDIENTE = 0 WHERE ID_PRODUCTO = '" & ListView2.ListItems(Con) & "' AND NO_PEDIDO = " & NoPed
                    Set tRs = cnn.Execute(sqlComanda)
                End If
            Next Con
        ' LISTO!!!
        End If
    Next Conta
    Actualizar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command4_Click()
    Dim Cont As Integer
    If Combo1.Text <> "" Then
        For Cont = 1 To ListView2.ListItems.Count
            If ListView2.ListItems.Item(Cont).Checked Then
                Text1(4).Text = ListView2.ListItems.Item(Cont).SubItems(2)  'CantSurt
                CantSurt = Val(Replace(ListView2.ListItems.Item(Cont).SubItems(2), ",", ""))
                CantPed = Val(Replace(ListView2.ListItems.Item(Cont).SubItems(4), ",", ""))
                CantSurt = CantSurt - CantPed
                Text1(0).Text = ListView2.ListItems.Item(Cont)
                Text1(1).Text = CantSurt
                Text1(2).Text = NoPed
                Text1(3).Text = CantPed
                Label3.Caption = NoOrdenCli
                DTPicker1.Value = ListView2.ListItems.Item(Cont).SubItems(5)
                FrmDeshacer.Show vbModal, Me
                Command4.Enabled = False
            End If
        Next Cont
        Actualizar
    Else
        MsgBox "Seleccione una sucursal", vbExclamation, "SACC"
    End If
End Sub
Private Sub Command5_Click()
    Dim sBuscar As String
    Dim sBuscar2 As String
    Dim Cont As Integer
    Dim tRs As ADODB.Recordset
    Dim CantO As Double
    Dim Ped As Double
    Dim NoPedido As Integer
    Dim completo As Boolean
    NoPedido = -1
    Cont = 1
    completo = True
    Do While Not (Cont > ListView2.ListItems.Count) And completo
        If Val(ListView2.ListItems.Item(Cont).SubItems(4)) <> 0 Then completo = False
        Cont = Cont + 1
    Loop
    If completo Then
        sBuscar2 = "UPDATE PED_CLIEN SET ESTADO = 'C', FECHA_CIERRE = DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())) WHERE NO_PEDIDO = " & NoPed
        cnn.Execute (sBuscar2)
        sBuscar2 = "UPDATE PED_CLIEN_DETALLE SET ACTIVO = 'N' WHERE NO_PEDIDO = " & NoPed
        cnn.Execute (sBuscar2)
    Else
        For Cont = 1 To ListView2.ListItems.Count
            If ListView2.ListItems.Item(Cont).Checked Then
                If NoPedido = -1 Then
                    sBuscar = "INSERT INTO PED_CLIEN (ID_CLIENTE, USUARIO, FECHA, ESTADO, NO_ORDEN, FECHA_CIERRE) VALUES (" & txtIDCliente.Text & ", '" & VarMen.Text1(0).Text & "', '" & Format(Date, "dd/mm/yyyy") & "', 'C', '" & NoOrdenCli & "', DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())));"
                    cnn.Execute (sBuscar)
                    sBuscar = "SELECT MAX(NO_PEDIDO) AS NO_PEDIDO FROM PED_CLIEN"
                    Set tRs = cnn.Execute(sBuscar)
                    NoPedido = tRs.Fields("NO_PEDIDO")
                End If
                sBuscar = "SELECT CANTIDAD_PEDIDA, CANT_ORIGINAL, CANTIDAD_PENDIENTE FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO = " & NoPed & " AND ID_PRODUCTO = '" & ListView2.ListItems.Item(Cont) & "'"
                Set tRs = cnn.Execute(sBuscar)
                If tRs.Fields("CANTIDAD_PENDIENTE") = 0 Then
                    ' INTENTO POR ELIMINAR LOS ENVIOS PARCIALES
                    sBuscar = "DELETE FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO  = " & NoPed & " AND ID_PRODUCTO = '" & ListView2.ListItems.Item(Cont) & "'"
                    cnn.Execute (sBuscar)
                    sBuscar = "INSERT INTO PED_CLIEN_DETALLE(ID_PRODUCTO, NO_PEDIDO, CANTIDAD_PEDIDA, CANTIDAD_EXISTENCIA, CANTIDAD_PENDIENTE) VALUES "
                    sBuscar = sBuscar & "('" & ListView2.ListItems.Item(Cont) & "', " & NoPedido & ", " & ListView2.ListItems.Item(Cont).SubItems(2)
                    sBuscar = sBuscar & ", " & ListView2.ListItems.Item(Cont).SubItems(3) & ", 0);"
                    sBuscar2 = "UPDATE PED_CLIEN_DETALLE SET ACTIVO = 'N' WHERE NO_PEDIDO = " & NoPedido & " AND ID_PRODUCTO = '" & ListView2.ListItems.Item(Cont) & "'"
                Else
                    If IsNull(tRs.Fields("CANT_ORIGINAL")) Then
                        CantO = tRs.Fields("CANTIDAD_PEDIDA")
                    Else
                        CantO = tRs.Fields("CANT_ORIGINAL")
                    End If
                    sBuscar = "INSERT INTO PED_CLIEN_DETALLE(ID_PRODUCTO, NO_PEDIDO, CANTIDAD_PEDIDA, CANTIDAD_EXISTENCIA, CANTIDAD_PENDIENTE) VALUES "
                    Ped = Abs(Val(ListView2.ListItems.Item(Cont).SubItems(2)) - Val(ListView2.ListItems.Item(Cont).SubItems(4)))
                    sBuscar = sBuscar & "('" & ListView2.ListItems.Item(Cont) & "', " & NoPedido & ", " & Ped
                    sBuscar = sBuscar & ", " & ListView2.ListItems.Item(Cont).SubItems(3) & ", 0);"
                    sBuscar2 = "UPDATE PED_CLIEN_DETALLE SET CANTIDAD_PEDIDA = " & ListView2.ListItems.Item(Cont).SubItems(4) & ",CANT_ORIGINAL = " & ListView2.ListItems.Item(Cont).SubItems(2) & ", CANTIDAD_PENDIENTE = " & ListView2.ListItems.Item(Cont).SubItems(4) & " WHERE NO_PEDIDO = " & NoPed & " AND ID_PRODUCTO = '" & ListView2.ListItems.Item(Cont) & "'"
                End If
                cnn.Execute (sBuscar)
                cnn.Execute (sBuscar2)
            End If
        Next Cont
    End If
    MsgBox "La venta se cerro con el folio " & NoPedido & ", Favor de notificar a ventas!", vbInformation, "SACC"
    Actualizar
End Sub
Private Sub Command6_Click()
On Error GoTo Error1
    Dim NRegistros As Integer
    Dim POSY As Integer
    Dim Con As Integer
    Dim Hojas As Integer
    Dim Hojastot As Integer
    Dim HojasAprox As Double
    If ListView2.ListItems.Count > 0 Then
        Hojas = 1
        Hojastot = ListView2.ListItems.Count
        HojasAprox = Hojastot / 30
        Hojastot = Hojastot / 30
        If HojasAprox - Hojastot > 0 Then
            Hojastot = Hojastot + 1
        End If
        CommonDialog1.Flags = 64
        CommonDialog1.CancelError = True
        CommonDialog1.ShowPrinter
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("SURTIR PEDIDO CLIENTE")) / 2
        Printer.Print "SURTIR PEDIDO CLIENTE"
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("LISTA DE PRODUCTOS")) / 2
        Printer.Print "LISTA DE PRODUCTOS"
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("FECHA DE ENTRGA: " & txtFechaEntrega.Text & "         FECHA DE CAPTURA: " & txtFechaCaptura.Text)) / 2
        Printer.Print "FECHA DE ENTRGA: " & txtFechaEntrega.Text & "         FECHA DE CAPTURA: " & txtFechaCaptura.Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("PEDIDO : " & TxtNoPed.Text)) / 2
        Printer.Print "PEDIDO : " & TxtNoPed.Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("CLIENTE : " & Label2.Caption)) / 2
        Printer.Print "CLIENTE : " & Label2.Caption
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        NRegistros = ListView2.ListItems.Count
        POSY = 2200
        Printer.CurrentY = POSY
        Printer.CurrentX = 400
        Printer.Print "PRODUCTO"
        Printer.CurrentY = POSY
        Printer.CurrentX = 2500
        Printer.Print "CANTIDAD PEDIDA"
        Printer.CurrentY = POSY
        Printer.CurrentX = 4500
        Printer.Print "CANTIDAD SURTIDA"
        Printer.CurrentY = POSY
        Printer.CurrentX = 6500
        Printer.Print "CANTIDAD EXISTENCIA"
        Printer.CurrentY = POSY
        Printer.CurrentX = 9000
        Printer.Print "CANTIDAD PENDIENTE"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        POSY = POSY + 200
        For Con = 1 To NRegistros
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 400
            Printer.Print ListView2.ListItems(Con)
            Printer.CurrentY = POSY
            Printer.CurrentX = (3500 - Printer.TextWidth(ListView2.ListItems(Con).SubItems(2)))
            Printer.Print ListView2.ListItems(Con).SubItems(2)
            Printer.CurrentY = POSY
            Printer.CurrentX = (5500 - Printer.TextWidth(ListView2.ListItems(Con).SubItems(3)))
            Printer.Print CDbl(ListView2.ListItems(Con).SubItems(2)) - CDbl(ListView2.ListItems(Con).SubItems(4))
            Printer.CurrentY = POSY
            Printer.CurrentX = (7500 - Printer.TextWidth(ListView2.ListItems(Con).SubItems(4)))
            Printer.Print ListView2.ListItems(Con).SubItems(3)
            Printer.CurrentY = POSY
            Printer.CurrentX = (10000 - Printer.TextWidth(ListView2.ListItems(Con).SubItems(4)))
            Printer.Print ListView2.ListItems(Con).SubItems(4)
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 0
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            If POSY >= 14200 Then
                Printer.CurrentX = (Printer.Width - 900 - Printer.TextWidth("Pagina " & Hojas & " de " & Hojastot))
                Printer.Print "Pagina " & Hojas & " de " & Hojastot
                Printer.NewPage
                Hojas = Hojas + 1
                POSY = 100
                Printer.Print ""
                Printer.Print ""
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                Printer.Print VarMen.Text5(0).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("SURTIR PEDIDO CLIENTE")) / 2
                Printer.Print "SURTIR PEDIDO CLIENTE"
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("LISTA DE PRODUCTOS")) / 2
                Printer.Print "LISTA DE PRODUCTOS"
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("FECHA DE ENTRGA: " & txtFechaEntrega.Text & "         FECHA DE CAPTURA: " & txtFechaCaptura.Text)) / 2
                Printer.Print "FECHA DE ENTRGA: " & txtFechaEntrega.Text & "         FECHA DE CAPTURA: " & txtFechaCaptura.Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("PEDIDO : " & TxtNoPed.Text)) / 2
                Printer.Print "PEDIDO : " & TxtNoPed.Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("CLIENTE : " & Label2.Caption)) / 2
                Printer.Print "CLIENTE : " & Label2.Caption
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                NRegistros = ListView2.ListItems.Count
                POSY = 2200
                Printer.CurrentY = POSY
                Printer.CurrentX = 400
                Printer.Print "PRODUCTO"
                Printer.CurrentY = POSY
                Printer.CurrentX = 2500
                Printer.Print "CANTIDAD PEDIDA"
                Printer.CurrentY = POSY
                Printer.CurrentX = 4500
                Printer.Print "CANTIDAD SURTIDA"
                Printer.CurrentY = POSY
                Printer.CurrentX = 6500
                Printer.Print "CANTIDAD EXISTENCIA"
                Printer.CurrentY = POSY
                Printer.CurrentX = 6500
                Printer.Print "CANTIDAD PENDIENTE"
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                POSY = POSY + 200
            End If
        Next Con
        Printer.CurrentX = (Printer.Width - 900 - Printer.TextWidth("Pagina " & Hojas & " de " & Hojastot))
        Printer.Print "Pagina " & Hojas & " de " & Hojastot
        Printer.Print "FIN DEL LISTADO"
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.EndDoc
        CommonDialog1.Copies = 1
    End If
Exit Sub
Error1:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub Command7_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If ListView1.ListItems.Count > 0 Then
        If Text2.Text <> "" Then
            sBuscar = "SELECT D.NO_PEDIDO FROM PED_CLIEN_DETALLE AS D JOIN PED_CLIEN AS P ON D.NO_PEDIDO = P.NO_PEDIDO WHERE P.ESTADO = 'I' AND ID_PRODUCTO LIKE '%" & Text2.Text & "%' ORDER BY D.NO_PEDIDO"
            Set tRs = cnn.Execute(sBuscar)
            If (tRs.BOF And tRs.EOF) Then
                MsgBox "EL PRODUCTO NO ESTA EN NINGUN PEDIDO", vbInformation, "SACC"
                ListView3.ListItems.Clear
            Else
                ListView3.ListItems.Clear
                tRs.MoveFirst
                Do While Not tRs.EOF
                    Set tLi = ListView3.ListItems.Add(, , tRs.Fields("NO_PEDIDO"))
                    tRs.MoveNext
                Loop
            End If
        End If
    End If
End Sub
Private Sub Command8_Click()
    Dim Cont As Integer
    For Cont = 1 To ListView2.ListItems.Count
        ListView2.ListItems.Item(Cont).Checked = False
    Next Cont
End Sub
Private Sub Command9_Click()
    Actualizar
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
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
        .ColumnHeaders.Add , , "No. Pedido", 1000
        .ColumnHeaders.Add , , "Capturo", 1500
        .ColumnHeaders.Add , , "Cliente", 6500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "IDCliente", 0
        .ColumnHeaders.Add , , "Fecha Captura", 1500
        .ColumnHeaders.Add , , "No. de Orden", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Producto", 2500
        .ColumnHeaders.Add , , "Descripcion", 7000
        .ColumnHeaders.Add , , "Cantidad Pedida", 2000
        .ColumnHeaders.Add , , "Cantidad en Existencia", 2000
        .ColumnHeaders.Add , , "Cantidad Pendiente", 2000
        .ColumnHeaders.Add , , "Fecha", 0
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "#Venta", 1000
    End With
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' GROUP BY NOMBRE ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            If Trim(tRs.Fields("NOMBRE")) <> "" Then Combo1.AddItem (tRs.Fields("NOMBRE"))
            tRs.MoveNext
        Loop
    End If
    Actualizar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image14_Click()
    FrmMuestraProgramadas.Show vbModal
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    NoPed = Item
    Dim sBuscar As String
    Dim tRs3 As ADODB.Recordset
    Dim tLi As ListItem
    sRastreaNoPed = Item
    txtIDCliente.Text = Item.SubItems(4)
    TxtUsuario.Text = Item.SubItems(1)
    TxtNoPed.Text = Item
    txtFechaEntrega.Text = Item.SubItems(3)
    txtFechaCaptura.Text = Item.SubItems(5)
    NoOrdenCli = Item.SubItems(6)
    Label2.Caption = Item.SubItems(2)
    'sBuscar = "SELECT I.ID_PRODUCTO, ALMACEN3.DESCRIPCION, I.NO_PEDIDO, I.CANTIDAD_PEDIDA, I.CANTIDAD_PENDIENTE, ISNULL(I.CANT_ORIGINAL, I.CANTIDAD_PEDIDA) AS CANT_ORIGINAL, ISNULL ((SELECT CANTIDAD From EXISTENCIAS WHERE (ID_PRODUCTO = I.ID_PRODUCTO) AND (SUCURSAL = '" & Combo1.Text & "')), 0) AS CANTIDAD FROM PED_CLIEN_DETALLE AS I INNER JOIN ALMACEN3 ON I.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO WHERE (I.ACTIVO = 'S') AND (I.NO_PEDIDO = " & Item & ")"
    sBuscar = "SELECT I.ID_PRODUCTO, ALMACEN3.DESCRIPCION, I.NO_PEDIDO, I.CANTIDAD_PEDIDA, I.CANTIDAD_PENDIENTE, ISNULL(I.CANT_ORIGINAL, I.CANTIDAD_PEDIDA) AS CANT_ORIGINAL, ISNULL ((SELECT CANTIDAD From EXISTENCIAS WHERE (ID_PRODUCTO = I.ID_PRODUCTO) AND (SUCURSAL = '" & Combo1.Text & "')), 0) AS CANTIDAD FROM PED_CLIEN_DETALLE AS I INNER JOIN ALMACEN3 ON I.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO WHERE (I.ACTIVO = 'S') AND (I.NO_PEDIDO = " & Item & ")"
    Set tRs3 = cnn.Execute(sBuscar)
    If (tRs3.BOF And tRs3.EOF) Then
        Me.ListView2.ListItems.Clear
    Else
        ListView2.ListItems.Clear
        Do While Not tRs3.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs3.Fields("ID_PRODUCTO"))
            tLi.SubItems(1) = tRs3.Fields("DESCRIPCION")
            tLi.SubItems(2) = tRs3.Fields("CANTIDAD_PEDIDA")
            tLi.SubItems(3) = tRs3.Fields("CANTIDAD")
            tLi.SubItems(4) = tRs3.Fields("CANTIDAD_PENDIENTE")
            tLi.SubItems(5) = Item.SubItems(3)
            tRs3.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView2_DblClick()
    FrmRastreoVentProg.Text2.Text = sRastreaProd
    FrmRastreoVentProg.Text1.Text = sRastreaNoPed
    FrmRastreoVentProg.Show vbModal
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Me.Command4.Enabled = True
        CantPed = Item.SubItems(4)
        CantSurt = Item.SubItems(2)
    Else
        Me.Command4.Enabled = False
        CantPed = 0
        CantSurt = 0
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Item.Checked = Not (Item.Checked)
    sRastreaProd = Item
    If Item.Checked Then
        Me.Command4.Enabled = True
    Else
        Me.Command4.Enabled = False
    End If
End Sub
Private Sub Actualizar()
On Error GoTo ManejaError
    Me.Command4.Enabled = False
    Dim sBuscar As String
    Dim BusClie As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT PED_CLIEN.NO_PEDIDO, PED_CLIEN.ID_CLIENTE, PED_CLIEN.USUARIO, PED_CLIEN.FECHA, PED_CLIEN.ESTADO, PED_CLIEN.FECHA_CAPTURA, PED_CLIEN.NO_ORDEN, PED_CLIEN.COMENTARIO, PED_CLIEN.FECHA_CIERRE, PED_CLIEN.FECHA_FACTURACION, CLIENTE.NOMBRE, CLIENTE.NOMBRE_COMERCIAL, USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS USUARIO FROM PED_CLIEN INNER JOIN CLIENTE ON PED_CLIEN.ID_CLIENTE = CLIENTE.ID_CLIENTE INNER JOIN USUARIOS ON PED_CLIEN.USUARIO = USUARIOS.ID_USUARIO WHERE (PED_CLIEN.ESTADO = 'I') AND (CLIENTE.NOMBRE LIKE '%" & Text3.Text & "%') ORDER BY PED_CLIEN.NO_PEDIDO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("NO_PEDIDO"))
                If Not IsNull(.Fields("USUARIO")) Then tLi.SubItems(1) = .Fields("USUARIO")
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(2) = .Fields("NOMBRE")
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(3) = .Fields("FECHA")
                If Not IsNull(.Fields("ID_CLIENTE")) Then tLi.SubItems(4) = .Fields("ID_CLIENTE")
                If Not IsNull(.Fields("FECHA_CAPTURA")) Then tLi.SubItems(5) = .Fields("FECHA_CAPTURA")
                If Not IsNull(.Fields("NO_ORDEN")) Then tLi.SubItems(6) = .Fields("NO_ORDEN")
                .MoveNext
            Loop
        End If
    End With
    ListView2.ListItems.Clear
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command7.Value = True
    End If
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Actualizar
    End If
End Sub
