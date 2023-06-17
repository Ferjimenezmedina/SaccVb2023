VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCancelaFactura 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelar Factura / Nota de Venta"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "UUID"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   34
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   360
      TabIndex        =   31
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   3360
      TabIndex        =   10
      Top             =   120
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Cancelar"
      TabPicture(0)   =   "frmCancelaFactura.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTitulo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblNoMov"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView2(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView2(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtIDVENTA"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtNo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Buscar"
      TabPicture(1)   =   "frmCancelaFactura.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text4"
      Tab(1).Control(1)=   "Command4"
      Tab(1).Control(2)=   "ListView4"
      Tab(1).Control(3)=   "ListView3"
      Tab(1).Control(4)=   "Option4"
      Tab(1).Control(5)=   "Option3"
      Tab(1).Control(6)=   "Text2"
      Tab(1).Control(7)=   "Label1"
      Tab(1).ControlCount=   8
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Top             =   6120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   -68160
         TabIndex        =   30
         Text            =   "Text4"
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command4 
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
         Left            =   -70200
         Picture         =   "frmCancelaFactura.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   28
         Top             =   3720
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   27
         Top             =   1080
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Factura"
         Height          =   195
         Left            =   -71520
         TabIndex        =   26
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "No. de Venta"
         Height          =   195
         Left            =   -71520
         TabIndex        =   25
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -73320
         TabIndex        =   24
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   5760
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Movimiento"
         Height          =   735
         Left            =   3840
         TabIndex        =   16
         Top             =   960
         Width           =   2895
         Begin VB.OptionButton Option1 
            Caption         =   "Notas de Venta"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   18
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Facturas"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   6135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
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
         Left            =   2640
         Picture         =   "frmCancelaFactura.frx":2A0A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5760
         Width           =   1455
      End
      Begin VB.TextBox txtIDVENTA 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   13
         Top             =   5760
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3855
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Visible         =   0   'False
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   6800
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3855
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6800
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Factura Nueva"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   6120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Numero de Venta :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblNoMov 
         Caption         =   "Factura No:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Facturas Pendientes de Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2280
      TabIndex        =   8
      Top             =   5400
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmCancelaFactura.frx":53DC
         MousePointer    =   99  'Custom
         Picture         =   "frmCancelaFactura.frx":56E6
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
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
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
      Picture         =   "frmCancelaFactura.frx":77C8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   3135
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Notas de Venta"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Facturas"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar Factura"
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
      Left            =   120
      Picture         =   "frmCancelaFactura.frx":A19A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblTipoMov 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Numero de Factura:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
End
Attribute VB_Name = "frmCancelaFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ClvCliente As Integer
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    If Option2(0).Value Then
        sBuscar = "SELECT ID_CLIENTE, ID_VENTA, FOLIO, UNA_EXIBICION TOTAL FROM VENTAS WHERE FOLIO ='" & Text14.Text & "'" '& " AND ID_CUENTA = 0 "
        Label2.Visible = True
        Text6.Visible = True
    Else
        If Option2(1).Value Then
            sBuscar = "SELECT ID_CLIENTE, ID_VENTA, FOLIO, UNA_EXIBICION TOTAL FROM VENTAS WHERE ID_VENTA = " & Val(Text14.Text)
        Else
            sBuscar = "SELECT ID_CLIENTE, ID_VENTA, FOLIO, UNA_EXIBICION TOTAL FROM VENTAS WHERE UUID LIKE '" & Val(Text14.Text) & "'"
        End If
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        ClvCliente = tRs.Fields("ID_CLIENTE")
        Actual (tRs.Fields("ID_CLIENTE"))
        If Command1.Caption = "Buscar Factura" Then
            Option1(0).Value = True
            txtNo.Text = tRs.Fields("FOLIO")
            lblNoMov.Caption = "Factura No:"
            txtIDVENTA.Text = ""
            sBuscar = "SELECT ID_VENTA FROM VENTAS WHERE FOLIO = '" & txtNo.Text & "'"
            Set tRs1 = cnn.Execute(sBuscar)
            Do While Not tRs1.EOF
                If txtIDVENTA.Text = "" Then
                    txtIDVENTA.Text = tRs1.Fields("ID_VENTA")
                Else
                    txtIDVENTA.Text = txtIDVENTA.Text & "," & tRs1.Fields("ID_VENTA")
                End If
                tRs1.MoveNext
            Loop
        Else
            Option1(1).Value = True
            txtNo.Text = tRs.Fields("ID_VENTA")
            lblNoMov.Caption = "Nota V. No:"
            txtIDVENTA.Text = txtNo.Text
        End If
        sBuscar = "SELECT NOMBRE FROM CLIENTE WHERE ID_CLIENTE = " & tRs.Fields("ID_CLIENTE")
        Set tRs = cnn.Execute(sBuscar)
        If Not IsNull(tRs.Fields("NOMBRE")) Then Text1.Text = tRs.Fields("NOMBRE")
        If Not IsNull(tRs.Fields("NOMBRE")) Then Text3.Text = tRs.Fields("NOMBRE")
    Else
        MsgBox "El movimiento no puede cancelarse, o no esta capturado"
    End If
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    Buscar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command3_Click()
    Dim sBusca As String
    Dim sBuscar As String
    Dim Id_vent As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tRs5 As ADODB.Recordset
    Dim num As String
    Dim TCant As Double
    Dim CanTodo As Boolean
    Dim OPERA As Double
    Dim Sucursal As String
    Dim IdProd As String
    Dim Canti As Double
    Dim ClvProd As String
    txtIDVENTA.Text = Replace(txtIDVENTA.Text, ",", " OR ID_VENTA = ")
    If Option2(0).Value Then
        sBusca = "SELECT ID_VENTA, SUCURSAL FROM VENTAS WHERE FOLIO = '" & Text14.Text & "' ORDER BY FOLIO"
        Set tRs3 = cnn.Execute(sBusca)
        If Not (tRs3.EOF And tRs3.BOF) Then
            If Not IsNull(tRs3.Fields("ID_VENTA")) Then Id_vent = tRs3.Fields("ID_VENTA")
            If Not IsNull(tRs3.Fields("SUCURSAL")) Then Sucursal = tRs3.Fields("SUCURSAL")
        Else
            MsgBox "EL FOLIO DE FACTURA NO EXISTE O YA FUE CANCELADA!", vbExclamation, "SACC"
            Exit Sub
        End If
    Else
        sBusca = "SELECT ID_VENTA, SUCURSAL FROM VENTAS WHERE ID_VENTA = '" & Text14.Text & "' ORDER BY ID_VENTA"
        Set tRs3 = cnn.Execute(sBusca)
        If Not IsNull(tRs3.Fields("SUCURSAL")) Then Sucursal = tRs3.Fields("SUCURSAL")
        Id_vent = Text14.Text
    End If
    If lblTipoMov.Caption = "Numero de Factura:" Then
        sBuscar = "SELECT ID_VENTA FROM VENTAS WHERE FOLIO = '" & txtNo.Text & "' AND UNA_EXIBICION = 'V'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While tRs.EOF
                sBuscar = "SELECT ID_CUENTA FROM CUENTA_VENTA WHERE ID_VENTA = " & txtIDVENTA
                Set tRs2 = cnn.Execute(sBuscar)
                If Not (tRs2.EOF And tRs2.BOF) Then
                    Do While Not tRs2.EOF
                        sBuscar = "DELETE FROM CUENTAS WHERE ID_CUENTA = " & tRs2.Fields("ID_CUENTA")
                        cnn.Execute (sBuscar)
                        tRs2.MoveNext
                    Loop
                End If
                tRs.MoveNext
            Loop
        End If
        If MsgBox("¿DESEA DEJAR LA VENTA ABIERTA PARA RE-FACTURARLA?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
            CanTodo = False
            sBusca = "UPDATE VENTAS SET FACTURADO = '0', FOLIO = '', ID_USUARIO_CANCELA = '" & VarMen.Text1(0).Text & "' WHERE FOLIO = '" & txtNo.Text & "'"
            cnn.Execute (sBusca)
        Else
            'SI LA FACTURA ES CANCELADA POR COMPLETO ENTONCES ELIMINA EL CREDITO
            CanTodo = True
            sBuscar = "SELECT ID_CUENTA FROM CUENTA_VENTA WHERE ID_VENTA = " & txtIDVENTA.Text
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                sBuscar = "DELETE FROM CUENTAS WHERE ID_CUENTA = " & tRs.Fields("ID_CUENTA")
                cnn.Execute (sBuscar)
            End If
            sBusca = "UPDATE VENTAS SET FACTURADO = '2', FOLIO = 'CANCELADO', ID_USUARIO_CANCELA = '" & VarMen.Text1(0).Text & "' WHERE FOLIO = '" & txtNo.Text & "'"
            cnn.Execute (sBusca)
        End If
        If Command1.Caption = "Buscar Factura" Then
            If InStr(1, txtIDVENTA, "OR") = 0 Then
                sBusca = "INSERT INTO FACTCAN (ID_VENTA, FOLIO, FECHA, ID_USUARIO, ID_CLIENTE, PASA_FACTU, NOMBRE) VALUES('" & txtIDVENTA & "', '" & txtNo.Text & "', '" & Format(Date, "dd/mm/yyyy") & "', '" & Text5.Text & "', '" & ClvCliente & "', '" & Text6.Text & "', '" & Text3.Text & "');"
                cnn.Execute (sBusca)
            End If
        Else
            sBusca = "INSERT INTO FACTCAN (ID_VENTA, FOLIO, FECHA, ID_USUARIO, ID_CLIENTE, NOMBRE) VALUES('" & Id_vent & "', '" & txtNo.Text & "', '" & Format(Date, "dd/mm/yyyy") & "','" & Text5.Text & "','" & ClvCliente & "','" & Text3.Text & "');"
            cnn.Execute (sBusca)
       End If
    Else
        sBusca = "SELECT ID_VENTA FROM VENTAS WHERE ID_VENTA = " & txtIDVENTA.Text & " AND FACTURADO <> 2"
        Set tRs = cnn.Execute(sBusca)
        If Not (tRs.EOF And tRs.BOF) Then
            sBusca = "UPDATE VENTAS SET FACTURADO = '2', FOLIO = 'CANCELADO', ID_USUARIO_CANCELA = '" & VarMen.Text1(0).Text & "' WHERE ID_VENTA = " & txtIDVENTA.Text
            cnn.Execute (sBusca)
        Else
            MsgBox "LA NOTA YA FUE CANCELADA ANTERIORMENTE!", vbInformation, "SACC"
            Exit Sub
        End If
    End If
    If Option2(0).Value = True Then
        sBusca = "SELECT * FROM VENTAS_DETALLE WHERE ID_VENTA = " & txtIDVENTA.Text
    Else
        sBusca = "SELECT * FROM VENTAS_DETALLE WHERE ID_VENTA = " & txtIDVENTA.Text
        CanTodo = True
    End If
    Set tRs = cnn.Execute(sBusca)
    If Not (tRs.EOF And tRs.BOF) And CanTodo Then
        Dim IdVenta As Integer
        'IdVenta = txtIDVENTA.Text
        sBusca = "SELECT * FROM VENTAS_DETALLE WHERE ID_VENTA IN (" & Replace(txtIDVENTA.Text, "OR ID_VENTA =", ", ") & ")"
        Set tRs = cnn.Execute(sBusca)
        If Option2(0).Value = False Then
            sBuscar = "SELECT ID_CUENTA FROM CUENTA_VENTA WHERE ID_VENTA IN (" & Replace(txtIDVENTA.Text, "OR ID_VENTA =", ", ") & ")"
            Set tRs4 = cnn.Execute(sBuscar)
            If Not (tRs4.EOF And tRs4.BOF) Then
                sBuscar = "DELETE FROM CUENTAS WHERE ID_CUENTA IN (" & "SELECT ID_CUENTA FROM CUENTA_VENTA WHERE ID_VENTA IN (" & Replace(txtIDVENTA.Text, "OR ID_VENTA =", ", ") & ")" & ")"
                cnn.Execute (sBuscar)
            End If
        End If
        If CanTodo Then
            sBusca = "SELECT SUCURSAL FROM VENTAS WHERE ID_VENTA = " & txtIDVENTA.Text
            Set tRs2 = cnn.Execute(sBusca)
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not tRs.EOF
                    ' ---------------------------- PRODUCTOS EQUIVALENTES ----------------------------
                    ' ------------------------------ 25/05/2021 H VALDEZ -----------------------------
                    ClvProd = tRs.Fields("ID_PRODUCTO")
                    Canti = CDbl(tRs.Fields("CANTIDAD"))
                    sBuscar = "SELECT JUEGO_REPARACION.ID_PRODUCTO, JUEGO_REPARACION.CANTIDAD FROM ALMACEN3 INNER JOIN JUEGO_REPARACION ON ALMACEN3.ID_PRODUCTO = JUEGO_REPARACION.ID_REPARACION WHERE (ALMACEN3.TIPO = 'EQUIVALE') AND  ALMACEN3.ID_PRODUCTO = '" & ClvProd & "'"
                    Set tRs3 = cnn.Execute(sBuscar)
                    If Not (tRs3.EOF And tRs3.BOF) Then
                        ClvProd = tRs3.Fields("ID_PRODUCTO")
                        Canti = tRs3.Fields("CANTIDAD") * Canti
                    End If
                    ' ++++++++++++++++++++++++++++ PRODUCTOS EQUIVALENTES ++++++++++++++++++++++++++++
                    sBusca = "SELECT CANTIDAD FROM EXISTENCIAS WHERE SUCURSAL = '" & Sucursal & "' AND ID_PRODUCTO = '" & ClvProd & "'"
                    Set tRs3 = cnn.Execute(sBusca)
                    If Not (tRs3.EOF And tRs3.BOF) Then
                        sBusca = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD + " & Canti & " WHERE ID_PRODUCTO = '" & ClvProd & "' AND SUCURSAL = '" & Sucursal & "'"
                        cnn.Execute (sBusca)
                    Else
                        sBuscar = "SELECT ID_EXIS_FIJA FROM EXISTENCIA_FIJA WHERE ID_PRODUCTO = '" & ClvProd & "' AND CANTIDAD = -" & tRs.Fields("CANTIDAD")
                        Set tRs5 = cnn.Execute(sBuscar)
                        If Not (tRs5.BOF And tRs5.EOF) Then
                            sBuscar = "DELETE FROM EXISTENCIA_FIJA WHERE ID_EXIS_FIJA = " & tRs5.Fields("ID_EXIS_FIJA")
                            cnn.Execute (sBuscar)
                            sBuscar = "SELECT * FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ClvProd & "'"
                            Set tRs5 = cnn.Execute(sBuscar)
                            If Not (tRs5.BOF And tRs5.EOF) Then
                                sBusca = "INSERT INTO EXISTENCIAS (CANTIDAD, ID_PRODUCTO, SUCURSAL) VALUES (" & tRs.Fields("CANTIDAD") & ", '" & ClvProd & "', '" & Sucursal & "');"
                                cnn.Execute (sBusca)
                            Else
                                sBusca = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD + " & tRs.Fields("CANTIDAD") & " WHERE ID_PRODUCTO = '" & ClvProd & "' AND SUCURSAL = '" & Sucursal & "' "
                            cnn.Execute (sBusca)
                            End If
                        Else
                            sBusca = "INSERT INTO EXISTENCIAS (CANTIDAD, ID_PRODUCTO, SUCURSAL) VALUES (" & Canti & ", '" & ClvProd & "', '" & Sucursal & "');"
                            cnn.Execute (sBusca)
                        End If
                    End If
                    tRs.MoveNext
                Loop
            End If
        End If
    End If
    txtIDVENTA = ""
    Actual (ClvCliente)
End Sub
Private Sub Command4_Click()
    If Text2.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        If Option3.Value = True Then
            sBuscar = "SELECT ID_VENTA, FOLIO, SUCURSAL, SUBTOTAL, TOTAL, IVA, NOMBRE, UNA_EXIBICION,FECHA FROM VENTAS WHERE ID_VENTA = " & Text2.Text
        Else
            sBuscar = "SELECT ID_VENTA, FOLIO, SUCURSAL, SUBTOTAL, TOTAL, IVA, NOMBRE, UNA_EXIBICION,FECHA FROM VENTAS WHERE FOLIO = '" & Text2.Text & "'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        ListView3.ListItems.Clear
        ListView4.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_VENTA"))
                If Not IsNull(tRs.Fields("FOLIO")) Then tLi.SubItems(1) = tRs.Fields("FOLIO")
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(2) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(3) = tRs.Fields("FECHA")
                If Not IsNull(tRs.Fields("SUBTOTAL")) Then tLi.SubItems(4) = tRs.Fields("SUBTOTAL")
                If Not IsNull(tRs.Fields("IVA")) Then tLi.SubItems(5) = tRs.Fields("IVA")
                If Not IsNull(tRs.Fields("TOTAL")) Then tLi.SubItems(6) = tRs.Fields("TOTAL")
                If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(7) = tRs.Fields("SUCURSAL")
                If Not IsNull(tRs.Fields("UNA_EXIBICION")) Then
                    If tRs.Fields("UNA_EXIBICION") = "S" Then
                        tLi.SubItems(8) = "CONTADO"
                    Else
                        tLi.SubItems(8) = "CREDITO"
                    End If
                End If
                tRs.MoveNext
            Loop
        Else
            MsgBox "No se ha encontrado el registro, es posible que se cancelara o no este capturado!", vbExclamation, "SACC"
        End If
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
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
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Cliente", 500
        .ColumnHeaders.Add , , "Nombre", 7450
    End With
    With ListView2(0)
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Folio", 1000
        .ColumnHeaders.Add , , "Total Compra", 2000
        .ColumnHeaders.Add , , "ID_VENTA", 0
    End With
    With ListView2(1)
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Nota", 1000
        .ColumnHeaders.Add , , "Total Compra", 2000
        .ColumnHeaders.Add , , "ID_VENTA", 0
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "NO. VENTA", 1000
        .ColumnHeaders.Add , , "FOLIO", 1000
        .ColumnHeaders.Add , , "CLIENTE", 4500
        .ColumnHeaders.Add , , "FECHA", 1500
        .ColumnHeaders.Add , , "SUBTOTAL", 1000
        .ColumnHeaders.Add , , "IVA", 1000
        .ColumnHeaders.Add , , "TOTAL", 1000
        .ColumnHeaders.Add , , "SUCURSAL", 1500
        .ColumnHeaders.Add , , "TIPO", 1000
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "CLAVE PRODUCTO", 4500
        .ColumnHeaders.Add , , "PRECIO DE VENTA", 1500
        .ColumnHeaders.Add , , "CANTIDAD", 1500
    End With
    Text5.Text = VarMen.Text1(0).Text
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    ClvCliente = Item
    Text3.Text = Item.SubItems(1)
    Actual (Item)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Actual(Item As String)
    Dim Acum As Double
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim tLi2 As ListItem
    Command3.Enabled = False
    ListView2(0).ListItems.Clear
    ListView2(1).ListItems.Clear
    txtNo.Text = ""
    'AQUI
    sBuscar = "SELECT ID_VENTA, FOLIO, TOTAL FROM VENTAS WHERE ID_CLIENTE = " & Item '& " AND ID_CUENTA = 0 "
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If .EOF And .BOF Then
            MsgBox "EL CLIENTE NO TIENE MOVIMIENTOS DE COMPRA", vbInformation, "SACC"
        Else
            Do While Not .EOF
                If (Not IsNull(.Fields("Folio"))) And (.Fields("Folio") <> "") Then
                    Set tLi = ListView2(0).ListItems.Add(, , .Fields("Folio"))
                Else
                    Set tLi = ListView2(1).ListItems.Add(, , .Fields("Id_Venta"))
                End If
                    If Not IsNull(.Fields("Total")) Then tLi.SubItems(1) = CDbl(.Fields("Total"))
                    tLi.SubItems(2) = .Fields("Id_Venta")
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub ListView2_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    txtNo.Text = Item
    If Option1(0).Value = False Then
        txtIDVENTA.Text = Item.SubItems(2)
    Else
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        txtIDVENTA.Text = ""
        sBuscar = "SELECT ID_VENTA FROM VENTAS WHERE FOLIO = '" & txtNo.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        Do While Not tRs.EOF
            If txtIDVENTA.Text = "" Then
                txtIDVENTA.Text = tRs.Fields("ID_VENTA")
            Else
                txtIDVENTA.Text = txtIDVENTA.Text & "," & tRs.Fields("ID_VENTA")
            End If
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Text4.Text = Item
    sBuscar = "SELECT ID_PRODUCTO, CANTIDAD, PRECIO_VENTA FROM VENTAS_DETALLE WHERE ID_VENTA = '" & Text4.Text & "' "
    Set tRs = cnn.Execute(sBuscar)
    ListView4.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView4.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(1) = tRs.Fields("PRECIO_VENTA")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Option1_Click(Index As Integer)
    If Option1(1).Value Then
        ListView2(1).Visible = True
        ListView2(0).Visible = False
        lblNoMov.Caption = "Nota V. No:"
        lblTitulo.Caption = "Notas Pendientes de Pago"
    Else
        ListView2(1).Visible = False
        ListView2(0).Visible = True
        lblNoMov.Caption = "Factura No:"
        lblTitulo.Caption = "Facturas Pendientes de Pago"
    End If
End Sub
Private Sub Option2_Click(Index As Integer)
    If Option2(1).Value Then
        lblTipoMov.Caption = "Numero de Nota V.:"
        Command1.Caption = "Buscar Nota V."
    Else
        lblTipoMov.Caption = "Numero de Factura:"
        Command1.Caption = "Buscar Factura"
    End If
End Sub
Private Sub Text1_Change()
On Error GoTo ManejaError
    If Text1.Text = "" Then
        Me.Command2.Enabled = False
    Else
        Me.Command2.Enabled = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If Text1.Text <> "" Then
        If KeyAscii = 13 Then
            Buscar
            ListView1.SetFocus
        End If
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text14_Change()
    If Text14.Text = "" Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text14.Text <> "" Then
            Command1.Value = True
        End If
    End If
    Dim Valido As String
    Valido = "1234567890ABCDEFGHIJKLMNÑOPQRSTUVWXYZ"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text14_GotFocus()
    Text14.BackColor = &HFFE1E1
End Sub
Private Sub Text14_LostFocus()
    Text14.BackColor = &H80000005
End Sub
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND LIMITE_CREDITO <> 0"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            Text1.Text = ""
            MsgBox "No se encontro cliente registrado a ese nombre"
        Else
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                tLi.SubItems(1) = .Fields("NOMBRE") & ""
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    If KeyAscii = 13 Then
        Command4.Value = True
    End If
    If Option3.Value = True Then
        Valido = "1234567890"
    Else
        Valido = "1234567890ABCDEFGHIJKLMANÑOPQRSTUVWXYZ-"
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub txtNo_Change()
    Command3.Enabled = False
    If txtNo.Text <> "" Then Command3.Enabled = True
End Sub
