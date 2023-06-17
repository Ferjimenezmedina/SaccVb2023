VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmReviComa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REVISAR COMANDAS"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Traer"
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
      Left            =   9480
      Picture         =   "frmReviComa.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9480
      TabIndex        =   25
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   23
      Top             =   2760
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
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image14 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmReviComa.frx":29D2
         MousePointer    =   99  'Custom
         Picture         =   "frmReviComa.frx":2CDC
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   21
      Top             =   5280
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmReviComa.frx":48DE
         MousePointer    =   99  'Custom
         Picture         =   "frmReviComa.frx":4BE8
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
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   19
      Top             =   3960
      Width           =   975
      Begin VB.Image cmdOrden 
         Height          =   705
         Left            =   120
         MouseIcon       =   "frmReviComa.frx":6CCA
         MousePointer    =   99  'Custom
         Picture         =   "frmReviComa.frx":6FD4
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Orden"
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
   End
   Begin VB.TextBox txtId_Reparacion 
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin TabDlg.SSTab sstReviComa 
      Height          =   4575
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   8070
      _Version        =   393216
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Tinta"
      TabPicture(0)   =   "frmReviComa.frx":890E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvwTinta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAceptarTinta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdEditarTinta"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancelarTinta"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Toner"
      TabPicture(1)   =   "frmReviComa.frx":892A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(1)=   "cmdCancelarToner"
      Tab(1).Control(2)=   "cdmEditarToner"
      Tab(1).Control(3)=   "cmdAceptarComandaToner"
      Tab(1).Control(4)=   "lvwToner"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Faltantes"
      TabPicture(2)   =   "frmReviComa.frx":8946
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvwFaltas"
      Tab(2).Control(1)=   "cmdPedir"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton Command2 
         Caption         =   "Rema"
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
         Left            =   -67200
         Picture         =   "frmReviComa.frx":8962
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelarToner 
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
         Height          =   375
         Left            =   -67200
         Picture         =   "frmReviComa.frx":B334
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelarTinta 
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
         Height          =   375
         Left            =   7800
         Picture         =   "frmReviComa.frx":DD06
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cdmEditarToner 
         Caption         =   "Editar"
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
         Left            =   -67200
         Picture         =   "frmReviComa.frx":106D8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditarTinta 
         Caption         =   "Editar"
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
         Left            =   7800
         Picture         =   "frmReviComa.frx":130AA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdPedir 
         Caption         =   "Pedir"
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
         Left            =   -67200
         Picture         =   "frmReviComa.frx":15A7C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarTinta 
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
         Left            =   7800
         Picture         =   "frmReviComa.frx":1844E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarComandaToner 
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
         Left            =   -67200
         Picture         =   "frmReviComa.frx":1AE20
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3480
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwTinta 
         Height          =   3735
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   6588
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwToner 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   8
         Top             =   600
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   6588
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwFaltas 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   12
         Top             =   600
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   6588
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   0
      TabIndex        =   14
      Top             =   -120
      Width           =   11055
      Begin VB.TextBox txtId_Comanda 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2640
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdTraer 
         Caption         =   "Traer"
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
         Left            =   600
         Picture         =   "frmReviComa.frx":1D7F2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtCantidadComanda 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   0
         Top             =   720
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2778
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Image Image1 
         Height          =   1545
         Left            =   3960
         Picture         =   "frmReviComa.frx":201C4
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Comanda"
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label lblEstado 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   6600
      Width           =   9135
   End
End
Attribute VB_Name = "frmReviComa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim lvSI As ListSubItem
Dim tRs As ADODB.Recordset
Dim tRs2 As ADODB.Recordset
Dim intIndex As Integer
Dim bBandExis As Boolean
Dim fechsistem As Date
Dim VarComVista As String
Private Sub cdmEditarToner_Click()
On Error GoTo ManejaError
    If VarMen.Text1(65).Text = "S" Then
        If Me.lvwToner.SelectedItem.Selected Then
            If Tiene_Autorizacion Then
                Me.txtId_Reparacion.Text = Me.lvwToner.SelectedItem.SubItems(1)
                frmEditarJR.Show vbModal, Me
            End If
        Else
            MsgBox "SELECCIONE EL PRODUCTO", vbInformation, "SACC"
        End If
    Else
        MsgBox "NO CUENTA CON PERMISOS!", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdAceptarComandaToner_Click()
On Error GoTo ManejaError
    DesSeleccionar_Faltantes
    Dim Cont As Integer
    Dim NiRe As Integer
    Dim tRs As ADODB.Recordset
    Dim Limpia As Boolean
    Limpia = False
    NoRe = Me.lvwToner.ListItems.Count
    For Cont = 1 To NoRe
        If Me.lvwToner.ListItems.Item(Cont).Checked = True Then
            If Me.lvwToner.ListItems.Item(Cont).SubItems(4) = "INCOMPLETO" Then
                Seleccionar_Faltantes Trim(Me.lvwToner.ListItems.Item(Cont).SubItems(1))
                Me.lblEstado.Caption = "No hay suficiente existencia en bodega para producir, haga un pedido"
                Me.lblEstado.ForeColor = vbRed
                Me.sstReviComa.Tab = 2
            Else
                If Me.lvwToner.ListItems.Item(Cont).SubItems(5) = "A" Then
                    sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'R' WHERE ID_COMANDA = " & Me.txtId_Comanda.Text & " AND ARTICULO = " & Me.lvwToner.ListItems.Item(Cont)
                Else
                    sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'S' WHERE ID_COMANDA = " & Me.txtId_Comanda.Text & " AND ARTICULO = " & Me.lvwToner.ListItems.Item(Cont)
                End If
                cnn.Execute (sqlQuery)
                '''''inventario
                sqlQuery = "SELECT COUNT(TEMPORAL) TEMPORAL FROM JR_TEMPORALES WHERE ID_COMANDA = " & Me.txtId_Comanda.Text & " AND ID_REPARACION = '" & lvwToner.ListItems.Item(Cont).SubItems(1) & "'"
                Set tRs = cnn.Execute(sqlQuery)
                If tRs.Fields("TEMPORAL") = 0 Then
                    sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & Me.lvwToner.ListItems.Item(Cont).SubItems(1) & "'"
                Else
                    sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JR_TEMPORALES WHERE ID_REPARACION = '" & Me.lvwToner.ListItems.Item(Cont).SubItems(1) & "' AND ID_COMANDA = " & Me.txtId_Comanda.Text
                End If
                Set tRs = cnn.Execute(sqlQuery)
                If Not (tRs.EOF And tRs.BOF) Then
                    Do While Not tRs.EOF
                        sqlQuery = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD - " & (CDbl(Me.lvwToner.ListItems.Item(Cont).SubItems(3)) * CDbl(tRs.Fields("CANTIDAD"))) & " WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND SUCURSAL = 'BODEGA'"
                        cnn.Execute (sqlQuery)
                        tRs.MoveNext
                    Loop
                End If
                Cerrar_Comanda
            End If
        End If
    Next Cont
    lvwToner.ListItems.Clear
    RemasPend
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdAceptarTinta_Click()
On Error GoTo ManejaError
    DesSeleccionar_Faltantes
    Dim Cont As Integer
    Dim NiRe As Integer
    Dim sqlQuery As String
    Dim Limpia As Boolean
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Limpia = False
    NoRe = Me.lvwTinta.ListItems.Count
    For Cont = 1 To NoRe
        If Me.lvwTinta.ListItems.Item(Cont).Checked = True Then
            If Me.lvwTinta.ListItems.Item(Cont).SubItems(4) = "INCOMPLETO" Then
                Seleccionar_Faltantes Trim(Me.lvwTinta.ListItems.Item(Cont).SubItems(1))
                Me.lblEstado.Caption = "No hay suficiente existencia en bodega para producir, haga un pedido"
                Me.lblEstado.ForeColor = vbRed
                Me.sstReviComa.Tab = 2
            Else
                If Me.lvwTinta.ListItems.Item(Cont).SubItems(5) = "A" Then
                    sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'R' WHERE ID_COMANDA = " & Me.txtId_Comanda.Text & " AND ARTICULO = " & Me.lvwTinta.ListItems.Item(Cont)
                Else
                    sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'S' WHERE ID_COMANDA = " & Me.txtId_Comanda.Text & " AND ARTICULO = " & Me.lvwTinta.ListItems.Item(Cont)
                End If
                cnn.Execute (sqlQuery)
                ' INICIA EL DESCUENTO DEL INVENTARIO
                sqlQuery = "SELECT COUNT(TEMPORAL) TEMPORAL FROM JR_TEMPORALES WHERE ID_COMANDA = " & Me.txtId_Comanda.Text & " AND ID_REPARACION = '" & Me.lvwTinta.ListItems.Item(Cont).SubItems(1) & "'"
                Set tRs = cnn.Execute(sqlQuery)
                If tRs.Fields("TEMPORAL") = 0 Then
                    sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & Me.lvwTinta.ListItems.Item(Cont).SubItems(1) & "'"
                Else
                    sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JR_TEMPORALES WHERE ID_REPARACION = '" & Me.lvwTinta.ListItems.Item(Cont).SubItems(1) & "' AND ID_COMANDA = " & Me.txtId_Comanda.Text
                End If
                Set tRs = cnn.Execute(sqlQuery)
                If Not (tRs.EOF And tRs.BOF) Then
                    Do While Not tRs.EOF
                        sqlQuery = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD - " & (CDbl(Me.lvwTinta.ListItems.Item(Cont).SubItems(3)) * CDbl(tRs.Fields("CANTIDAD"))) & " WHERE ID_PRODUCTO = '" & tRs.Fields("Id_Producto") & "' AND SUCURSAL = 'BODEGA'"
                        'VERIFICACION DEL DESCUENTO DE INVENTARIO
                        Dim iAfectados As Long
                        Set tRs1 = cnn.Execute(sqlQuery, iAfectados, adCmdText)
                        If iAfectados < 1 Then
                            sqlQuery = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD - " & (CDbl(Me.lvwTinta.ListItems.Item(Cont).SubItems(3)) * CDbl(tRs.Fields("CANTIDAD"))) & " WHERE ID_PRODUCTO = '" & tRs.Fields("Id_Producto") & "' AND SUCURSAL = 'BODEGA'"
                            Set tRs1 = cnn.Execute(sqlQuery, iAfectados, adCmdText)
                            If iAfectados < 1 Then
                                MsgBox "No se realizo el descuento del inventario del producto " & tRs.Fields("Id_Producto"), vbExclamation, "SACC"
                            End If
                        End If
                        'FIN DE LA VERIFICACION
                        tRs.MoveNext
                    Loop
                Else
                    MsgBox "No se realizo el descuento del inventario apropiadamente", vbExclamation, "SACC"
                End If
                ' TERMINA EL DESCUENTO DEL INVENTARIO
                Cerrar_Comanda
            End If
        End If
    Next Cont
    lvwTinta.ListItems.Clear
    RemasPend
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdCancelarTinta_Click()
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim NoRe As Integer
    Dim bBanderaTinta As Boolean
    Dim cNota As String
    bBanderaTinta = False
    NoRe = Me.lvwTinta.ListItems.Count
    For Cont = 1 To NoRe
        If Me.lvwTinta.ListItems.Item(Cont).Checked = True Then
            bBanderaTinta = True
            cNota = cNota + " " + Me.lvwTinta.ListItems.Item(Cont).SubItems(1)
        End If
    Next Cont
    If bBanderaTinta = True Then
        If MsgBox("SE CANCELARAN SOLO LOS ARTICULOS SELECCIONADOS, ¿DESEA CONTINUAR?", vbInformation + vbYesNo + vbDefaultButton1, "MESAJE DEL SISTEMA") = vbYes Then
            cNota = InputBox("INTRODUSCA EL MOTIVO DE LA CANCELACIÓN", "CANCELANDO COMANDA " & Me.txtId_Comanda.Text & " DE TINTA", "Los siguientes productos de Tinta no llegaron:" & cNota & ".")
            ''
            For Cont = 1 To NoRe
                If Me.lvwTinta.ListItems.Item(Cont).Checked = True Then
                    sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = '0' WHERE ID_COMANDA = " & Me.txtId_Comanda.Text & " AND ARTICULO = " & Me.lvwTinta.ListItems.Item(Cont)
                    Set tRs = cnn.Execute(sqlQuery)
                End If
            Next Cont
            sqlQuery = "UPDATE COMANDAS_2 SET NOTAS = '" & cNota & "' WHERE ID_COMANDA = " & Me.txtId_Comanda.Text
            Set tRs = cnn.Execute(sqlQuery)
            Me.lvwTinta.ListItems.Clear
            Me.lblEstado.Caption = "Comanda " & Me.txtId_Comanda.Text & " cancelada"
            Me.lblEstado.ForeColor = vbBlack
            Me.cmdTraer.Value = True
        End If
    Else
        Me.lblEstado.Caption = "Seleccione un articulo de la lista"
        Me.lblEstado.ForeColor = vbRed
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdCancelarToner_Click()
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim NoRe As Integer
    Dim bBanderaToner As Boolean
    Dim cNota As String
    bBanderaToner = False
    NoRe = Me.lvwToner.ListItems.Count
    For Cont = 1 To NoRe
        If Me.lvwToner.ListItems.Item(Cont).Checked = True Then
            bBanderaToner = True
            cNota = cNota + " " + Me.lvwToner.ListItems.Item(Cont).SubItems(1)
        End If
    Next Cont
    If bBanderaToner = True Then
        If MsgBox("SE CANCELARAN SOLO LOS ARTICULOS SELECCIONADOS, ¿DESEA CONTINUAR?", vbInformation + vbYesNo + vbDefaultButton1, "MESAJE DEL SISTEMA") = vbYes Then
            cNota = InputBox("INTRODUSCA EL MOTIVO DE LA CANCELACIÓN", "CANCELANDO COMANDA " & Me.txtId_Comanda.Text & " DE TINTA", "Los siguientes productos de Tinta no llegaron:" & cNota & ".")
            ''
            For Cont = 1 To NoRe
                If Me.lvwToner.ListItems.Item(Cont).Checked = True Then
                    sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'C'WHERE ID_COMANDA = " & Me.txtId_Comanda.Text & " AND ARTICULO = " & Me.lvwToner.ListItems.Item(Cont)
                    Set tRs = cnn.Execute(sqlQuery)
                End If
            Next Cont
            sqlQuery = "UPDATE COMANDAS_2 SET NOTAS = '" & cNota & "' WHERE ID_COMANDA = " & Me.txtId_Comanda.Text
            Set tRs = cnn.Execute(sqlQuery)
            Me.lvwToner.ListItems.Clear
            Me.lblEstado.Caption = "Comanda " & Me.txtId_Comanda.Text & " cancelada"
            Me.lblEstado.ForeColor = vbBlack
            Me.cmdTraer.Value = True
        End If
    Else
        Me.lblEstado.Caption = "Seleccione un articulo de la lista"
        Me.lblEstado.ForeColor = vbRed
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdEditarTinta_Click()
On Error GoTo ManejaError
    If VarMen.Text1(65).Text = "S" Then
            If Me.lvwTinta.SelectedItem.Selected Then
                If Tiene_Autorizacion Then
                    Me.txtId_Reparacion.Text = Me.lvwTinta.SelectedItem.SubItems(1)
                    frmEditarJR.Show vbModal, Me
                End If
            Else
                MsgBox "SELECCIONE EL PRODUCTO", vbInformation, "SACC"
            End If
    Else
        MsgBox "NO CUENTA CON PERMISOS!", vbInformation, "SACC"
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdOrden_Click()
    frmVerOrdenesProduccion.Show vbModal
    RemasPend
End Sub
Private Sub cmdPedir_Click()
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim NoRe As Integer
    Dim nPedido As Integer
    Dim cProducto As String
    Dim Uno As String
    Dim cMin As Double
    Dim cMax As Double
    Dim CantPed As Integer
    Dim CantPedF As Integer
    Dim exist As Integer
    Dim Almacen As String
    Dim Marca As String
    cMin = 0
    cMax = 0
    CantPed = 0
    If Tiene_Autorizacion Then
        If MsgBox("¿DESEA HACER UN PEDIDO POR LOS ARTICULOS SELECCIONADOS?", vbQuestion + vbYesNo + vbDefaultButton1, "SACC") = vbYes Then
            Uno = "S"
            sqlQuery = "INSERT INTO PEDIDO (SUCURSAL, PIDIO, FECHA, TIPO, COMENTARIO) VALUES ('BODEGA', '1', '" & Format(Date, "dd/mm/yyyy") & "', 'I', 'MAXIMOS Y MINIMOS')"
            cnn.Execute (sqlQuery)
            NoRe = Me.lvwFaltas.ListItems.Count
            For Cont = 1 To NoRe
                If Me.lvwFaltas.ListItems.Item(Cont).Checked = True Then
                    Me.lblEstado.Caption = "Pidiendo " & Me.lvwFaltas.ListItems.Item(Cont).SubItems(2) & " de " & Me.lvwFaltas.ListItems.Item(Cont).SubItems(1)
                    Me.lblEstado.ForeColor = vbBlack
                    DoEvents
                    sqlQuery = "SELECT A.Descripcion, A.C_MINIMA, A.C_MAXIMA, ISNULL(E.CANTIDAD,0), A.MARCA AS EXISTENCIA FROM ALMACEN2 AS A LEFT JOIN EXISTENCIAS AS E ON A.ID_PRODUCTO = E.ID_PRODUCTO WHERE A.ID_PRODUCTO = '" & Me.lvwFaltas.ListItems.Item(Cont).SubItems(1) & "'"
                    Set tRs = cnn.Execute(sqlQuery)
                    If Not tRs.EOF And Not tRs.BOF Then
                        cProducto = tRs.Fields("Descripcion")
                        If Not IsNull(tRs.Fields("C_MINIMA")) Then cMin = tRs.Fields("C_MINIMA")
                        If Not IsNull(tRs.Fields("C_MAXIMA")) Then cMax = tRs.Fields("C_MAXIMA")
                        exist = tRs.Fields("EXISTENCIA")
                        Marca = tRs.Fields("MARCA")
                        Almacen = "A2"
                    Else
                        sqlQuery = "SELECT A.Descripcion, A.C_MINIMA, A.C_MAXIMA, ISNULL(E.CANTIDAD,0), A.MARCA AS EXISTENCIA FROM ALMACEN1 AS A LEFT JOIN EXISTENCIAS AS E ON A.ID_PRODUCTO = E.ID_PRODUCTO WHERE ID_PRODUCTO = '" & Trim(Me.lvwFaltas.ListItems.Item(Cont).SubItems(1)) & "'"
                        Set tRs = cnn.Execute(sqlQuery)
                        If Not tRs.EOF And Not tRs.BOF Then
                            cProducto = tRs.Fields("Descripcion")
                            If Not IsNull(tRs.Fields("C_MINIMA")) Then cMin = tRs.Fields("C_MINIMA")
                            If Not IsNull(tRs.Fields("C_MAXIMA")) Then cMax = tRs.Fields("C_MAXIMA")
                            exist = tRs.Fields("EXISTENCIA")
                            Marca = tRs.Fields("MARCA")
                            Almacen = "A1"
                        Else
                            MsgBox "NO SE ENCONTRO EL REGISTRO EN ALMACENES DE " & Me.lvwFaltas.ListItems.Item(Cont).SubItems(1), vbCritical, "MENSAJE DEL SISITEMA"
                            cProducto = ""
                        End If
                    End If
                    If cProducto <> "" Then
                        If cMax > cMin Then
                            sqlQuery = "SELECT D.CANTIDAD FROM DETALLE_PEDIDO AS D JOIN PEDIDO AS P WHERE D.ID_PRODUCTO = '" & lvwFaltas.ListItems.Item(Cont) & "' AND (D.ENTREGADO = '0' OR D.ENTREGADO = 'R') AND P.COMENTARIO = 'MAXIMOS Y MINIMOS' AND P.PIDIO = '1'"
                            Set tRs = cnn.Execute(sqlQuery)
                            If Not (tRs.EOF And tRs.BOF) Then
                                Do While Not tRs.EOF
                                    CantPed = CantPed + tRs.Fields("CANTIDAD")
                                Loop
                            End If
                        End If
                        CantPedF = cMax - CantPed - exist
                        If (CantPedF > 0) Or (cMax = 0) Then
                            If Uno = "S" Then
                                sqlQuery = "INSERT INTO PEDIDO (SUCURSAL, PIDIO, FECHA, TIPO, COMENTARIO) VALUES ('BODEGA', '1', '" & Format(Date, "dd/mm/yyyy") & "', 'I', 'MAXIMOS Y MINIMOS')"
                                cnn.Execute (sqlQuery)
                                sqlQuery = "SELECT TOP 1 ID_PEDIDO FROM PEDIDO ORDER BY ID_PEDIDO DESC"
                                Set tRs = cnn.Execute(sqlQuery)
                                nPedido = tRs.Fields("ID_PEDIDO")
                                Uno = "N"
                            End If
                            If (cMax = 0) Then
                                CantPedF = CantPed
                            End If
                            sqlQuery = "INSERT INTO DETALLE_PEDIDO (ID_PRODUCTO, CANTIDAD, ID_PEDIDO, DESCRIPCION, ALMACEN, MARCA) VALUES ('" & Trim(Me.lvwFaltas.ListItems.Item(Cont).SubItems(1)) & "', " & (cMax - CantPed - exist) & ", " & nPedido & ", '" & cProducto & "', '" & Almacen & "', '" & Marca & "')"
                            cnn.Execute (sqlQuery)
                        End If
                    End If
                    
                End If
            Next Cont
            Me.lblEstado.Caption = "Pedido listo"
            DoEvents
        End If
    Else
        'NO TIENE AUTORIZACION
        MsgBox "AUTORIZACION DENEGADA", vbExclamation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdTraer_Click_Click()
On Error GoTo ManejaError
    lvwTinta.ListItems.Clear
    lvwToner.ListItems.Clear
    lvwFaltas.ListItems.Clear
    If Puede_Traer_Comanda(Val(Me.txtCantidadComanda.Text)) Then
        If Hay_Comanda(Me.txtCantidadComanda.Text) Then
            'Me.lblEstado.Caption = ""
            Me.lvwFaltas.ListItems.Clear
            Llenar_Lista_Comandas (Me.txtCantidadComanda.Text)
            Revisar_Juegos_Reparacion
            'bBandera = False
            Me.txtId_Comanda.Text = Me.txtCantidadComanda.Text
            If Me.lvwTinta.ListItems.Count <> 0 Then
                Me.sstReviComa.Tab = 0
            Else
                Me.sstReviComa.Tab = 1
            End If
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Puede_Traer_Comanda(nComanda As Integer) As Boolean
On Error GoTo ManejaError
    If Trim(Me.txtCantidadComanda.Text) = "" Then
        Puede_Traer_Comanda = False
        Me.lblEstado.Caption = "Introsusca el numero de comanda"
        Me.lblEstado.ForeColor = vbRed
        Me.txtCantidadComanda.SetFocus
        Exit Function
    End If
    Puede_Traer_Comanda = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Llenar_Lista_Comandas(nID_Comanda As Integer)
On Error GoTo ManejaError
    sqlQuery = "SELECT C.ID_COMANDA, C.FECHA_INICIO, C.SUCURSAL, CD.ARTICULO, CD.ID_PRODUCTO, CD.CANTIDAD, CD.TIPO, CD.ESTADO_ACTUAL, A.Descripcion FROM COMANDAS_2 AS C JOIN COMANDAS_DETALLES_2 AS CD ON C.ID_COMANDA = CD.ID_COMANDA JOIN ALMACEN3 AS A ON A.ID_PRODUCTO = CD.ID_PRODUCTO WHERE C.ID_COMANDA = " & nID_Comanda & " AND C.ESTADO_ACTUAL = 'A' AND (CD.ESTADO_ACTUAL = 'A' OR CD.ESTADO_ACTUAL = 'B') ORDER BY C.ID_COMANDA"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        Me.lvwTinta.ListItems.Clear
        Me.lvwToner.ListItems.Clear
        Do While Not .EOF
            If Not IsNull(.Fields("TIPO")) Then
                If .Fields("TIPO") = "I" Then
                    Set tLi = Me.lvwTinta.ListItems.Add(, , .Fields("ARTICULO"))
                    If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = .Fields("ID_PRODUCTO")
                    If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(2) = .Fields("Descripcion")
                    If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(3) = .Fields("CANTIDAD")
                    If Not IsNull(.Fields("ESTADO_ACTUAL")) Then tLi.SubItems(5) = .Fields("ESTADO_ACTUAL")
                    If Not IsNull(.Fields("SUCURSAL")) Then tLi.SubItems(6) = .Fields("SUCURSAL")
                ElseIf .Fields("TIPO") = "T" Then
                    Set tLi = Me.lvwToner.ListItems.Add(, , .Fields("ARTICULO"))
                    If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = .Fields("ID_PRODUCTO")
                    If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(2) = .Fields("Descripcion")
                    If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(3) = .Fields("CANTIDAD")
                    If Not IsNull(.Fields("ESTADO_ACTUAL")) Then tLi.SubItems(5) = .Fields("ESTADO_ACTUAL")
                    If Not IsNull(.Fields("SUCURSAL")) Then tLi.SubItems(6) = .Fields("SUCURSAL")
                End If
            End If
            .MoveNext
        Loop
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub cmdTraer_Click()
On Error GoTo ManejaError
    lvwTinta.ListItems.Clear
    lvwToner.ListItems.Clear
    lvwFaltas.ListItems.Clear
    If Puede_Traer_Comanda(Val(Me.txtCantidadComanda.Text)) Then
        If Hay_Comanda(Me.txtCantidadComanda.Text) Then
            Me.lvwFaltas.ListItems.Clear
            Llenar_Lista_Comandas (Me.txtCantidadComanda.Text)
            Revisar_Juegos_Reparacion
            Me.txtId_Comanda.Text = Me.txtCantidadComanda.Text
            If Me.lvwTinta.ListItems.Count <> 0 Then
                Me.sstReviComa.Tab = 0
            Else
                Me.sstReviComa.Tab = 1
            End If
            VarComVista = txtCantidadComanda.Text
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command1_Click()
    FrmComiciones.Show vbModal
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim Con As Integer
    If lvwToner.ListItems.Count > 0 Then
        For Con = 1 To lvwToner.ListItems.Count
            If lvwToner.ListItems(Con).Checked Then
                FrmCalidad3.Caption = "CALIDAD DE " & lvwToner.ListItems(Con).SubItems(1)
                FrmCalidad3.lblArticulo.Caption = lvwToner.ListItems(Con)
                FrmCalidad3.lblComanda.Caption = VarComVista
                FrmCalidad3.lblCantidad.Caption = lvwToner.ListItems(Con).SubItems(3)
                FrmCalidad3.txtNumArticulo.Text = lvwToner.ListItems(Con)
                FrmCalidad3.txtEdo.Text = "A"
                FrmCalidad3.txtNoSirve.Text = "0"
                FrmCalidad3.Show
            End If
        Next
        cmdTraer.Value = True
    End If
ManejaError:
    MsgBox Err.Number & ": " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Text1.Text = Str(Date)
    fechsistem = Text1.Text
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "COMANDA", 1440
    End With
    With lvwTinta
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ARTICULO", 500
        .ColumnHeaders.Add , , "CLAVE", 2000
        .ColumnHeaders.Add , , "Descripcion", 4100
        .ColumnHeaders.Add , , "CANTIDAD", 2000
        .ColumnHeaders.Add , , "EXISTENCIA", 0
        .ColumnHeaders.Add , , "ESTADO", 0
        .ColumnHeaders.Add , , "SUCURSAL", 1000
    End With
    With lvwToner
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ARTICULO", 500
        .ColumnHeaders.Add , , "CLAVE", 2000
        .ColumnHeaders.Add , , "Descripcion", 4100
        .ColumnHeaders.Add , , "CANTIDAD", 2000
        .ColumnHeaders.Add , , "EXISTENCIA", 0
        .ColumnHeaders.Add , , "ESTADO", 0
        .ColumnHeaders.Add , , "SUCURSAL", 1000
    End With
    With lvwFaltas
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "PRODUCTO", 500
        .ColumnHeaders.Add , , "REPARACION", 2000
        .ColumnHeaders.Add , , "CANTIDAD", 4100
    End With
    RemasPend
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image14_Click()
    FrmReviAutRema.Show vbModal
End Sub
Private Sub Image2_Click()
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim Cont As Integer
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
    Printer.Print VarMen.Text5(0).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
    Printer.Print "FECHA : " & tRs.Fields("FECHA")
    Printer.Print "SUCURSAL : " & VarMen.Text4(0).Text
    Printer.Print "TELEFONO SUCURSAL : " & VarMen.Text4(5).Text
    Printer.Print "No. DE VALE DE CAJA : " & tRs.Fields("ID_VALE")
    Printer.Print "No. DE VENTA : " & tRs.Fields("ID_VENTA")
    Printer.Print "ATENDIDO POR : " & VarMen.Text1(1).Text
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtCantidadComanda.Text = Item
End Sub
Private Sub lvwTinta_DblClick()
On Error GoTo ManejaError
   Me.cmdEditarTinta.Value = True
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwToner_DblClick()
On Error GoTo ManejaError
    Me.cdmEditarToner.Value = True
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtCantidadComanda_GotFocus()
On Error GoTo ManejaError
    Me.txtCantidadComanda.BackColor = &HFFE1E1
    Me.txtCantidadComanda.SelStart = 0
    Me.txtCantidadComanda.SelLength = Len(Me.txtCantidadComanda.Text)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtCantidadComanda_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.lvwTinta.SetFocus
        Me.cmdTraer.Value = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Revisar_Juegos_Reparacion()
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim NoRe As Integer
    NoRe = Me.lvwTinta.ListItems.Count
    For Cont = 1 To NoRe
        If Hay_existencias(Trim(Me.lvwTinta.ListItems.Item(Cont).SubItems(1)), Me.lvwTinta.ListItems.Item(Cont).SubItems(3)) = False Then
            Colorear_Item Me.lvwTinta.ListItems.Item(Cont).Index
            Me.lvwTinta.ListItems.Item(Cont).SubItems(4) = "INCOMPLETO"
        End If
    Next
    NoRe = Me.lvwToner.ListItems.Count
    For Cont = 1 To NoRe
        If Hay_existencias(Trim(Me.lvwToner.ListItems.Item(Cont).SubItems(1)), Me.lvwToner.ListItems.Item(Cont).SubItems(3)) = False Then
            Colorear_Item_Toner Me.lvwToner.ListItems.Item(Cont).Index
            Me.lvwToner.ListItems.Item(Cont).SubItems(4) = "INCOMPLETO"
        End If
    Next
    If Me.lvwFaltas.ListItems.Count <> 0 Then
        Me.lblEstado.Caption = "Faltan existencias"
        Me.lblEstado.ForeColor = vbRed
        DoEvents
    Else
        Me.lblEstado.Caption = ""
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Hay_Juego_Reparacion(cId_Producto As String) As Boolean
On Error GoTo ManejaError
    Me.lblEstado.Caption = "Buscando juego de reparacion: " & cId_Producto
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    sqlQuery = "SELECT COUNT(ID_REPARACION)ID_REPARACION FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & cId_Producto & "'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_REPARACION") <> 0 Then
            Hay_Juego_Reparacion = True
        Else
            Hay_Juego_Reparacion = False
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Colorear_Item(RowNbr As Integer)
On Error GoTo ManejaError
    Dim ItMx As ListItem
    Dim lvSI As ListSubItem
    Dim intIndex As Integer
    Set ItMx = Me.lvwTinta.ListItems(RowNbr)
    ItMx.ForeColor = vbRed
    For intIndex = 1 To Me.lvwTinta.ColumnHeaders.Count - 2
        Set lvSI = ItMx.ListSubItems(intIndex)
        lvSI.ForeColor = vbRed
    Next
    DoEvents
    Set ItMx = Nothing
    Set lvSI = Nothing
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Hay_Juego_Reparacion_Toner(cId_Producto As String) As Boolean
On Error GoTo ManejaError
    Me.lblEstado.Caption = "Buscando juego de reparacion: " & cId_Producto
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    sqlQuery = "SELECT COUNT(ID_REPARACION)ID_REPARACION FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & cId_Producto & "'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_REPARACION") <> 0 Then
            Hay_Juego_Reparacion_Toner = True
        Else
            Hay_Juego_Reparacion_Toner = False
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Colorear_Item_Toner(RowNbr As Integer)
On Error GoTo ManejaError
    Dim ItMx As ListItem
    Dim lvSI As ListSubItem
    Dim intIndex As Integer
    Set ItMx = Me.lvwToner.ListItems(RowNbr)
    ItMx.ForeColor = vbRed
    For intIndex = 1 To Me.lvwToner.ColumnHeaders.Count - 2
        Set lvSI = ItMx.ListSubItems(intIndex)
        lvSI.ForeColor = vbRed
    Next
    DoEvents
    Set ItMx = Nothing
    Set lvSI = Nothing
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Hay_existencias(cId_Producto As String, nCantidad As Double) As Boolean
On Error GoTo ManejaError
    Dim tRs3 As ADODB.Recordset
    Dim cMin As Double
    Dim Pedido As Boolean
    bBandExis = False
    Me.lblEstado.Caption = "Buscando existencias: " & cId_Producto
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    sqlQuery = "SELECT COUNT(TEMPORAL) TEMPORAL FROM JR_TEMPORALES WHERE ID_COMANDA = " & Me.txtCantidadComanda.Text & " AND ID_REPARACION = '" & cId_Producto & "'"
    Set tRs = cnn.Execute(sqlQuery)
    If tRs.Fields("TEMPORAL") = 0 Then
        sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & cId_Producto & "'"
        Set tRs = cnn.Execute(sqlQuery)
    Else
        sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JR_TEMPORALES WHERE ID_REPARACION = '" & cId_Producto & "' AND ID_COMANDA = " & Me.txtCantidadComanda.Text
        Set tRs = cnn.Execute(sqlQuery)
    End If
    With tRs
        Do While Not .EOF
                Pedido = True
                sqlQuery = "SELECT ID_EXISTENCIA, ID_PRODUCTO, CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Trim(.Fields("ID_PRODUCTO")) & "' AND SUCURSAL = 'BODEGA'"
                Set tRs2 = cnn.Execute(sqlQuery)
                If Not (tRs2.EOF And tRs2.BOF) Then
                    If Not (tRs2.Fields("CANTIDAD") >= .Fields("CANTIDAD")) Then
                        'NO HAY SUFICIENTE EXISTENCIA
                        Set tLi = Me.lvwFaltas.ListItems.Add(, , cId_Producto)
                            tLi.SubItems(1) = .Fields("ID_PRODUCTO")
                            tLi.SubItems(2) = nCantidad - tRs2.Fields("CANTIDAD")
                        bBandExis = True
                    End If
                Else
                    'NO HAY REGISTRO EN LA TABLA
                    If tRs.Fields("CANTIDAD") <> 0 Then
                        Set tLi = Me.lvwFaltas.ListItems.Add(, , cId_Producto)
                        tLi.SubItems(1) = .Fields("ID_PRODUCTO")
                        tLi.SubItems(2) = .Fields("CANTIDAD") * nCantidad
                        bBandExis = True
                    End If
                End If
            .MoveNext
        Loop
    End With
    If bBandExis = False Then
        Hay_existencias = True
    Else
        Hay_existencias = False
    End If
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Puede_Aceptar_Tinta() As Boolean
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim NoRe As Integer
    Dim bBanderaTinta As Boolean
    bBanderaTinta = False
    For Cont = 1 To NoRe
        If Me.lvwTinta.ListItems.Item(Cont).SubItems(4) = "INCOMPLETO" Then
            bBanderaTinta = True
            Seleccionar_Faltantes Trim(Me.lvwTinta.ListItems.Item(Cont).SubItems(1))
        End If
    Next Cont
    If bBanderaTinta = True Then
        Puede_Aceptar_Tinta = False
        Me.lblEstado.Caption = "No hay suficiente existencia en bodega para producir, haga un pedido"
        Me.lblEstado.ForeColor = vbRed
        Me.sstReviComa.Tab = 2
        Exit Function
    Else
        Puede_Aceptar_Tinta = True
    End If
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Seleccionar_Faltantes(cId_Producto As String)
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim NoRe As Integer
    NoRe = Me.lvwFaltas.ListItems.Count
    For Cont = 1 To NoRe
        If Me.lvwFaltas.ListItems.Item(Cont) = cId_Producto Then
            Me.lvwFaltas.ListItems.Item(Cont).Checked = True
        End If
    Next Cont
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function DesSeleccionar_Faltantes()
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim NoRe As Integer
    NoRe = Me.lvwFaltas.ListItems.Count
    For Cont = 1 To NoRe
        Me.lvwFaltas.ListItems.Item(Cont).Checked = False
    Next Cont
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Puede_Aceptar_Toner() As Boolean
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim NoRe As Integer
    Dim bBanderaToner As Boolean
    Dim cNota As String
    bBanderaToner = False
    NoRe = Me.lvwToner.ListItems.Count
    For Cont = 1 To NoRe
        If Me.lvwToner.ListItems.Item(Cont).Checked = False Then
            bBanderaToner = True
            cNota = cNota + " " + Me.lvwToner.ListItems.Item(Cont).SubItems(1)
        End If
    Next Cont
    If bBanderaToner = True Then
        Puede_Aceptar_Toner = False
        If MsgBox("LA COMANDA LLEGO INCOMPLETA Y SERA CANCELADA, ¿DESEA CONTINUAR?", vbInformation + vbYesNo + vbDefaultButton1, "MESAJE DEL SISTEMA") = vbYes Then
            cNota = InputBox("INTRODUSCA EL MOTIVO DE LA CANCELACIÓN", "CANCELANDO COMANDA " & Me.txtId_Comanda.Text & " DE TONER", "Los siguientes productos de Toner no llegaron:" & cNota & ".")
            For Cont = 1 To NoRe
                sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'C'WHERE ID_COMANDA = " & Me.txtId_Comanda.Text & " AND ARTICULO = " & Me.lvwToner.ListItems.Item(Cont)
                Set tRs = cnn.Execute(sqlQuery)
            Next Cont
            sqlQuery = "UPDATE COMANDAS_2 SET NOTAS = '" & cNota & "' WHERE ID_COMANDA = " & Me.txtId_Comanda.Text
            Set tRs = cnn.Execute(sqlQuery)
            Me.lvwToner.ListItems.Clear
            Me.lblEstado.Caption = ""
            Me.txtCantidadComanda.Text = ""
            ''
        End If
        Exit Function
    Else
        Puede_Aceptar_Toner = True
    End If
    For Cont = 1 To NoRe
        If Me.lvwToner.ListItems.Item(Cont).SubItems(4) = "INCOMPLETO" Then
            bBanderaToner = True
            Seleccionar_Faltantes Trim(Me.lvwToner.ListItems.Item(Cont).SubItems(1))
        End If
    Next Cont
    If bBanderaToner = True Then
        Puede_Aceptar_Toner = False
        Me.lblEstado.Caption = "No hay suficiente existencia en bodega para producir, haga un pedido"
        Me.lblEstado.ForeColor = vbRed
        Me.sstReviComa.Tab = 2
        Exit Function
    Else
        Puede_Aceptar_Toner = True
    End If
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Hay_Comanda(nComanda As Integer) As Boolean
On Error GoTo ManejaError
    Me.lblEstado.Caption = "Buscando comanda: " & nComanda
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    sqlQuery = "SELECT COUNT(C.ID_COMANDA)ID_COMANDA FROM COMANDAS_2 AS C JOIN COMANDAS_DETALLES_2 AS CD ON C.ID_COMANDA = CD.ID_COMANDA WHERE C.ESTADO_ACTUAL = 'A' AND CD.ESTADO_ACTUAL = 'A' AND C.ID_COMANDA = " & nComanda
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_COMANDA") <> 0 Then
            Hay_Comanda = True
            Me.lblEstado.Caption = "Se encontraron " & .Fields("ID_COMANDA") & " productos"
            Me.lblEstado.ForeColor = vbBlue
            DoEvents
        Else
            Hay_Comanda = False
            Me.lblEstado.Caption = "No se encontraron productos"
            Me.lblEstado.ForeColor = vbRed
            Me.txtCantidadComanda.SetFocus
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Descontar_Inventarios(cProducto As String, cCantidad As Double)
On Error GoTo ManejaError
    Dim POSY As String
    Me.lblEstado.Caption = "Descontando de inventarios: " & cProducto
    Me.lblEstado.ForeColor = vbBlack
    DoEvents
    '''''''''''''''''''aquita
    sqlQuery = "SELECT COUNT(TEMPORAL) TEMPORAL FROM JR_TEMPORALES WHERE ID_COMANDA = " & Me.txtId_Comanda.Text & " AND ID_REPARACION = '" & cProducto & "'"
    Set tRs = cnn.Execute(sqlQuery)
    If tRs.Fields("TEMPORAL") = 0 Then
        sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & cProducto & "'"
    Else
        sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JR_TEMPORALES WHERE ID_REPARACION = '" & cProducto & "' AND ID_COMANDA = " & Me.txtId_Comanda.Text
    End If
    Set tRs = cnn.Execute(sqlQuery)
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub txtCantidadComanda_LostFocus()
    txtCantidadComanda.BackColor = &H80000005
End Sub
Private Sub Cerrar_Comanda()
    If Me.lvwTinta.ListItems.Count = 0 And Me.lvwToner.ListItems.Count = 0 Then
        sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'R' WHERE ID_COMANDA = '" & Me.txtId_Comanda.Text & "'"
        cnn.Execute (sqlQuery)
    End If
End Sub
Function Tiene_Autorizacion() As Boolean
    Dim nClave As String
    ' Se le pasa el Hwnd del formulario y el caracter a usar _
     como contraseña
    Call inputbox_Password(Me, "*")
    'Abre el InputBox
    nClave = InputBox("INTRODUSCA CLAVE DE SUPERVISOR", "REQUIERE AUTORIZACION")
    sqlQuery = "SELECT PUESTO FROM USUARIOS WHERE ID_USUARIO = " & nClave
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.EOF And .BOF) Then
            Tiene_Autorizacion = True
        Else
            Tiene_Autorizacion = False
        End If
    End With
End Function
Private Sub RemasPend()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT ID_COMANDA FROM COMANDAS_DETALLES_2 WHERE ID_PRODUCTO LIKE '%AREM' AND ESTADO_ACTUAL = 'A' GROUP BY ID_COMANDA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_COMANDA"))
            tRs.MoveNext
        Loop
    End If
End Sub
