VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAutRema 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AUTORIZAR REMANUFACTURACIONES"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9360
      TabIndex        =   23
      Top             =   5400
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
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmAutRema.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmAutRema.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmAutRema.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblNombre"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ListaPendientes"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ListaSelec"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text2(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text2(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdAceptar"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtNoCom"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdNoComOk"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdActualizar"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdDenegar"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      Begin VB.CommandButton cmdDenegar 
         Caption         =   "Denegar"
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
         Left            =   7680
         Picture         =   "frmAutRema.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar"
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
         Picture         =   "frmAutRema.frx":4DDA
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdNoComOk 
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
         Left            =   7800
         Picture         =   "frmAutRema.frx":77AC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtNoCom 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4560
         MaxLength       =   15
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Autorizar"
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
         Left            =   6120
         Picture         =   "frmAutRema.frx":A17E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
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
         Left            =   5640
         Picture         =   "frmAutRema.frx":CB50
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4560
         MaxLength       =   15
         TabIndex        =   4
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   7080
         TabIndex        =   12
         Top             =   3480
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   7320
         TabIndex        =   11
         Top             =   3480
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5880
         MaxLength       =   15
         TabIndex        =   7
         Top             =   5400
         Width           =   3135
      End
      Begin MSComctlLib.ListView ListaSelec 
         Height          =   2055
         Left            =   120
         TabIndex        =   6
         Top             =   4320
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3625
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
      Begin MSComctlLib.ListView ListaPendientes 
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4048
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
      Begin VB.Label Label4 
         Caption         =   "...o escriba aquí el numero de comanda."
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
         Left            =   4560
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "1)   Seleccione la comanda de la lista..."
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
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "2) Lista de articulos seleccionados"
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
         Left            =   120
         TabIndex        =   20
         Top             =   3960
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   495
         Left            =   5880
         TabIndex        =   19
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Tel."
         Height          =   255
         Left            =   5880
         TabIndex        =   18
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label lblNombre 
         Height          =   495
         Left            =   6720
         TabIndex        =   17
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label lblTel 
         Height          =   255
         Left            =   6720
         TabIndex        =   16
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label Label7 
         Caption         =   "Cantidad:"
         Height          =   375
         Left            =   3720
         TabIndex        =   14
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Autorizo/Denego"
         Height          =   255
         Left            =   5880
         TabIndex        =   13
         Top             =   5160
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAutRema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlCom As String
Dim sqlComDet As String
Dim tRs As ADODB.Recordset
Dim tLi As ListItem
Private Sub CmdAceptar_Click()
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim sqlQuery As String
    Dim NoSirve As Integer
    Dim tRs As ADODB.Recordset
    If ListaSelec.ListItems.Count > 0 Then
        For Cont = 1 To ListaSelec.ListItems.Count
            NoSirve = ListaSelec.ListItems.Item(Cont).SubItems(12) - ListaSelec.ListItems.Item(Cont).SubItems(1)
            sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'A', CANTIDAD = " & ListaSelec.ListItems.Item(Cont).SubItems(1) & ",CANTIDAD_NO_SIRVIO  = " & NoSirve & " WHERE ID_COMANDA = " & ListaSelec.ListItems.Item(Cont) & " AND ARTICULO = '" & ListaSelec.ListItems.Item(Cont).SubItems(11) & "'"
            cnn.Execute (sqlQuery)
            sqlQuery = "INSERT INTO AUTORIZAREMA (NOMBRE, ACCION, ID_COMANDA, FECHA) VALUES( '" & Text3.Text & "', 'AUTORIZO', " & ListaSelec.ListItems.Item(Cont) & ", '" & Format(Date, "dd/mm/yyyy") & "');"
            cnn.Execute (sqlQuery)
            sqlQuery = "SELECT ID_PRODUCTO, CANTIDAD FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & ListaSelec.ListItems(Cont).SubItems(2) & "' AND ID_PRODUCTO NOT IN (SELECT ID_PRODUCTO FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & Replace(ListaSelec.ListItems(Cont).SubItems(2), "REM", "REC") & "')"
            Set tRs = cnn.Execute(sqlQuery)
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not tRs.EOF
                    sqlQuery = "INSERT INTO JR_TEMPORALES (ID_COMANDA, ID_REPARACION, ID_PRODUCTO, CANTIDAD) VALUES (" & ListaSelec.ListItems.Item(Cont) & ", '" & ListaSelec.ListItems(Cont).SubItems(2) & "', '" & tRs.Fields("ID_PRODUCTO") & "', '" & tRs.Fields("CANTIDAD") & "');"
                    cnn.Execute (sqlQuery)
                    tRs.MoveNext
                Loop
            End If
        Next Cont
        '//////TICKET DE AUTORIZACION
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
        Printer.Print "SUCURSAL : " & VarMen.Text4(0).Text
        Printer.Print "ATENDIDO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
        Printer.Print "CLIENTE : " & lblNombre.Caption
        Printer.Print "AUTORIZADO POR : " & Text3.Text
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                          AUTORIZACION REMANOFACTURA"
        Printer.Print "--------------------------------------------------------------------------------"
        POSY = 2200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 1900
        Printer.Print "Cant."
        Printer.CurrentY = POSY
        Printer.CurrentX = 3000
        Printer.Print "Comanda"
        For Cont = 1 To ListaSelec.ListItems.Count
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print ListaSelec.ListItems(Cont).SubItems(2)
            Printer.CurrentY = POSY
            Printer.CurrentX = 1900
            Printer.Print ListaSelec.ListItems(Cont).SubItems(1)
            Printer.CurrentY = POSY
            Printer.CurrentX = 2900
            Printer.Print ListaSelec.ListItems(Cont).Text
        Next Cont
        Printer.Print ""
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.EndDoc
        ListaSelec.ListItems.Clear
        lblNombre.Caption = ""
        lblTel.Caption = ""
        Text3.Text = ""
    Else
        MsgBox "DEBE SELECCIONAR PRIMERO LOS PRODUCTOS", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdActualizar_Click()
On Error GoTo ManejaError
    Dim tRs As ADODB.Recordset
    Dim sBusqueda As String
    Dim tLi As ListItem
    sBusqueda = "SELECT * FROM VSCOMREMA WHERE ESTADO_ACTUAL = 'Z' AND ID_SUCURSAL = " & VarMen.Text4(6).Text
    Set tRs = cnn.Execute(sBusqueda)
    ListaPendientes.ListItems.Clear
    ListaSelec.ListItems.Clear
    lblNombre.Caption = ""
    lblTel.Caption = ""
    With tRs
        If .EOF And .BOF Then
            MsgBox "No hay articulos pendientes de Autorizacion"
        Else
            Do While Not .EOF
                Set tLi = ListaPendientes.ListItems.Add(, , .Fields("ID_COMANDA") & "")
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = .Fields("ID_PRODUCTO") & ""
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(3) = .Fields("NOMBRE") & ""
                If Not IsNull(.Fields("TELEFONO")) Then tLi.SubItems(4) = .Fields("TELEFONO") & ""
                If Not IsNull(.Fields("DIRECCION")) Then tLi.SubItems(5) = .Fields("DIRECCION") & ""
                If Not IsNull(.Fields("NUMERO_EXTERIOR")) Then tLi.SubItems(6) = .Fields("NUMERO_EXTERIOR") & ""
                If Not IsNull(.Fields("NUMERO_INTERIOR")) Then tLi.SubItems(7) = .Fields("NUMERO_INTERIOR") & ""
                If Not IsNull(.Fields("COLONIA")) Then tLi.SubItems(8) = .Fields("COLONIA") & ""
                If Not IsNull(.Fields("CIUDAD")) Then tLi.SubItems(9) = .Fields("CIUDAD") & ""
                If Not IsNull(.Fields("ESTADO")) Then tLi.SubItems(10) = .Fields("ESTADO") & ""
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdDenegar_Click()
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim sqlQuery As String
    If ListaSelec.ListItems.Count > 0 Then
        For Cont = 1 To ListaSelec.ListItems.Count
            sqlQuery = "UPDATE COMANDAS_DETALLES_2 SET ESTADO_ACTUAL = 'C' WHERE ID_COMANDA = " & ListaSelec.ListItems.Item(Cont) & " AND ARTICULO = '" & ListaSelec.ListItems.Item(Cont).SubItems(11) & "'" ' PONER BIEN ESTADO
            cnn.Execute (sqlQuery)
            sqlQuery = "INSERT INTO AUTORIZAREMA (NOMBRE, ACCION, ID_COMANDA, FECHA) VALUES( '" & Text3.Text & "', 'DENEGO', " & ListaSelec.ListItems.Item(Cont) & ", '" & Format(Date, "dd/mm/yyyy") & "');"
            cnn.Execute (sqlQuery)
        Next Cont
        ListaSelec.ListItems.Clear
        lblNombre.Caption = ""
        lblTel.Caption = ""
    Else
        MsgBox "DEBE SELECCIONAR PRIMERO LOS PRODUCTOS", vbInformation, "SACC"
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdNoComOk_Click()
On Error GoTo ManejaError
    Dim tRs As ADODB.Recordset
    Dim sBusqueda As String
    Dim tLi As ListItem
    sBusqueda = "SELECT * FROM VSCOMREMA WHERE ESTADO_ACTUAL = 'Z' AND ID_SUCURSAL = '" & VarMen.Text4(6).Text & "' AND ID_COMANDA = '" & txtNoCom.Text & "'"
    Set tRs = cnn.Execute(sqlQuery)
    ListaSelec.ListItems.Clear
    With tRs
        If .EOF And .BOF Then
            MsgBox "No hay articulos pendientes de Autorizacion"
        Else
            Do While Not .EOF
                Set tLi = ListaSelec.ListItems.Add(, , .Fields("ID_COMANDA") & "")
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = .Fields("ID_PRODUCTO") & ""
                If Not IsNull(.Fields("NOMBRE_COMERCIAL")) Then tLi.SubItems(3) = .Fields("NOMBRE_COMERCIAL") & ""
                If Not IsNull(.Fields("TELEFONO_CASA")) Then tLi.SubItems(5) = .Fields("TELEFONO_CASA") & ""
                If Not IsNull(.Fields("DIRECCION")) Then tLi.SubItems(7) = .Fields("DIRECCION") & ""
                If Not IsNull(.Fields("NUMERO_EXTERIOR")) Then tLi.SubItems(8) = .Fields("NUMERO_EXTERIOR") & ""
                If Not IsNull(.Fields("NUMERO_INTERIOR")) Then tLi.SubItems(9) = .Fields("NUMERO_INTERIOR") & ""
                If Not IsNull(.Fields("COLONIA")) Then tLi.SubItems(10) = .Fields("COLONIA") & ""
                If Not IsNull(.Fields("CIUDAD")) Then tLi.SubItems(11) = .Fields("CIUDAD") & ""
                If Not IsNull(.Fields("ESTADO")) Then tLi.SubItems(11) = .Fields("ESTADO") & ""
                .MoveNext
            Loop
        End If
    End With
    lblNombre.Caption = ""
    lblTel.Caption = ""
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim tLi As ListItem
    Dim Selec As Integer
    Selec = Val(Text2(1).Text)
    If (lblNombre.Caption = ListaPendientes.ListItems.Item(Selec).SubItems(3)) Or (lblNombre.Caption = "") Then
        If Val(ListaPendientes.ListItems.Item(Selec).SubItems(1)) >= Val(Text1.Text) Then
            Set tLi = ListaSelec.ListItems.Add(, , ListaPendientes.ListItems.Item(Selec))
            tLi.SubItems(1) = Text1.Text
            tLi.SubItems(2) = ListaPendientes.ListItems.Item(Selec).SubItems(2)
            tLi.SubItems(3) = ListaPendientes.ListItems.Item(Selec).SubItems(3)
            tLi.SubItems(4) = ListaPendientes.ListItems.Item(Selec).SubItems(4)
            tLi.SubItems(5) = ListaPendientes.ListItems.Item(Selec).SubItems(5)
            tLi.SubItems(6) = ListaPendientes.ListItems.Item(Selec).SubItems(6)
            tLi.SubItems(7) = ListaPendientes.ListItems.Item(Selec).SubItems(7)
            tLi.SubItems(8) = ListaPendientes.ListItems.Item(Selec).SubItems(8)
            tLi.SubItems(9) = ListaPendientes.ListItems.Item(Selec).SubItems(9)
            tLi.SubItems(10) = ListaPendientes.ListItems.Item(Selec).SubItems(10)
            tLi.SubItems(11) = ListaPendientes.ListItems.Item(Selec).SubItems(11)
            tLi.SubItems(12) = ListaPendientes.ListItems.Item(Selec).SubItems(1)
            lblNombre.Caption = ListaPendientes.ListItems.Item(Selec).SubItems(3)
            lblTel.Caption = ListaPendientes.ListItems.Item(Selec).SubItems(4)
            ListaPendientes.ListItems.Remove (Selec)
            Label6.Caption = ""
            Text1.Text = ""
        Else
            MsgBox "LA CANTIDAD AUTORIZADA NO PUEDE SER MAYOR QUE LA CANTIDAD SOLICITADA", vbInformation, "SACC"
        End If
    Else
        MsgBox "NO SE PUEDE AUTORIZAR REMANOFACTURAS DE DIFERENTES CLIENTES AL MISMO TIEMPO", vbInformation, "SACC"
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    Dim tRs As ADODB.Recordset
    Dim sBusqueda As String
    Dim tLi As ListItem
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    lblNombre.Caption = ""
    lblTel.Caption = ""
    With ListaPendientes
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "COMANDA", 1200
        .ColumnHeaders.Add , , "CANTIDAD", 1200
        .ColumnHeaders.Add , , "ID_PROD", 1200
        .ColumnHeaders.Add , , "CLIENTE", 2600
        .ColumnHeaders.Add , , "TELEFONO", 1200
        .ColumnHeaders.Add , , "DIRECCION", 3200
        .ColumnHeaders.Add , , "NO. INT.", 800
        .ColumnHeaders.Add , , "NO. EXT.", 800
        .ColumnHeaders.Add , , "COLONIA", 1200
        .ColumnHeaders.Add , , "CIUDAD", 1200
        .ColumnHeaders.Add , , "ESTADO", 1200
        .ColumnHeaders.Add , , "ARTICULO", 0
    End With
    With ListaSelec
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "COMANDA", 1200
        .ColumnHeaders.Add , , "CANTIDAD", 1200
        .ColumnHeaders.Add , , "ID_PROD", 1200
        .ColumnHeaders.Add , , "CLIENTE", 2600
        .ColumnHeaders.Add , , "TELEFONO", 1200
        .ColumnHeaders.Add , , "DIRECCION", 3200
        .ColumnHeaders.Add , , "NO. EXT.", 800
        .ColumnHeaders.Add , , "NO. INT.", 800
        .ColumnHeaders.Add , , "COLONIA", 1200
        .ColumnHeaders.Add , , "CIUDAD", 1200
        .ColumnHeaders.Add , , "ESTADO", 1200
        .ColumnHeaders.Add , , "ARTICULO", 0
        .ColumnHeaders.Add , , "CANTORIG", 0
    End With
    sBusqueda = "SELECT * FROM VSCOMREMA WHERE ESTADO_ACTUAL = 'Z' AND ID_SUCURSAL = " & VarMen.Text4(6).Text
    Set tRs = cnn.Execute(sBusqueda)
    With tRs
        If .EOF And .BOF Then
            MsgBox "No hay articulos pendientes de Autorizacion"
        Else
            Do While Not .EOF
             Set tLi = ListaPendientes.ListItems.Add(, , .Fields("ID_COMANDA") & "")
                    If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                    If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = .Fields("ID_PRODUCTO") & ""
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(3) = .Fields("NOMBRE") & ""
                    If Not IsNull(.Fields("TELEFONO")) Then tLi.SubItems(4) = .Fields("TELEFONO") & ""
                    If Not IsNull(.Fields("DIRECCION")) Then tLi.SubItems(5) = .Fields("DIRECCION") & ""
                    If Not IsNull(.Fields("NUMERO_EXTERIOR")) Then tLi.SubItems(6) = .Fields("NUMERO_EXTERIOR") & ""
                    If Not IsNull(.Fields("NUMERO_INTERIOR")) Then tLi.SubItems(7) = .Fields("NUMERO_INTERIOR") & ""
                    If Not IsNull(.Fields("COLONIA")) Then tLi.SubItems(8) = .Fields("COLONIA") & ""
                    If Not IsNull(.Fields("CIUDAD")) Then tLi.SubItems(9) = .Fields("CIUDAD") & ""
                    If Not IsNull(.Fields("ESTADO")) Then tLi.SubItems(10) = .Fields("ESTADO") & ""
                    If Not IsNull(.Fields("ARTICULO")) Then tLi.SubItems(11) = .Fields("ARTICULO") & ""
                    .MoveNext
            Loop
        End If
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
Private Sub ListaPendientes_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListaPendientes.ListItems.Count > 0 Then
        Label6.Caption = Item.SubItems(2)
        Text1.Text = Item.SubItems(1)
        Text2(0).Text = Item
        Text2(1).Text = Item.Index
    End If
End Sub
Private Sub Text3_Change()
    cmdAceptar.Enabled = False
    cmdDenegar.Enabled = False
    If Text3.Text <> "" Then
        cmdAceptar.Enabled = True
        cmdDenegar.Enabled = True
    End If
End Sub
Private Sub txtNoCom_Change()
On Error GoTo ManejaError
    If txtNoCom.Text <> "" Then
        cmdNoComOk.Enabled = True
    Else
        cmdNoComOk.Enabled = False
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtNoCom_GotFocus()
On Error GoTo ManejaError
    txtNoCom.SelStart = 0
    txtNoCom.SelLength = Len(txtNoCom.Text)
    txtNoCom.BackColor = &HFFE1E1
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtNoCom_LostFocus()
On Error GoTo ManejaError
    txtNoCom.BackColor = &H80000005
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtNoCom_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And txtNoCom.Text <> "" Then
        cmdNoComOk.Value = True
    End If
    Dim Valido As String
    Valido = "1234567890"
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
