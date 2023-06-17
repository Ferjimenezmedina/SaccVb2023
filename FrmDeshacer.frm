VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmDeshacer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deshacer (Regresar a Existencia)"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6000
      TabIndex        =   19
      Top             =   2880
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
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmDeshacer.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmDeshacer.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmDeshacer.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label16"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label15"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label7"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label8"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label9"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DTPicker1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
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
         Left            =   3360
         Picture         =   "FrmDeshacer.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   3360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   3480
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Format          =   50724865
         CurrentDate     =   39127
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   1920
         Width           =   3375
      End
      Begin VB.Label Label8 
         Caption         =   "Cantidad en Existencia :"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad Surtida :"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Numero de Orden :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Clave del Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Cantidad Pendiente :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Cantidad Pedida :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Label Label15 
         Height          =   135
         Left            =   4680
         TabIndex        =   5
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label16 
         Height          =   135
         Left            =   4680
         TabIndex        =   4
         Top             =   3600
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      Height          =   255
      Left            =   6000
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   255
      Left            =   6000
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "FrmDeshacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim Pend As Double
    Dim NuevaExis As Double
    Dim sqlComanda As String
    Dim tRs As ADODB.Recordset
    Dim descrip As String
    Dim Marca As String
    Dim Urgente As String
    Dim sBuscar2 As String
    Dim sBuscar As String
    If CDbl(Text1.Text) > CDbl(Label7.Caption) Then
        MsgBox "IMPOSIBLE REGRESAR MAS DE LO APARTADO!", vbInformation, "SACC"
    Else
        Pend = CDbl(Label12.Caption) + CDbl(Text1.Text)
        sqlComanda = "UPDATE PED_CLIEN_DETALLE SET CANTIDAD_PENDIENTE = " & Pend & " WHERE ID_PRODUCTO = '" & Label6.Caption & "' AND NO_PEDIDO = " & Label5.Caption
        Set tRs = cnn.Execute(sqlComanda)
        sqlComanda = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Label6.Caption & "' AND SUCURSAL = '" & frmShowPediC.Combo1.Text & "'"
        Set tRs = cnn.Execute(sqlComanda)
        If tRs.EOF And tRs.BOF Then
            NuevaExis = Text1.Text
            sBuscar = "INSERT INTO EXISTENCIAS (CANTIDAD, SUCURSAL, ID_PRODUCTO) VALUES (" & NuevaExis & ", '" & frmShowPediC.Combo1.Text & "', '" & Label6.Caption & "' );"
            cnn.Execute (sBuscar)
        Else
            NuevaExis = CDbl(tRs.Fields("CANTIDAD")) + CDbl(Text1.Text)
            sqlComanda = "UPDATE EXISTENCIAS SET CANTIDAD = " & NuevaExis & " WHERE ID_PRODUCTO = '" & Label6.Caption & "' AND SUCURSAL = '" & frmShowPediC.Combo1.Text & "'"
            Set tRs = cnn.Execute(sqlComanda)
        End If
        If MsgBox("QUIERE HACER UNA REQUISICION POR " & Text1.Text & "? " & Label6.Caption, vbYesNo, "SACC") = vbYes Then
            sBuscar = "SELECT Descripcion, MARCA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & Label6.Caption & "'"
            Set tRs = cnn.Execute(sBuscar)
            descrip = ""
            If Not (tRs.BOF And tRs.EOF) Then
                descrip = tRs.Fields("Descripcion")
                Marca = tRs.Fields("MARCA")
            End If
            If Label15.Caption = "COMPUESTO" Then
                sBuscar = "INSERT INTO PEDIDO (SUCURSAL, PIDIO, FECHA, TIPO, COMENTARIO) VALUES ('" & VarMen.Text4(0).Text & "', '1', '" & Format(Date, "dd/mm/yyyy") & "', 'D', 'REPOSISCION DE PEDIDO')"
                cnn.Execute (sBuscar)
                sBuscar = "SELECT ID_PEDIDO FROM PEDIDO WHERE SUCURSAL = '" & VarMen.Text4(0).Text & "' ORDER BY ID_PEDIDO DESC"
                Set tRs = cnn.Execute(sBuscar)
                sBuscar = "INSERT INTO DETALLE_PEDIDO (ID_PEDIDO, CANTIDAD, ID_PRODUCTO, ENTREGADO, DESCRIPCION, ALMACEN, MARCA) VALUES ('" & tRs.Fields("ID_PEDIDO") & "', '" & Replace(Text1.Text, ",", "") & "', '" & Label6.Caption & "', 0, '" & descrip & "', 'A3', '" & Marca & "')"
                cnn.Execute (sBuscar)
            Else
                sBuscar = "SELECT ID_REQUISICION, CANTIDAD FROM REQUISICION WHERE ACTIVO = 0 AND URGENTE = 'S' AND ID_PRODUCTO = '" & Label6.Caption & "'"
                Set tRs = cnn.Execute(sBuscar)
                If tRs.BOF And tRs.EOF Or (DTPicker1.Value <= (Date + 10)) Then
                    Urgente = "N"
                    If DTPicker1.Value <= (Date + 10) Then Urgente = "S"
                    sBuscar = "SELECT COMENTARIO FROM PED_CLIEN WHERE NO_PEDIDO = " & Pend
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        sBuscar = "INSERT INTO REQUISICION (FECHA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, ACTIVO, CONTADOR, COTIZADA, ALMACEN, URGENTE, COMENTARIO) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', '" & Label6.Caption & "' , '" & descrip & "'," & Replace(Text1.Text, ",", "") & ", 0, 0, 0, 'A3', '" & Urgente & "', '" & tRs.Fields("COMENTARIO") & "')"
                        cnn.Execute (sBuscar)
                    Else
                        sBuscar = "INSERT INTO REQUISICION (FECHA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, ACTIVO, CONTADOR, COTIZADA, ALMACEN, URGENTE, COMENTARIO) VALUES ('" & Format(Date, "dd/mm/yyyy") & "', '" & Label6.Caption & "' , '" & descrip & "'," & Replace(Text1.Text, ",", "") & ", 0, 0, 0, 'A3', '" & Urgente & "', 'VENTA PROGRAMADA No. " & Pend & " PEDIDO DEL SISTEMA')"
                        cnn.Execute (sBuscar)
                    End If
                    sBuscar = "SELECT ID_REQUISICION FROM REQUISICION ORDER BY ID_REQUISICION DESC"
                    Set tRs = cnn.Execute(sBuscar)
                    sBuscar2 = "INSERT INTO RASTREOREQUI (ID_REQUI, SOLICITO, CANTIDAD, COMENTARIO,FECHA) Values(" & tRs.Fields("ID_REQUISICION") & ", '" & VarMen.Text1(1).Text & "'  ," & Replace(Text1.Text, ",", "") & ", 'V.Prog. No. " & Label5.Caption & " No.Orden " & Label17.Caption & " Cliente: " & Label14.Caption & "','" & Format(Date, "dd/mm/yyyy") & "')"
                    cnn.Execute (sBuscar2)
                Else
                    sBuscar = "UPDATE REQUISICION SET CANTIDAD = " & Replace(Val(tRs.Fields("CANTIDAD")) + Val(Text1.Text), ",", "") & "WHERE ID_REQUISICION = " & tRs.Fields("ID_REQUISICION") & " AND ID_PRODUCTO = '" & Label6.Caption & "'"
                    cnn.Execute (sBuscar)
                    sBuscar2 = "INSERT INTO RASTREOREQUI (ID_REQUI, SOLICITO, CANTIDAD, COMENTARIO,FECHA) Values(" & tRs.Fields("ID_REQUISICION") & ", '" & VarMen.Text1(1).Text & "'  ," & Replace(Text1.Text, ",", "") & ", '" & "V.Prog. No. " & Label5.Caption & " No.Orden " & Label17.Caption & " Cliente: " & Label14.Caption & "','" & Format(Date, "dd/mm/yyyy") & "')"
                    cnn.Execute (sBuscar2)
                End If
                tRs.Close
            End If
        End If
        Command1.Enabled = False
        Unload Me
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Label5.Caption = frmShowPediC.Text1(2).Text
    Label6.Caption = frmShowPediC.Text1(0).Text
    Label7.Caption = frmShowPediC.Text1(1).Text
    Label12.Caption = frmShowPediC.Text1(3).Text
    Label13.Caption = frmShowPediC.Text1(4).Text
    Label14.Caption = frmShowPediC.Label2.Caption
    Label15.Caption = frmShowPediC.Label3.Caption
    DTPicker1.Value = frmShowPediC.DTPicker1.Value
    Command1.Enabled = False
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.txtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE SUCURSAL = '" & frmShowPediC.Combo1.Text & "' AND ID_PRODUCTO = '" & Label6.Caption & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF Or tRs.EOF) Then
        Label9.Caption = tRs.Fields("CANTIDAD") & ""
    Else
        Label9.Caption = "0"
    End If
    sBuscar = "SELECT TIPO, Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO = '" & Label6.Caption & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF Or tRs.EOF) Then
        Label15.Caption = tRs.Fields("TIPO")
        Label16.Caption = tRs.Fields("Descripcion")
    Else
        Label15.Caption = "X"
        Label16.Caption = ""
    End If
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
Private Sub Text1_Change()
On Error GoTo ManejaError
    If Text1.Text = "" Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_GotFocus()
On Error GoTo ManejaError
    Text1.BackColor = &HFFE1E1
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    Dim Valido As String
    Valido = "1234567890."
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
