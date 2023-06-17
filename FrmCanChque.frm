VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCanChque 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelar Cheques a Orden de Compra"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuscar 
      Cancel          =   -1  'True
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
      Left            =   4320
      Picture         =   "FrmCanChque.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7440
      TabIndex        =   10
      Top             =   3240
      Width           =   975
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   15
         Top             =   1320
         Width           =   975
         Begin VB.Image Image15 
            Height          =   720
            Left            =   120
            MouseIcon       =   "FrmCanChque.frx":29D2
            MousePointer    =   99  'Custom
            Picture         =   "FrmCanChque.frx":2CDC
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Guardar"
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
            TabIndex        =   16
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   13
         Top             =   1320
         Width           =   975
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "FrmCanChque.frx":469E
            MousePointer    =   99  'Custom
            Picture         =   "FrmCanChque.frx":49A8
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   11
         Top             =   0
         Width           =   975
         Begin VB.Label Label18 
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
            TabIndex        =   12
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image17 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmCanChque.frx":61D2
            MousePointer    =   99  'Custom
            Picture         =   "FrmCanChque.frx":64DC
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Eliminar"
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
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image18 
         Height          =   750
         Left            =   120
         MouseIcon       =   "FrmCanChque.frx":7F8E
         MousePointer    =   99  'Custom
         Picture         =   "FrmCanChque.frx":8298
         Top             =   240
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmCanChque.frx":9FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ListView1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3413
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
         Caption         =   "Rapida"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Indirecta"
         Height          =   255
         Left            =   5520
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Internacional"
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nacional"
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1935
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3413
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Abonos Cancelados"
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
         Left            =   240
         TabIndex        =   21
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Abonos Activos"
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
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "No. Orden :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7440
      TabIndex        =   0
      Top             =   4440
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCanChque.frx":9FDE
         MousePointer    =   99  'Custom
         Picture         =   "FrmCanChque.frx":A2E8
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmCanChque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim num_orden As String
Dim sTipo As String
Dim sTipoOrden As String
Dim sBuscaTodo As String
Dim NumOrdenes As String
Dim IdProveedor As String
Private Sub cmdBuscar_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tLi As ListItem
    num_orden = Text1.Text
    If Option1.Value Then
        sTipo = "N"
        sTipoOrden = "NACIONAL"
    End If
    If Option2.Value Then
        sTipo = "I"
        sTipoOrden = "INTERNACIONAL"
    End If
    If Option3.Value Then
        sTipo = "X"
        sTipoOrden = "INDIRECTA"
    End If
    If Option4.Value Then
        sTipo = "R"
        sTipoOrden = "RAPIDA"
    End If
    ListView1.ListItems.Clear
    sBuscaTodo = "SELECT ID_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN = " & Text1.Text & " AND TIPO = '" & sTipo & "'"
    sBuscar = "SELECT ID_ABONO, FECHA, BANCO, NUMTRANS, NUMCHEQUE, CANT_ABONO FROM ABONOS_PAGO_OC WHERE NUM_ORDEN = " & Text1.Text & " AND TIPO = '" & sTipo & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , Format(CDbl(tRs.Fields("CANT_ABONO")), "###,###,###,##0.00"))
            If Not IsNull(tRs.Fields("BANCO")) Then tLi.SubItems(1) = tRs.Fields("BANCO")
            If Not IsNull(tRs.Fields("NUMTRANS")) Then tLi.SubItems(2) = tRs.Fields("NUMTRANS")
            If Not IsNull(tRs.Fields("NUMCHEQUE")) Then tLi.SubItems(3) = tRs.Fields("NUMCHEQUE")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(4) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("ID_ABONO")) Then tLi.SubItems(5) = tRs.Fields("ID_ABONO")
            tRs.MoveNext
        Loop
    End If
    ListView2.ListItems.Clear
    sBuscar = "SELECT USR_CANCELO, FECHA, BANCO, NUMTRANS, NUMCHEQUE, CANT_ABONO FROM ABONOS_OC_CANCELADOS WHERE NUM_ORDEN = " & Text1.Text & " AND TIPO = '" & sTipo & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , Format(CDbl(tRs.Fields("CANT_ABONO")), "###,###,###,##0.00"))
            If Not IsNull(tRs.Fields("BANCO")) Then tLi.SubItems(1) = tRs.Fields("BANCO")
            If Not IsNull(tRs.Fields("NUMTRANS")) Then tLi.SubItems(2) = tRs.Fields("NUMTRANS")
            If Not IsNull(tRs.Fields("NUMCHEQUE")) Then tLi.SubItems(3) = tRs.Fields("NUMCHEQUE")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(4) = tRs.Fields("FECHA")
            sBuscar = "SELECT NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs.Fields("USR_CANCELO")
            Set tRs1 = cnn.Execute(sBuscar)
            If Not (tRs1.EOF And tRs1.BOF) Then
                tLi.SubItems(5) = tRs1.Fields("NOMBRE") & " " & tRs1.Fields("APELLIDOS")
            Else
                tLi.SubItems(5) = "<USUARIO NO ENCONTRADO>"
            End If
            tRs.MoveNext
        Loop
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
On Error GoTo ManejaError
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
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .Checkboxes = True
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ABONO", 1200
        .ColumnHeaders.Add , , "BANCO", 1200
        .ColumnHeaders.Add , , "No. TRANSFERENCIA", 1200
        .ColumnHeaders.Add , , "NO. CHEQUE", 1200
        .ColumnHeaders.Add , , "FECHA", 1200
        .ColumnHeaders.Add , , "ID_ABONO", 0
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .Checkboxes = True
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ABONO", 1200
        .ColumnHeaders.Add , , "BANCO", 1200
        .ColumnHeaders.Add , , "No. TRANSFERENCIA", 1200
        .ColumnHeaders.Add , , "NO. CHEQUE", 1200
        .ColumnHeaders.Add , , "FECHA", 1200
        .ColumnHeaders.Add , , "USUARIO", 1200
    End With
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image18_Click()
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim sTipoC As String
    For Cont = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(Cont).Checked Then
            'sBuscar = "SELECT * FROM ABONOS_PAGO_OC WHERE ID_ABONO IN (" & ListView1.ListItems(Cont).SubItems(5) & ")"
            sBuscar = "SELECT * FROM ABONOS_PAGO_OC WHERE ID_ABONO IN (" & ListView1.ListItems(Cont).SubItems(5) & ")"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                IdProveedor = tRs.Fields("ID_PROVEEDOR")
                Do While Not tRs.EOF
                    sBuscar = "INSERT INTO ABONOS_OC_CANCELADOS (ID_ABONO, ID_PROVEEDOR, CANT_ABONO, NO_CHEQUE, BANCO, CANTIDAD, NUMTRANS, NUMCHEQUE, NUM_ORDEN, TIPO, ID_ORDEN, TIPOPAGO, FECHA, PROVEEDOR, USR_CANCELO) VALUES ('" & tRs.Fields("ID_ABONO") & "', '" & tRs.Fields("ID_PROVEEDOR") & "', '" & tRs.Fields("CANT_ABONO") & "', '" & tRs.Fields("NO_CHEQUE") & "', '" & tRs.Fields("BANCO") & "', '" & tRs.Fields("CANTIDAD") & "', '" & tRs.Fields("NUMTRANS") & "', '" & tRs.Fields("NUMCHEQUE") & "', '" & tRs.Fields("NUM_ORDEN") & "', '" & tRs.Fields("TIPO") & "', '" & tRs.Fields("ID_ORDEN") & "', '" & tRs.Fields("TIPOPAGO") & "', '" & tRs.Fields("FECHA") & "', '" & tRs.Fields("PROVEEDOR") & "', '" & VarMen.Text1(0).Text & "');"
                    cnn.Execute (sBuscar)
                    tRs.MoveNext
                Loop
            End If
            sBuscar = "DELETE FROM ABONOS_PAGO_OC WHERE ID_ABONO IN (" & ListView1.ListItems(Cont).SubItems(5) & ")"
            cnn.Execute (sBuscar)
            If sTipo = "N" Then
                sTipoC = "NACIONAL"
            End If
            If sTipo = "I" Then
                sTipoC = "INTERNACIONAL"
            End If
            If sTipo = "X" Then
                sTipoC = "INDIRECTA"
            End If
            If sTipo = "R" Then
                sTipoC = "RAPIDA"
            End If
            'sBuscar = "SELECT * FROM CHEQUES WHERE NUM_ORDEN LIKE '%" & num_orden & ", %'  AND TIPO_ORDEN = '" & sTipoC & "'"
            sBuscar = "SELECT * FROM CHEQUES WHERE (NUM_ORDEN LIKE '%, " & num_orden & ", %' OR NUM_ORDEN LIKE '" & num_orden & ", %') AND (TIPO_ORDEN = '" & sTipoC & "')"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                If sTipo <> "R" Then
                    sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'X' WHERE NUM_ORDEN IN (" & Mid(tRs.Fields("NUM_ORDEN"), 1, Len(tRs.Fields("NUM_ORDEN")) - 2) & ") AND TIPO = '" & sTipo & "'"
                    cnn.Execute (sBuscar)
                    NumOrdenes = Mid(tRs.Fields("NUM_ORDEN"), 1, Len(tRs.Fields("NUM_ORDEN")) - 2)
                    MsgBox "LAS ORDENES REGRESADAS SON LAS SIGUIENTES: " & Mid(tRs.Fields("NUM_ORDEN"), 1, Len(tRs.Fields("NUM_ORDEN")) - 2), vbExclamation, "SACC"
                    sBuscar = "SELECT EMAIL FROM PROVEEDOR WHERE ID_PROVEEDOR = " & IdProveedor & " AND EMAIL IS NOT NULL"
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        EnviaCorreo (tRs.Fields("EMAIL"))
                    End If
                Else
                    sBuscar = "UPDATE ORDEN_RAPIDA SET ESTADO = 'A' WHERE ID_ORDEN_RAPIDA IN (" & Mid(tRs.Fields("NUM_ORDEN"), 1, Len(tRs.Fields("NUM_ORDEN")) - 2) & ")"
                    cnn.Execute (sBuscar)
                    NumOrdenes = Mid(tRs.Fields("NUM_ORDEN"), 1, Len(tRs.Fields("NUM_ORDEN")) - 2)
                MsgBox "LAS ORDENES REGRESADAS SON LAS SIGUIENTES: " & Mid(tRs.Fields("NUM_ORDEN"), 1, Len(tRs.Fields("NUM_ORDEN")) - 2), vbExclamation, "SACC"
                    sBuscar = "SELECT EMAIL FROM PROVEEDOR_CONSUMO WHERE ID_PROVEEDOR = " & IdProveedor & " AND EMAIL IS NOT NULL"
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        EnviaCorreo (tRs.Fields("EMAIL"))
                    End If
                End If
            Else
                If sTipo <> "R" Then
                    sBuscar = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'X' WHERE NUM_ORDEN = " & num_orden & " AND TIPO = '" & sTipo & "'"
                    cnn.Execute (sBuscar)
                    sBuscar = "SELECT EMAIL FROM PROVEEDOR WHERE ID_PROVEEDOR = " & IdProveedor & " AND EMAIL IS NOT NULL"
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        EnviaCorreo (tRs.Fields("EMAIL"))
                    End If
                Else
                    sBuscar = "UPDATE ORDEN_RAPIDA SET ESTADO = 'A' WHERE ID_ORDEN_RAPIDA = " & num_orden
                    cnn.Execute (sBuscar)
                    sBuscar = "SELECT EMAIL FROM PROVEEDOR_CONSUMO WHERE ID_PROVEEDOR = " & IdProveedor & " AND EMAIL IS NOT NULL"
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        EnviaCorreo (tRs.Fields("EMAIL"))
                    End If
                End If
                NumOrdenes = num_orden
            End If
        End If
    Next Cont
    cmdBuscar.Value = True
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub EnviaCorreo(MAIL As String)
On Error GoTo ManejaError
    If GetSetting("APTONER", "ConfigSACC", "Correo", "") <> "" Then
        Dim email As CDO.Message
        Dim correo As String
        Dim passwd As String
        Dim destino As String
        Dim Asunto As String
        Dim cuerpo As String
        Set email = New CDO.Message
        correo = GetSetting("APTONER", "ConfigSACC", "Correo", "")  ' "sistemas2@aptoner.com.mx"
        passwd = GetSetting("APTONER", "ConfigSACC", "CorreoPass", "")  ' "@Pt171218." Contraseña Generada por Gmail
        destino = MAIL '"control.sistemas.aptoner@gmail.com"
        Asunto = "Reprogramación de Pago"
        cuerpo = "La empresa " & VarMen.TxtEmp(0).Text & " acaba de posponer el pago de la(s) orden(es) de compra " & NumOrdenes & ", esta(s) será(n) reptrogramada(s) para nueva fecha de pago."
        email.Configuration.Fields(cdoSMTPServer) = GetSetting("APTONER", "ConfigSACC", "SMTP", "")  '"aptoner.com.mx"
        email.Configuration.Fields(cdoSendUsingMethod) = 2
        With email.Configuration.Fields
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(GetSetting("APTONER", "ConfigSACC", "PuertoCorreo", ""))  ' 26
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = Abs(1)
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = correo
            .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = passwd
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
        End With
        With email
            .To = destino
            .From = correo
            .Subject = Asunto
            .TextBody = cuerpo
            .Configuration.Fields.Update
            .Send
        End With
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
