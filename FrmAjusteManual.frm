VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmAjusteManual 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ver Estado de Pedidos de Clientes - (Ordenes de Compra)"
   ClientHeight    =   6510
   ClientLeft      =   720
   ClientTop       =   1035
   ClientWidth     =   9855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame20 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   15
      Top             =   1560
      Width           =   975
      Begin VB.Image Image19 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmAjusteManual.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmAjusteManual.frx":030A
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Manual"
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
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   13
      Top             =   3960
      Width           =   975
      Begin VB.Image Image6 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmAjusteManual.frx":1DBC
         MousePointer    =   99  'Custom
         Picture         =   "FrmAjusteManual.frx":20C6
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label1 
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
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame26 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   11
      Top             =   2760
      Width           =   975
      Begin VB.Image Image24 
         Height          =   765
         Left            =   240
         MouseIcon       =   "FrmAjusteManual.frx":3B78
         MousePointer    =   99  'Custom
         Picture         =   "FrmAjusteManual.frx":3E82
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ajuste Automatico"
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   8760
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   8880
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   9000
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   9120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   9240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   3
      Top             =   5160
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmAjusteManual.frx":5910
         MousePointer    =   99  'Custom
         Picture         =   "FrmAjusteManual.frx":5C1A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Pedidos"
      TabPicture(0)   =   "FrmAjusteManual.frx":7CFC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin MSComctlLib.ListView ListView2 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   3600
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4471
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4683
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
End
Attribute VB_Name = "FrmAjusteManual"
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
        Else
        ' AQUI VA EL MERO MOLE... LO SABROSO DE ESTO.... EL CODIGO QUE HACE TODO EL DESMADRE... WACHA!!!
            Dim NoReg As Integer
            NoReg = ListView2.ListItems.Count
            Dim Con As Integer
            Dim IDPro As String
            For Con = 1 To NoReg
                IDPro = ListView2.ListItems(Con)
                sqlComanda = "SELECT * FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & IDPro & "' AND CANTIDAD >= " & ListView2.ListItems(Con).SubItems(3) & " AND SUCURSAL = 'BODEGA'"
                Set tRs = cnn.Execute(sqlComanda)
                If Not (tRs.BOF And tRs.EOF) Then
                    Dim NewEx As Double
                    NewEx = tRs.Fields("CANTIDAD")
                    NewEx = NewEx - CDbl(ListView2.ListItems(Con).SubItems(3))
                    sqlComanda = "UPDATE EXISTENCIAS SET CANTIDAD = " & NewEx & " WHERE ID_PRODUCTO = '" & ListView2.ListItems(Con) & "' AND SUCURSAL = 'BODEGA'"
                    Set tRs = cnn.Execute(sqlComanda)
                    sqlComanda = "UPDATE PED_CLIEN_DETALLE SET CANTIDAD_PENDIENTE = 0 WHERE ID_PRODUCTO = '" & ListView2.ListItems(Con) & "' AND NO_PEDIDO = " & NoPed
                    Set tRs = cnn.Execute(sqlComanda)
                End If
            Next Con
        End If
    Next Conta
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
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
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. Pedido", 1000
        .ColumnHeaders.Add , , "Capturo", 1500
        .ColumnHeaders.Add , , "Cliente", 6500
        .ColumnHeaders.Add , , "Fecha", 1500
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
        .ColumnHeaders.Add , , "Clave del Producto", 4500
        .ColumnHeaders.Add , , "Cantidad Pedida", 2000
        .ColumnHeaders.Add , , "Cantidad en Existencia", 2000
        .ColumnHeaders.Add , , "Cantidad Pendiente", 2000
    End With
    Actualizar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image19_Click()
    If CantPed > 0 And CantSurt > 0 Then
        On Error GoTo ManejaError
        Text1(0).Text = SelMod
        Text1(1).Text = CantPed
        Text1(2).Text = NoPed
        FrmModSurt.Show vbModal, Me
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image24_Click()
    On Error GoTo ManejaError
    Actualizar
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Image6_Click()
    If CantPed > 0 And CantSurt > 0 Then
        On Error GoTo ManejaError
        Text1(4).Text = CantSurt
        CantSurt = CantSurt - CantPed
        Text1(0).Text = Trim(SelMod)
        Text1(1).Text = CantSurt
        Text1(2).Text = NoPed
        Text1(3).Text = CantPed
        FrmDeshacer2.Show vbModal
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
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    NoPed = Item
    Dim sBuscar As String
    Dim tRs3 As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT * FROM PED_CLIEN_DETALLE WHERE NO_PEDIDO = " & CDbl(Item)
    Set tRs3 = cnn.Execute(sBuscar)
    If (tRs3.BOF And tRs3.EOF) Then
        'MsgBox "El pedido esta vacio", vbCritical, "ERROR"
        Me.ListView2.ListItems.Clear
    Else
        ListView2.ListItems.Clear
        'tRs3.MoveFirst
        Do While Not tRs3.EOF
            Set tLi = ListView2.ListItems.Add(, , Trim(tRs3.Fields("ID_PRODUCTO")) & "")
            tLi.SubItems(1) = tRs3.Fields("CANTIDAD_PEDIDA") & ""
            tLi.SubItems(2) = tRs3.Fields("CANTIDAD_EXISTENCIA") & ""
            tLi.SubItems(3) = tRs3.Fields("CANTIDAD_PENDIENTE") & ""
            tRs3.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    SelMod = Item
    CantPed = Item.SubItems(3)
    CantSurt = Item.SubItems(1)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Actualizar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim BusClie As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT * FROM PED_CLIEN WHERE ESTADO = 'I'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            ListView1.ListItems.Clear
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("NO_PEDIDO") & "")
                tLi.SubItems(1) = .Fields("USUARIO") & ""
                BusClie = "SELECT NOMBRE FROM CLIENTE WHERE ID_CLIENTE =" & .Fields("ID_CLIENTE")
                Set tRs2 = cnn.Execute(BusClie)
                tLi.SubItems(2) = tRs2.Fields("NOMBRE") & ""
                tLi.SubItems(3) = .Fields("FECHA") & ""
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
