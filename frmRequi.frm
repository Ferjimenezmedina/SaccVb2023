VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmRequi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "REQUISICIONES"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   5895
      Left            =   5040
      TabIndex        =   25
      Top             =   0
      Width           =   4815
      Begin VB.OptionButton opnNomCom 
         Caption         =   "NOMBRE COMERCIAL"
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   5520
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton opnNom 
         Caption         =   "NOMBRE"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   5520
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComctlLib.ListView lvwL 
         Height          =   5175
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   9128
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "USUARIO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NOMBRE"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "NOMBRE COMERCIAL"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "PROVEEDOR:"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "CLIENTE:"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "EJECUTIVO:"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtNP 
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtNC 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtNE 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdVE 
         Caption         =   ">"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4320
         TabIndex        =   1
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdVC 
         Caption         =   ">"
         Height          =   255
         Left            =   4320
         TabIndex        =   3
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdVP 
         Caption         =   ">"
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   1080
         Width           =   255
      End
      Begin VB.Frame Frame3 
         Height          =   4215
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   4455
         Begin VB.CommandButton cmdListo 
            Caption         =   "ENVIAR"
            Height          =   375
            Left            =   3000
            TabIndex        =   11
            Top             =   1560
            Width           =   1215
         End
         Begin VB.ComboBox cboUnidad 
            Height          =   315
            Left            =   2280
            TabIndex        =   7
            Top             =   480
            Width           =   1815
         End
         Begin VB.ComboBox cboCant 
            Height          =   315
            Left            =   360
            TabIndex        =   6
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtDescripcion 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   360
            TabIndex        =   8
            Top             =   1080
            Width           =   3735
         End
         Begin VB.CommandButton cmdBorrar 
            Caption         =   "BORRAR"
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "AGREGAR"
            Height          =   375
            Left            =   1560
            TabIndex        =   9
            Top             =   1560
            Width           =   1335
         End
         Begin MSComctlLib.ListView lvwRequi 
            Height          =   1935
            Left            =   120
            TabIndex        =   24
            Top             =   2040
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   3413
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CANTIDAD"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "UNIDAD"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "DESCRIPCION"
               Object.Width           =   5115
            EndProperty
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "UNIDAD"
            Height          =   255
            Left            =   2280
            TabIndex        =   23
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "CANTIDAD"
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "DESCRIPCION"
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   840
            Width           =   3735
         End
      End
      Begin VB.TextBox txtProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtEjec 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.OptionButton opnP 
      Caption         =   "USO PERSONAL"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton opnV 
      Caption         =   "VENTA"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   1695
      Begin VB.Label lblCom 
         Alignment       =   2  'Center
         Caption         =   "UTILICE ESTA OPCION PARA PEDIR UN ARTICULO DE USO PERSONAL Y NO PARA VENTAS"
         Height          =   1335
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmRequi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim H As Integer
Dim w As Integer
Dim CONT As Integer
Dim ITMX As ListItem
Dim EXPANDIDO As Boolean
Dim OPCIONES As Byte
Dim CANT As Integer
Dim UNID As String
Dim LvwNoRe As Integer
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1

Private Sub cboCant_DropDown()

    Me.cboCant.Clear
    Me.cboCant.AddItem "1", 0
    Me.cboCant.AddItem "2", 1
    Me.cboCant.AddItem "3", 2
    Me.cboCant.AddItem "4", 3
    Me.cboCant.AddItem "5", 4
    Me.cboCant.AddItem "6", 5
    Me.cboCant.AddItem "7", 6
    Me.cboCant.AddItem "8", 7
    Me.cboCant.AddItem "9", 8
    Me.cboCant.AddItem "10", 9
    Me.cboCant.AddItem "11", 10
    Me.cboCant.AddItem "12", 11

End Sub

Private Sub cboCant_GotFocus()

        Me.cboCant.SelStart = 0
        Me.cboCant.SelLength = Len(Me.cboCant.Text)

End Sub

Private Sub cboCant_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    End If

End Sub

Private Sub cboUnidad_DropDown()

    Me.cboUnidad.Clear
    Me.cboUnidad.AddItem "PIEZA", 0
    Me.cboUnidad.AddItem "LITRO", 1
    Me.cboUnidad.AddItem "KILO", 2
    Me.cboUnidad.AddItem "ONZA", 3
    Me.cboUnidad.AddItem "LIBRA", 4

End Sub

Private Sub cboUnidad_GotFocus()

    Me.cboUnidad.SelStart = 0
    Me.cboUnidad.SelLength = Len(Me.cboUnidad.Text)

End Sub

Private Sub cboUnidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    End If

End Sub

Private Sub cmdAgregar_Click()

    If Trim(Me.txtDescripcion.Text = "") Then
        MsgBox "LA DESCRIPCIÓN ES NECESARIA", vbInformation, "MENSAJE DEL SISTEMA"
        Me.txtDescripcion.SetFocus
    Else
        If Val(Me.cboCant.Text) = 0 Then
            CANT = 1
        Else
            CANT = Trim(Val(Me.cboCant.Text))
        End If
        If Trim(Me.cboUnidad.Text) = "" Then
            UNID = "Pza"
        Else
            UNID = Trim(Me.cboUnidad.Text)
        End If
        
        Set ITMX = Me.lvwRequi.ListItems.Add(, , CANT)
        ITMX.SubItems(1) = UNID
        ITMX.SubItems(2) = Trim(Me.txtDescripcion.Text)
        Me.cboCant.SetFocus
        Me.txtDescripcion.Text = ""
    End If
    
End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Me.cmdAgregar.Value = True
    End If
    
End Sub

Private Sub cmdBorrar_Click()
    
    If Me.lvwRequi.ListItems.Count <> 0 Then
        Me.lvwRequi.ListItems.Remove (Me.lvwRequi.SelectedItem.Index)
    End If
    
End Sub

Private Sub cmdListo_Click()
    
    If Me.lvwRequi.ListItems.Count <> 0 Then
        Set cnn = New ADODB.Connection
        Set rst = New ADODB.Recordset
        Dim sBuscar As String
        Const sPathBase As String = "NEWSERVER"
        With cnn
            .ConnectionString = _
                "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;" & _
                "Data Source=" & sPathBase & ";"
            .Open
        End With
        
        Dim fecha As Date
        fecha = Date
        
        sBuscar = "INSERT INTO REQUISICION (FECHA, PROVEEDOR, CLIENTE, EJECUTIVO) VALUES ('" & fecha & "', '" & txtProveedor.Text & "', '" & txtCliente.Text & "', '" & txtEjec.Text & "')"
        cnn.Execute (sBuscar)
        
        Dim tRs As Recordset
        sBuscar = "SELECT ID_REQUISICION FROM REQUISICION ORDER BY ID_REQUISICION DESC"
        Set tRs = cnn.Execute(sBuscar)
        tRs.MoveFirst
        Dim IDRequi As String
        IDRequi = tRs.Fields("ID_REQUISICION") & ""
        Dim NumeroRegistros As Integer
        NumeroRegistros = lvwRequi.ListItems.Count
        Dim conta As Integer
        For conta = 1 To NumeroRegistros
            sBuscar = "INSERT INTO REQUISICION_PRODUCTO (DESCRIPCION, CANTIDAD, ID_REQUISICION) VALUES ('" & lvwRequi.ListItems(conta).SubItems(2) & "', '" & lvwRequi.ListItems(conta) & "', " & Val(IDRequi) & ")"
            cnn.Execute (sBuscar)
        Next conta
        
        Unload Me
    Else
        MsgBox "AGREGE ELEMENTOS", vbInformation, "MENSAJE DEL SISTEMA"
    End If
    
End Sub

Private Sub cmdOk_Click()
    
    CONTROLES_INVISIBLES
    TAMAÑO
    Me.Frame2.Visible = True
    If Me.opnV.Value = True Then Me.Caption = "REQUISICIONES PARA VENTA"
    If Me.opnP.Value = True Then Me.Caption = "REQUISICIONES PERSONALES"
    
End Sub

Private Sub cmdVC_Click()
    
    'EXPANDIR
    LLENAR_LISTA_CLIENTES
    Me.opnNom.Visible = True
    Me.opnNomCom.Visible = True
    OPCIONES = 2
    Me.lvwL.SetFocus
    
End Sub

Private Sub cmdVE_Click()
    
    'EXPANDIR
    'LLENAR_LISTA_USUARIOS
    'Me.opnNom.Visible = False
    'Me.opnNomCom.Visible = False
    'OPCIONES = 1

End Sub

Private Sub cmdVP_Click()
    
    EXPANDIR
    LLENAR_LISTA_PROVEEDORES
    Me.opnNom.Visible = False
    Me.opnNomCom.Visible = False
    OPCIONES = 4
    Me.lvwL.SetFocus

End Sub

Private Sub Form_Load()

    Me.Height = 3360
    Me.Width = 2025
    Me.txtNE.Text = MENU.Text1(0).Text
    TRAER_NOMBRE_EJECUTIVO

End Sub

Private Sub Form_Resize()

    Dim TopCorner As Integer
    Dim LeftCorner As Integer
    If Me.WindowState <> 0 Then Exit Sub
    TopCorner = (Screen.Height - Me.Height) \ 2
    LeftCorner = (Screen.Width - Me.Width) \ 2
    Me.Move LeftCorner, TopCorner
    
End Sub

Private Sub lvwL_Click()
    
    'CONTRAER
    LLENAR
    
End Sub

Private Sub lvwL_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        LLENAR
    End If
    
End Sub

Private Sub opnNom_Click()
    
    Me.lvwL.ColumnHeaders.Item(2).Width = 4000
    Me.lvwL.ColumnHeaders.Item(3).Width = 0
    OPCIONES = 2

End Sub

Private Sub opnNomCom_Click()

    Me.lvwL.ColumnHeaders.Item(2).Width = 0
    Me.lvwL.ColumnHeaders.Item(3).Width = 4000
    OPCIONES = 3
    
End Sub

Private Sub opnP_Click()

    Me.lblCom.Caption = "UTILICE ESTA OPCION PARA PEDIR UN ARTICULO DE USO PERSONAL Y NO PARA VENTAS"
    
End Sub

Private Sub opnV_Click()

    Me.lblCom.Caption = "UTILICE ESTA OPCION PARA PEDIR UN ARTICULO DE VENTAS QUE NO ESTE EN LA LISTA DE ARICULOS"

End Sub

Private Function TAMAÑO()

    'H = Me.Height
    'w = Me.Width
    
    'For cont = 1 To 3000
        'H = H + 1
        'w = w + 1
        Me.Height = 6435
        Me.Width = 10065
    'Next cont
    
End Function

Private Function CONTROLES_INVISIBLES()

    Me.cmdOk.Visible = False
    Me.opnP.Visible = False
    Me.opnV.Visible = False
    Me.lblCom.Visible = False
    Me.Frame1.Visible = False
    
End Function

Sub LLENAR_LISTA_USUARIOS()

    Me.lvwL.ListItems.Clear
    deAPTONER.LISTA_USUARIOS
    With deAPTONER.rsLISTA_USUARIOS
        While Not .EOF
            Set ITMX = Me.lvwL.ListItems.Add(, , Trim(!ID_USUARIO))
            If Not IsNull(!NOM) Then ITMX.SubItems(1) = Trim(!NOM)
            .MoveNext
        Wend
        .Close
    End With
    
End Sub

Sub LLENAR_LISTA_CLIENTES()

    Me.lvwL.ListItems.Clear
    deAPTONER.LISTA_CLIENTES Trim(Me.txtCliente.Text)
    With deAPTONER.rsLISTA_CLIENTES
        While Not .EOF
            Set ITMX = Me.lvwL.ListItems.Add(, , Trim(!ID_CLIENTE))
            If Not IsNull(!NOMBRE) Then ITMX.SubItems(1) = Trim(!NOMBRE)
            If Not IsNull(!NOMBRE_COMERCIAL) Then ITMX.SubItems(2) = Trim(!NOMBRE_COMERCIAL)
            .MoveNext
        Wend
        .Close
    End With
    
End Sub

Sub LLENAR_LISTA_PROVEEDORES()

    Me.lvwL.ListItems.Clear
    deAPTONER.LISTA_PROVEEDOR Trim(Me.txtProveedor.Text)
    With deAPTONER.rsLISTA_PROVEEDOR
        While Not .EOF
            Set ITMX = Me.lvwL.ListItems.Add(, , Trim(!ID_PROVEEDOR))
            If Not IsNull(!NOMBRE) Then ITMX.SubItems(1) = Trim(!NOMBRE)
            .MoveNext
        Wend
        .Close
    End With
    
End Sub

Sub EXPANDIR()
    
    'If EXPANDIDO = False Then
        
        'w = Me.Width
        
        'For cont = 1 To 5025
            'w = w + 1
            'Me.Width = w
        'Next cont

        'EXPANDIDO = True
        
    'End If
    
End Sub

Sub CONTRAER()
    
    If EXPANDIDO = True Then
        
        'w = Me.Width
        
        'For cont = 1 To 25
            'w = w - 200
            'Me.Width = w
        'Next cont

        'EXPANDIDO = False
        
    End If
    
End Sub

Sub LLENAR()
On Error GoTo MANEJAERROR:
    Select Case OPCIONES
    Case 1:
            'If Me.lvwL.SelectedItem.Selected = True Then
                'Me.txtEjec.Text = Trim(Me.lvwL.SelectedItem.SubItems(1))
                'Me.txtNE.Text = Trim(Me.lvwL.SelectedItem)
            'End If
            MsgBox "Error"
            
    Case 2:
            If Me.lvwL.SelectedItem.Selected = True Then
                Me.txtCliente.Text = Trim(Me.lvwL.SelectedItem.SubItems(1))
                Me.txtNC.Text = Trim(Me.lvwL.SelectedItem)
            End If
            
    Case 3:
            If Me.lvwL.SelectedItem.Selected = True Then
                Me.txtCliente.Text = Trim(Me.lvwL.SelectedItem.SubItems(2))
                Me.txtNC.Text = Trim(Me.lvwL.SelectedItem)
            End If
            
    Case 4:
            If Me.lvwL.SelectedItem.Selected = True Then
                Me.txtProveedor.Text = Trim(Me.lvwL.SelectedItem.SubItems(1))
                Me.txtNP.Text = Trim(Me.lvwL.SelectedItem)
            End If
            
    End Select
    
MANEJAERROR:
    Err.Clear
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Me.cmdVC.Value = True
    End If
    
End Sub

Private Sub txtDescripcion_GotFocus()

        Me.txtDescripcion.SelStart = 0
        Me.txtDescripcion.SelLength = Len(Me.txtDescripcion.Text)

End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    Else
        Me.txtDescripcion.Visible = True
    End If

End Sub

Sub TRAER_NOMBRE_EJECUTIVO()

    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Const sPathBase As String = "NEWSERVER"
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With

    Dim intSQL As String
    intSQL = "SELECT NOMBRE FROM USUARIOS WHERE ID_USUARIO = " & "'" & Trim(Me.txtNE.Text) & "'"
    Dim tRs As Recordset
    Set tRs = cnn.Execute(intSQL)
    tRs.MoveFirst
    Me.txtEjec.Text = tRs.Fields("NOMBRE")

End Sub

Private Sub txtProveedor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Me.cmdVP.Value = True
    End If
    
End Sub
