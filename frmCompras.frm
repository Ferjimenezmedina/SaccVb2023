VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCompras 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compras"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCancel 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9480
      TabIndex        =   1
      Top             =   5760
      Width           =   975
      Begin VB.Label Label12 
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Image cmdCancelar 
         Height          =   705
         Left            =   120
         MouseIcon       =   "frmCompras.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmCompras.frx":030A
         Top             =   240
         Width           =   705
      End
   End
   Begin TabDlg.SSTab tabCompras 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   11668
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Requisiciones"
      TabPicture(0)   =   "frmCompras.frx":1DBC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAgregado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lvwRequisiciones"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lvwProveedores"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtId_Proveedor"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdImprimir"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAgregar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Cotizar"
      TabPicture(1)   =   "frmCompras.frx":1DD8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Finalizar"
      TabPicture(2)   =   "frmCompras.frx":1DF4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton cmdAgregar 
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
         Left            =   7560
         Picture         =   "frmCompras.frx":1E10
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton cmdImprimir 
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
         Height          =   375
         Left            =   6120
         Picture         =   "frmCompras.frx":47E2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtId_Proveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   6120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   5280
         TabIndex        =   11
         Top             =   480
         Width           =   3495
         Begin VB.TextBox txtNotas 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   915
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   600
            Width           =   3255
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "NOTAS"
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
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   4935
         Begin VB.TextBox txtDireccion 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   915
            Left            =   1920
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txtTel2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtTel1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtTel3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "DIRECCIÓN"
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
            TabIndex        =   10
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "TELEFONOS"
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
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1695
         End
      End
      Begin MSComctlLib.ListView lvwProveedores 
         Height          =   1335
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwRequisiciones 
         Height          =   1695
         Left            =   240
         TabIndex        =   17
         Top             =   4200
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2990
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label5 
         Caption         =   "Proveedor"
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
         TabIndex        =   20
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Proveedor"
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
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lblAgregado 
         Alignment       =   2  'Center
         Caption         =   "----------------------------------------------"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   6120
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As Recordset
Private Sub Form_Resize()

    On Error GoTo ManejaError
    
    Dim Centro As Single
    Dim Alto_Tab As Single
    
    ' Para centrar el Frame
    Centro = (Me.ScaleWidth - Me.fraCancel.Width) - 200
    
    ' Lo Posiciona: el -50 es para dejar un borde
    Me.fraCancel.Move Centro, Me.ScaleHeight - Me.fraCancel.Height - 200
    
    ' El alto del Tab es el alto del formulario _
     menos la posición Top del Frame y menos 50 _
     para dejar un espacio entre el Tab y el Fame
    
    Alto_Tab = Me.fraCancel.Top - 50
    
    ' Esto chequea que el valor Height del text no sea negativo _
      ya que si no da error
    If Alto_Tab <= 0 Then Alto_Tab = 100
    
    'Posiciona y redimensiona el Tab
    Me.tabCompras.Move 200, 200, (Me.ScaleWidth - 400), (Alto_Tab - 200)

Exit Sub
ManejaError:
    MsgBox Err.Number & " " & Err.Description
    Err.Clear
    
End Sub

Private Sub cmdAgregar_Click()
On Error GoTo ManejaError
    If Puede_Agregar Then
        Dim ID_REQUISICION As Integer
        Dim ID_PROVEEDOR As Integer
        Dim ID_PRODUCTO As String
        Dim DESCRIPCION As String
        Dim CANTIDAD As Double
        Dim DIAS_ENTREGA As Integer
        Dim Precio As Double
        
        ID_REQUISICION = Me.lvwRequisiciones.SelectedItem
        ID_PROVEEDOR = Me.txtId_Proveedor.Text
        ID_PRODUCTO = Me.lvwRequisiciones.SelectedItem.ListSubItems(1)
        DESCRIPCION = Me.lvwRequisiciones.SelectedItem.ListSubItems(2)
        CANTIDAD = Me.lvwRequisiciones.SelectedItem.ListSubItems(3)
        
        sqlQuery = "INSERT INTO COTIZA_REQUI (ID_REQUISICION, ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, CANTIDAD, FECHA) VALUES (" & ID_REQUISICION & ", " & ID_PROVEEDOR & ", '" & ID_PRODUCTO & "', '" & DESCRIPCION & "', " & CANTIDAD & ", '" & Format(Date, "dd/mm/yyyy") & "')"
        Set tRs = cnn.Execute(sqlQuery)
        
        sqlQuery = "UPDATE REQUISICION SET CONTADOR = CONTADOR + 1 WHERE ID_REQUISICION = " & ID_REQUISICION
        Set tRs = cnn.Execute(sqlQuery)
        
        Me.lblAgregado.Caption = Me.lvwRequisiciones.SelectedItem.ListSubItems(1) & " AGREGADO"
        Llenar_Lista_Requisiciones
    End If
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub

Private Sub cmdCancelar_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()
On Error GoTo ManejaError
    With Me.lvwRequisiciones
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "ID REQUISICION", 0
        .ColumnHeaders.Add , , "CLAVE DEL PRODUCTO", 2200
        .ColumnHeaders.Add , , "DESCRIPCION", 3500
        .ColumnHeaders.Add , , "CANTIDAD", 1000, 2
        .ColumnHeaders.Add , , "FECHA", 2000, 2
        .ColumnHeaders.Add , , "CONTADOR", 1300, 2
    End With
    
    With Me.lvwProveedores
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "ID PROVEEDOR", 0
        .ColumnHeaders.Add , , "PROVEEDOR", 4500, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
    End With
    
    If Hay_Proveedores Then
        Llenar_Lista_Proveedores
    Else
        MsgBox "NO HAY PROVEEDORES", vbInformation, "MENSAJE DEL SISTEMA"
    End If
    
    If Hay_Requisiciones Then
        Llenar_Lista_Requisiciones
    End If
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Dim sPathBase As String
    sPathBase = "LINUX"
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub

Sub Llenar_Lista_Requisiciones()
On Error GoTo ManejaError
    sqlQuery = "SELECT ID_REQUISICION, ID_PRODUCTO, DESCRIPCION, CANTIDAD, FECHA, CONTADOR FROM REQUISICION WHERE ACTIVO = 0"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.BOF And .EOF) Then
            Me.lvwRequisiciones.ListItems.Clear
            Do While Not .EOF
                Set tLi = lvwRequisiciones.ListItems.Add(, , .Fields("ID_REQUISICION"))
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("DESCRIPCION")) Then tLi.SubItems(2) = Trim(.Fields("DESCRIPCION"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(3) = Trim(.Fields("CANTIDAD"))
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(4) = Trim(.Fields("FECHA"))
                If Not IsNull(.Fields("CONTADOR")) Then
                    tLi.SubItems(5) = Trim(.Fields("CONTADOR"))
                Else
                    tLi.SubItems(5) = "0"
                End If
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub

Function Hay_Requisiciones() As Boolean
On Error GoTo ManejaError
    sqlQuery = "SELECT COUNT(ID_REQUISICION)ID_REQUISICION FROM REQUISICION WHERE ACTIVO = 0"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_REQUISICION") <> 0 Then
            Hay_Requisiciones = True
        Else
            Hay_Requisiciones = False
        End If
    End With
Exit Function
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Function

Sub Llenar_Lista_Proveedores()
On Error GoTo ManejaError
    sqlQuery = "SELECT ID_PROVEEDOR, NOMBRE, DIRECCION, COLONIA, CIUDAD, CP, RFC, TELEFONO1, TELEFONO2, TELEFONO3, NOTAS, ESTADO, PAIS FROM PROVEEDOR WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.BOF And .EOF) Then
            Me.lvwProveedores.ListItems.Clear
            Do While Not .EOF
                Set tLi = lvwProveedores.ListItems.Add(, , .Fields("ID_PROVEEDOR"))
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = Trim(.Fields("NOMBRE"))
                If Not IsNull(.Fields("DIRECCION")) Then tLi.SubItems(2) = Trim(.Fields("DIRECCION"))
                If Not IsNull(.Fields("COLONIA")) Then tLi.SubItems(3) = Trim(.Fields("COLONIA"))
                If Not IsNull(.Fields("CP")) Then tLi.SubItems(4) = Trim(.Fields("CP"))
                If Not IsNull(.Fields("RFC")) Then tLi.SubItems(5) = Trim(.Fields("RFC"))
                If Not IsNull(.Fields("TELEFONO1")) Then tLi.SubItems(6) = Trim(.Fields("TELEFONO1"))
                If Not IsNull(.Fields("TELEFONO2")) Then tLi.SubItems(7) = Trim(.Fields("TELEFONO2"))
                If Not IsNull(.Fields("TELEFONO3")) Then tLi.SubItems(8) = Trim(.Fields("TELEFONO3"))
                If Not IsNull(.Fields("NOTAS")) Then tLi.SubItems(9) = Trim(.Fields("NOTAS"))
                If Not IsNull(.Fields("ESTADO")) Then tLi.SubItems(10) = Trim(.Fields("ESTADO"))
                If Not IsNull(.Fields("PAIS")) Then tLi.SubItems(11) = Trim(.Fields("PAIS"))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub

Function Hay_Proveedores() As Boolean
On Error GoTo ManejaError
    sqlQuery = "SELECT COUNT(ID_PROVEEDOR)ID_PROVEEDOR FROM PROVEEDOR WHERE ELIMINADO = 'N'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_PROVEEDOR") <> 0 Then
            Hay_Proveedores = True
        Else
            Hay_Proveedores = False
        End If
    End With
Exit Function
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Function

Private Sub lvwProveedores_Click()

    Me.lblAgregado.Caption = "----------------------------------------------"

End Sub

Private Sub lvwProveedores_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Me.txtId_Proveedor.Text = Item
    Label4.Caption = Item.SubItems(1)
    Me.txtDireccion.Text = Item.SubItems(2) + " " + Item.SubItems(3) + " " + Item.SubItems(4) + " " + Item.SubItems(5) + " " + Item.SubItems(10) + " " + Item.SubItems(11)
    Me.txtTel1.Text = Item.SubItems(6)
    Me.txtTel2.Text = Item.SubItems(7)
    Me.txtTel3.Text = Item.SubItems(8)
    Me.txtNotas.Text = Item.SubItems(9)
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Sub

Private Sub lvwRequisiciones_Click()

    Me.lblAgregado.Caption = "----------------------------------------------"
    
End Sub

Function Puede_Agregar() As Boolean
On Error GoTo ManejaError
    If Me.lvwRequisiciones.ListItems.Count = 0 Then
        MsgBox "NO HAY REQUISICIONES", vbInformation, "MENSAJE DEL SISTEMA"
        Puede_Agregar = False
        Exit Function
    End If

    If Me.txtId_Proveedor.Text = "" Then
        MsgBox "SELECCIONE EL PROVEEDOR", vbInformation, "MENSAJE DEL SISTEMA"
        Puede_Agregar = False
        Exit Function
    End If
    
    If Me.lvwRequisiciones.SelectedItem.Selected = False Then
        MsgBox "SELECCIONE LA REQUISICION", vbInformation, "MENSAJE DEL SISTEMA"
        Puede_Agregar = False
        Exit Function
    End If
    
    Puede_Agregar = True
Exit Function
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
    End If
End Function

Sub Reconexion()
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Const sPathBase As String = "LINUX"
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
        MsgBox "LA CONEXION SE RESABLECIO CON EXITO. PUEDE CONTINUAR CON SU TRABAJO.", vbInformation, "MENSAJE DEL SISTEMA"
    End With
Exit Sub
ManejaError:
    MsgBox Err.Number & Err.Description
    Err.Clear
    If MsgBox("NO PUDIMOS RESTABLECER LA CONEXIÓN, ¿DESEA REINTENTARLO?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
End Sub



