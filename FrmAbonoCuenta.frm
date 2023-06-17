VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmAbonoCuenta 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargar Abono a Cuenta"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9720
      TabIndex        =   24
      Top             =   5040
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
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmAbonoCuenta.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmAbonoCuenta.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmAbonoCuenta.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   360
         Width           =   6855
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
         Left            =   7920
         Picture         =   "FrmAbonoCuenta.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7800
         TabIndex        =   20
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3120
         Width           =   5295
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cantidad de Abono"
         Height          =   1095
         Left            =   5520
         TabIndex        =   18
         Top             =   3600
         Width           =   3855
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   1935
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Registrar"
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
            Left            =   2520
            Picture         =   "FrmAbonoCuenta.frx":4DDA
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ingrese Clave"
         Height          =   1095
         Left            =   5520
         TabIndex        =   17
         Top             =   4800
         Width           =   3855
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   1935
         End
         Begin VB.CommandButton Command4 
            Caption         =   "A Factura"
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
            Left            =   2520
            Picture         =   "FrmAbonoCuenta.frx":77AC
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos del Cheque"
         Height          =   2295
         Left            =   120
         TabIndex        =   13
         Top             =   3600
         Width           =   5295
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1800
            MaxLength       =   18
            TabIndex        =   3
            Top             =   480
            Width           =   3375
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1800
            TabIndex        =   4
            Top             =   960
            Width           =   2055
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Pago con Efectivo"
            Height          =   255
            Left            =   1800
            TabIndex        =   6
            Top             =   1920
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1800
            TabIndex        =   5
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   56557569
            CurrentDate     =   38833
         End
         Begin VB.Label Label4 
            Caption         =   "Numero de Cheque :"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Banco :"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha para Depositar :"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1560
            Width           =   1815
         End
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   6000
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "FrmAbonoCuenta.frx":A17E
         Top             =   3120
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4048
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
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad Pendiente :"
         Height          =   255
         Left            =   6240
         TabIndex        =   22
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3120
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmAbonoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim ClvCliente As Integer
Dim VarLimCred As Double
Dim ClvUsuario As String
Private Sub Check1_Click()
On Error GoTo ManejaError
    If Check1.Value = 1 Then
        DTPicker1.Enabled = False
        Combo1.Enabled = False
        Text6.Enabled = False
    Else
        DTPicker1.Enabled = True
        Combo1.Enabled = True
        Text6.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        DTPicker1.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HFFE1E1
End Sub
Private Sub Combo1_LostFocus()
    Combo1.BackColor = &H80000005
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
On Error GoTo ManejaError
    If Text2.Text = "" Or Text3.Text = "" Or Text3.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Then
        MsgBox "FALTA INFORMACION NECESARIA PARA EL REGISTRO!", vbInformation, "SACC"
    Else
        Dim sBuscar As String
        Dim tRs As Recordset
        Dim tRs1 As Recordset
        Text7.Text = ""
        Text4.Text = Replace(Text4.Text, ",", ".")
        Text2.Text = Replace(Text2.Text, ",", ".")
        If Check1.Value = 1 Then
            sBuscar = "INSERT INTO ABONOS_CUENTA (ID_CLIENTE, CANT_ABONO, DEUDA, FECHA, ID_USUARIO, EFECTIVO) VALUES (" & ClvCliente & ", " & Text4.Text & ", " & Text2.Text & ", '" & Format(Date, "dd/mm/yyyy") & "', '" & VarMen.Text1(0).Text & "', 'S');"
        Else
            sBuscar = "INSERT INTO ABONOS_CUENTA (ID_CLIENTE, CANT_ABONO, DEUDA, FECHA, ID_USUARIO, EFECTIVO, NO_CHEQUE, BANCO, FECHA_CHEQUE) VALUES (" & ClvCliente & ", " & Text4.Text & ", " & Text2.Text & ", '" & Format(Date, "dd/mm/yyyy") & "', '" & VarMen.Text1(0).Text & "', 'N', '" & Text6.Text & "', '" & Combo1.Text & "', '" & DTPicker1.Value & "');"
        End If
        cnn.Execute (sBuscar)
        sBuscar = "SELECT ID_CUENTA, TOTAL_COMPRA FROM CUENTAS WHERE ID_CLIENTE = " & ClvCliente & " AND PAGADA = 'N' ORDER BY ID_CUENTA"
        Set tRs = cnn.Execute(sBuscar)
        'Si tiene saldo a favor se le suma a su abono
        If CDbl(Text2.Text) < 0 Then
            Text4.Text = Abs(CDbl(Text2.Text)) + CDbl(Text4.Text)
        End If
        If (tRs.BOF And tRs.EOF) Then
            tRs.MoveFirst
            Do While Not tRs.EOF
                ' Buscar cuantas facturas se pagan con el abono dado
                If tRs.Fields("TOTAL_COMPRA") <= CDbl(Text4.Text) Then
                    sBuscar = "UPDATE CUENTAS SET PAGADA = 'S' WHERE ID_CUENTA = " & tRs.Fields("ID_CUENTA")
                    Set tRs1 = cnn.Execute(sBuscar)
                    sBuscar = "SELECT ID_VENTA FROM CUENTA_VENTA WHERE ID_CUENTE = " & tRs.Fields("ID_CUENTA")
                    Set tRs1 = cnn.Execute(sBuscar)
                    Text7.Text = Text7.Text & tRs1.Fields("ID_VENTA") & " " & vbCrLf
                Else
                    tRs.MoveLast
                End If
                tRs.MoveNext
            Loop
        End If
        If Text7.Text = "" Then
            MsgBox "NO SE LIBERO NINGUNA FACTURA!", vbInformation, "SACC"
        Else
            MsgBox "LAS VENTAS LIBERADAS SON : " & Text7.Text, vbInformation, "SACC"
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
    Err.Clear
End Sub
Private Sub Command3_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        'Me.Command1.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Command4_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        'Me.Command1 = 13
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Text4.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    Combo1.AddItem "BANAMEX"
    Combo1.AddItem "BANCOMER"
    Combo1.AddItem "BANORTE"
    Combo1.AddItem "HSBC"
    Combo1.AddItem "SANTANDER"
    Combo1.AddItem "SCOTIABANK"
    Combo1.AddItem "Otros"
    
    ClvUsuario = VarMen.Text1(0).Text
    Command3.Enabled = False
    Me.Command4.Enabled = False
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .CheckBoxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Cliente", 1800
        .ColumnHeaders.Add , , "Nombre", 7450
        .ColumnHeaders.Add , , "RFC", 2450
        .ColumnHeaders.Add , , "Limite de credito", 2450
        .ColumnHeaders.Add , , "Descuento", 1550
    End With
    Me.Command3.Enabled = False
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
    Err.Clear
End Sub
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_CLIENTE, NOMBRE, RFC, LIMITE_CREDITO, DESCUENTO FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND LIMITE_CREDITO <> 0"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            Text1.Text = ""
            MsgBox "No se encontro cliente con credito registrado a ese nombre"
        Else
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                tLi.SubItems(1) = .Fields("NOMBRE") & ""
                tLi.SubItems(2) = .Fields("RFC") & ""
                tLi.SubItems(3) = .Fields("LIMITE_CREDITO") & ""
                If .Fields("DESCUENTO") = "" Then
                    tLi.SubItems(4) = "0.00"
                Else
                    tLi.SubItems(4) = .Fields("DESCUENTO") & ""
                End If
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
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
    ClvCliente = Item
    Text3.Text = Item.SubItems(1)
    VarLimCred = Item.SubItems(3)
    Dim Acum As Double
    Dim sBuscar As String
    Dim tRs As Recordset
    sBuscar = "SELECT TOTAL_COMPRA FROM CUENTAS WHERE ID_CLIENTE = " & Item
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                If .Fields("TOTAL_COMPRA") <> Null Then
                    Acum = Acum + CDbl(.Fields("TOTAL_COMPRA"))
                End If
                .MoveNext
            Loop
        End If
    End With
    sBuscar = "SELECT CANT_ABONO FROM ABONOS_CUENTA WHERE ID_CLIENTE = " & Item
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                If .Fields("TOTAL_COMPRA") <> Null Then
                    Acum = Acum - CDbl(.Fields("CANT_ABONO"))
                End If
                .MoveNext
            Loop
        End If
    End With
    Text2.Text = Acum
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
    Err.Clear
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Text6.SetFocus
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
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
Private Sub Text4_Change()
On Error GoTo ManejaError
    If Text4.Text <> "" Then
        Command3.Enabled = True
    Else
        Command3.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        If Text4.Text <> "" Then
            Me.Command3.SetFocus
        Else
            Text5.SetFocus
        End If
    End If
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
Private Sub Text4_GotFocus()
    Text4.BackColor = &HFFE1E1
End Sub
Private Sub Text4_LostFocus()
    Text4.BackColor = &H80000005
End Sub
Private Sub Text5_Change()
On Error GoTo ManejaError
    If Text5.Text = "" Then
        Me.Command4.Enabled = False
    Else
        Me.Command4.Enabled = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.Command4.SetFocus
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text5_GotFocus()
    Text5.BackColor = &HFFE1E1
End Sub
Private Sub Text5_LostFocus()
    Text5.BackColor = &H80000005
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Combo1.SetFocus
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text6_GotFocus()
    Text6.BackColor = &HFFE1E1
End Sub
Private Sub Text6_LostFocus()
    Text6.BackColor = &H80000005
End Sub
