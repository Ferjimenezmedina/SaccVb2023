VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmCapCoti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capturar Cotizacion"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10320
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   9120
      ScaleHeight     =   6315
      ScaleWidth      =   1515
      TabIndex        =   19
      Top             =   0
      Width           =   1575
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   20
         Top             =   4920
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
            TabIndex        =   21
            Top             =   960
            Width           =   975
         End
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmCapCoti.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "FrmCapCoti.frx":030A
            Top             =   240
            Width           =   705
         End
      End
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
      Left            =   7440
      Picture         =   "FrmCapCoti.frx":1DBC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Left            =   5640
      Picture         =   "FrmCapCoti.frx":478E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Top             =   3960
      Width           =   4215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3480
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3413
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
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4320
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3000
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3000
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2295
      Left            =   4680
      TabIndex        =   2
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
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
   Begin VB.Label Label8 
      Caption         =   "Clave del Proveedor"
      Height          =   255
      Left            =   6000
      TabIndex        =   18
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Precio/U"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Clave"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   4560
      X2              =   4560
      Y1              =   120
      Y2              =   2880
   End
   Begin VB.Label Label2 
      Caption         =   "Articulos :"
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Pendientes :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "FrmCapCoti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim NoCoti As String
Private Sub Command1_Click()
On Error GoTo ManejaError
    Buscar
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cmdCancelar_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    Dim sqlComanda As String
    Dim Precio As String
    Precio = Format(CDbl(Text4.Text) * 1.3, "0.00")
    sqlComanda = "INSERT INTO COTIZACION_PROV (ID_COTIZACION, ID_PRODUCTO, DESCRIPCION, ID_PROVEEDOR, PRECIO) VALUES (" & NoCoti & ", '" & Text1.Text & "', '" & Text2.Text & "', " & Text5.Text & ", " & Precio & ");"
    cnn.Execute (sqlComanda)
    Me.Command3.Enabled = False
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
Private Sub Command3_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        ListView2.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Form_Load()
On Error GoTo ManejaError
    Me.Command1.Enabled = False
    Me.Command3.Enabled = False
    Const sPathBase As String = "LINUX"
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "NO. COTIZACION", 0
        .ColumnHeaders.Add , , "NOMBRE", 3700
        .ColumnHeaders.Add , , "DIRECCION", 1500
        .ColumnHeaders.Add , , "COLONIA", 1500
        .ColumnHeaders.Add , , "CIUDAD", 1500
        .ColumnHeaders.Add , , "TELEFONO", 1000
        .ColumnHeaders.Add , , "COMENTARIOS", 1000
    End With
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "CLAVE", 2000
        .ColumnHeaders.Add , , "DESCRIPCION", 3700
        .ColumnHeaders.Add , , "CANTIDAD", 1500
    End With
    With ListView3
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "NO. PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE", 3700
        .ColumnHeaders.Add , , "DIRECCION", 1500
        .ColumnHeaders.Add , , "COLONIA", 1500
        .ColumnHeaders.Add , , "CIUDAD", 1500
        .ColumnHeaders.Add , , "TELEFONO", 1000
    End With
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_COTIZACION, NOMBRE, DIRECCION, COLONIA, CIUDAD, TELEFONO, COMENTARIOS FROM COTIZACION WHERE PENDIENTE = 'S'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        ListView1.ListItems.Clear
        tRs.MoveFirst
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_COTIZACION") & "")
            tLi.SubItems(1) = tRs.Fields("NOMBRE") & ""
            tLi.SubItems(2) = tRs.Fields("DIRECCION") & ""
            tLi.SubItems(3) = tRs.Fields("COLONIA") & ""
            tLi.SubItems(4) = tRs.Fields("CIUDAD") & ""
            tLi.SubItems(5) = tRs.Fields("TELEFONO") & ""
            tLi.SubItems(6) = tRs.Fields("COMENTARIOS") & ""
            tRs.MoveNext
        Loop
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
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    NoCoti = Item
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM COTIZACION_DETALLE WHERE ID_COTIZACION = " & Item
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        ListView2.ListItems.Clear
        tRs.MoveFirst
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO") & "")
            tLi.SubItems(1) = tRs.Fields("DESCRIPCION") & ""
            tLi.SubItems(2) = tRs.Fields("CANTIDAD") & ""
            tRs.MoveNext
        Loop
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
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_PROVEEDOR, NOMBRE, DIRECCION, COLONIA, CIUDAD, TELEFONO1 FROM PROVEEDOR WHERE NOMBRE LIKE '%" & Text6.Text & "%'"
    Set tRs = cnn.Execute(sBuscar)
    ListView3.ListItems.Clear
    If Not (tRs.BOF And tRs.EOF) Then
        tRs.MoveFirst
        Do While Not tRs.EOF
            Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR") & "")
            tLi.SubItems(1) = tRs.Fields("NOMBRE") & ""
            tLi.SubItems(2) = tRs.Fields("DIRECCION") & ""
            tLi.SubItems(3) = tRs.Fields("COLONIA") & ""
            tLi.SubItems(4) = tRs.Fields("CIUDAD") & ""
            tLi.SubItems(5) = tRs.Fields("TELEFONO1") & ""
            tRs.MoveNext
        Loop
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
Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        ListView2.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Text1.Text = Item
    Text2.Text = Item.SubItems(1)
    Text3.Text = Item.SubItems(2)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub ListView2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Text4.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Text5.Text = Item
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub ListView3_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.Command3.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text1_Change()
On Error GoTo ManejaError
    If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
        Me.Command3.Enabled = True
    Else
        Me.Command3.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
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
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text2_Change()
On Error GoTo ManejaError
    If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
        Me.Command3.Enabled = True
    Else
        Me.Command3.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
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
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text3_Change()
On Error GoTo ManejaError
    If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
        Me.Command3.Enabled = True
    Else
        Me.Command3.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text3_GotFocus()
On Error GoTo ManejaError
    Text3.SetFocus
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
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
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text4_Change()
On Error GoTo ManejaError
    If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
        Me.Command3.Enabled = True
    Else
        Me.Command3.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text4_GotFocus()
On Error GoTo ManejaError
    Text4.BackColor = &HFFE1E1
    Text4.SetFocus
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text4_LostFocus()
      Text4.BackColor = &H80000005
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Text6.SetFocus
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
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text5_Change()
On Error GoTo ManejaError
    If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
        Me.Command3.Enabled = True
    Else
        Me.Command3.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text5_GotFocus()
On Error GoTo ManejaError
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text6_Change()
On Error GoTo ManejaError
    If Text6.Text <> "" Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text6_GotFocus()
    Text6.BackColor = &HFFE1E1
End Sub
Private Sub Text6_LostFocus()
      Text6.BackColor = &H80000005
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And Text6.Text <> "" Then
        Buscar
        ListView3.SetFocus
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
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub

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


