VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmNueCoti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nueva Cotizacion"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   7335
      Left            =   10080
      ScaleHeight     =   7275
      ScaleWidth      =   1395
      TabIndex        =   30
      Top             =   0
      Width           =   1455
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   240
         TabIndex        =   31
         Top             =   5880
         Width           =   975
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FrmNueCoti.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "FrmNueCoti.frx":030A
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label11 
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
            TabIndex        =   32
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin VB.TextBox Text10 
      Height          =   1095
      Left            =   6360
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   5520
      Width           =   3495
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   840
      TabIndex        =   27
      Top             =   3000
      Width           =   6015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Terminar"
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
      Left            =   8640
      Picture         =   "FrmNueCoti.frx":1DBC
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
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
      Left            =   8640
      Picture         =   "FrmNueCoti.frx":478E
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Height          =   375
      Left            =   5280
      Picture         =   "FrmNueCoti.frx":7160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6840
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   1695
      Left            =   120
      TabIndex        =   22
      Top             =   5520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2990
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
   Begin VB.CommandButton Command2 
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
      Left            =   8640
      Picture         =   "FrmNueCoti.frx":9B32
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   20
      Top             =   4560
      Width           =   3615
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   6240
      MaxLength       =   100
      TabIndex        =   18
      Top             =   4080
      Width           =   3615
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   6240
      MaxLength       =   20
      TabIndex        =   16
      Top             =   3600
      Width           =   3615
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3201
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
   Begin VB.OptionButton Option2 
      Caption         =   "Por Descripcion"
      Height          =   195
      Left            =   6960
      TabIndex        =   13
      Top             =   3120
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Clave"
      Height          =   195
      Left            =   6960
      TabIndex        =   12
      Top             =   2880
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      MaxLength       =   14
      TabIndex        =   11
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   9
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   5
      Top             =   720
      Width           =   3615
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3201
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
      Left            =   8640
      Picture         =   "FrmNueCoti.frx":C504
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   1
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label Label10 
      Caption         =   "Comentario"
      Height          =   255
      Left            =   5400
      TabIndex        =   28
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Ciudad"
      Height          =   255
      Left            =   5280
      TabIndex        =   26
      Top             =   1680
      Width           =   855
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   9840
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label9 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   5280
      TabIndex        =   19
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Clave"
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   3600
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   5160
      X2              =   5160
      Y1              =   3480
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   5160
      X2              =   5160
      Y1              =   720
      Y2              =   2520
   End
   Begin VB.Label Label6 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Telefono"
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Colonia"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Direccion"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre del cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "FrmNueCoti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim INDI As Integer
Dim IdClien As String

Private Sub cmdCancelar_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub

Private Sub Command1_Click()
On Error GoTo ManejaError
    Buscar
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    If Text8.Text = "" Or Val(Text8.Text) = 0 Then
        MsgBox "                   Debe dar cantidad de articulo                   "
        Text8.SetFocus
    Else
        Dim tLi As ListItem
        Me.Command2.Enabled = False
        Set tLi = ListView3.ListItems.Add(, , Text6.Text & "")
        tLi.SubItems(1) = Text7.Text & ""
        tLi.SubItems(2) = Text8.Text & ""
        Text6.Text = ""
        Text7.Text = ""
        Text8.Text = ""
        Me.Command6.Enabled = True
        Text9.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    ListView3.ListItems.Remove (INDI)
    Me.Command3.Enabled = False
    If ListView3.ListItems.Count = 0 Then
        Me.Command6.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Command5_Click()
On Error GoTo ManejaError
    BusProd
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Command6_Click()
On Error GoTo ManejaError
    Dim NumeroRegistros As Integer
    Dim sqlComanda As String
    Dim NoCot As String
    Dim tRs As Recordset
    sqlComanda = "INSERT INTO COTIZACION (ID_CLIENTE, PENDIENTE, NOMBRE, DIRECCION, COLONIA, CIUDAD, TELEFONO, COMENTARIOS, FECHA) VALUES (" & IdClien & ", 'S', '" & Text1.Text & "', '" & Text2.Text & "', '" & Text3.Text & "', '" & Text4.Text & "', '" & Text5.Text & "', '" & Text10.Text & "', '" & Date & "');"
    cnn.Execute (sqlComanda)
    sqlComanda = "SELECT ID_COTIZACION FROM COTIZACION WHERE FECHA = '" & Date & "' ORDER BY ID_COTIZACION DESC"
    Set tRs = cnn.Execute(sqlComanda)
    NoCot = tRs.Fields("ID_COTIZACION")
    NumeroRegistros = ListView3.ListItems.Count
    Dim Conta As Integer
    For Conta = 1 To NumeroRegistros
        sqlComanda = "INSERT INTO COTIZACION_DETALLE (ID_PRODUCTO, DESCRIPCION, CANTIDAD, ID_COTIZACION) VALUES ('" & ListView3.ListItems(Conta) & "', '" & ListView3.ListItems(Conta).SubItems(1) & "', " & ListView3.ListItems(Conta).SubItems(2) & ", " & NoCot & ");"
        cnn.Execute (sqlComanda)
    Next Conta
    Me.Command6.Enabled = False
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
    IdClien = "0"
    Me.Command1.Enabled = False
    Me.Command2.Enabled = False
    Me.Command3.Enabled = False
    Me.Command5.Enabled = False
    Me.Command6.Enabled = False
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
        .ColumnHeaders.Add , , "No. CLIENTE", 0
        .ColumnHeaders.Add , , "NOMBRE", 3700
        .ColumnHeaders.Add , , "DIRECCION", 1500
        .ColumnHeaders.Add , , "COLONIA", 1500
        .ColumnHeaders.Add , , "CIUDAD", 1500
        .ColumnHeaders.Add , , "TELEFONO", 1000
    End With
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "CLAVE", 1100
        .ColumnHeaders.Add , , "DESCRIPCION", 3700
    End With
    With ListView3
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "CLAVE", 1100
        .ColumnHeaders.Add , , "DESCRIPCION", 3700
        .ColumnHeaders.Add , , "CANTIDAD", 1500
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
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT ID_CLIENTE, NOMBRE, DIRECCION, COLONIA, CIUDAD, TELEFONO_TRABAJO FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                tLi.SubItems(1) = .Fields("NOMBRE") & ""
                tLi.SubItems(2) = .Fields("DIRECCION") & ""
                tLi.SubItems(3) = .Fields("COLONIA") & ""
                tLi.SubItems(4) = .Fields("CIUDAD") & ""
                tLi.SubItems(5) = .Fields("TELEFONO_TRABAJO") & ""
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
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    IdClien = Item
    Text1.Text = Item.SubItems(1)
    Text2.Text = Item.SubItems(2)
    Text3.Text = Item.SubItems(3)
    Text4.Text = Item.SubItems(4)
    Text5.Text = Item.SubItems(5)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Text9.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Text6.Text = Item
    Text7.Text = Item.SubItems(1)
    Me.Command2.Enabled = True
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub ListView2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Text8.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    INDI = Item.Index
    Me.Command3.Enabled = True
    Me.Command3.SetFocus
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text1_Change()
On Error GoTo ManejaError
    If Text1.Text <> "" Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
    End If
    If ListView3.ListItems.Count <> 0 And Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
        Me.Command6.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
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
    If KeyAscii = 13 And Text1.Text <> "" Then
        Buscar
        Me.ListView1.SetFocus
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
Private Sub BusProd()
    On Error GoTo ManejaError
    'If KeyAscii = 13 And Text2.Text <> "" Then
        Dim tRs As Recordset
        Dim tLi As ListItem
        Dim sBus As String
        Dim SUC As String
        SUC = Menu.Text4(0).Text
        If Option1.Value = True Then
            sBus = "SELECT ID_PRODUCTO, DESCRIPCION, GANANCIA, PRECIO_COSTO, CANTIDAD FROM VSVENTAS WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%' AND SUCURSAL = '" & SUC & "'" 'Cambiado 25/09/06
        End If                                                                                                                                     'Se cambio Almacen3 por VsVentas
        If Option2.Value = True Then
            sBus = "SELECT ID_PRODUCTO, DESCRIPCION, GANANCIA, PRECIO_COSTO, CANTIDAD FROM VSVENTAS WHERE DESCRIPCION LIKE '%" & Text2.Text & "%' AND SUCURSAL = '" & SUC & "'" 'Cambiado 25/09/06
        End If                                                                                                                           'Se cambio Almacen3 por VsVentas
        If sBus <> "" Then
            Set tRs = cnn.Execute(sBus)
            With tRs
                ListView2.ListItems.Clear
                Do While Not .EOF
                    If Not IsNull(.Fields("GANANCIA")) And Not IsNull(.Fields("PRECIO_COSTO")) Then
                        Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                            If Not IsNull(.Fields("DESCRIPCION")) Then tLi.SubItems(1) = .Fields("DESCRIPCION") & ""
                            If Not IsNull(.Fields("GANANCIA")) And Not IsNull(.Fields("PRECIO_COSTO")) Then
                                tLi.SubItems(2) = Format((1 + CDbl(.Fields("GANANCIA"))) * CDbl(.Fields("PRECIO_COSTO")), "0.00")
                            End If
                    End If
                    .MoveNext
                Loop
            End With
        End If
        Me.ListView2.SetFocus
    'Dim Valido As String
        'Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'If KeyAscii > 26 Then
            'If InStr(Valido, Chr(KeyAscii)) = 0 Then
                'KeyAscii = 0
            'End If
    'End If
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
Private Sub Text10_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Command6.SetFocus
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
Private Sub Text10_GotFocus()
    Text10.BackColor = &HFFE1E1
End Sub
Private Sub Text10_LostFocus()
      Text10.BackColor = &H80000005
End Sub
Private Sub Text2_Change()
On Error GoTo ManejaError
    If ListView3.ListItems.Count <> 0 And Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
        Me.Command6.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text3_Change()
On Error GoTo ManejaError
    If ListView3.ListItems.Count <> 0 And Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
        Me.Command6.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text4_Change()
On Error GoTo ManejaError
    If ListView3.ListItems.Count <> 0 And Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
        Me.Command6.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text5_Change()
On Error GoTo ManejaError
    If ListView3.ListItems.Count <> 0 And Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
        Me.Command6.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    Dim Valido As String
    Valido = "1234567890()-"
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
Private Sub Text6_Change()
On Error GoTo ManejaError
    If Text6.Text <> "" And Text7.Text <> "" And Text8.Text <> "" Then
        Me.Command2.Enabled = True
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
Private Sub Text7_Change()
On Error GoTo ManejaError
    If Text6.Text <> "" And Text7.Text <> "" And Text8.Text <> "" Then
        Me.Command2.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text7_GotFocus()
    Text7.BackColor = &HFFE1E1
End Sub
Private Sub Text7_LostFocus()
      Text7.BackColor = &H80000005
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
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
Private Sub Text8_Change()
On Error GoTo ManejaError
    If Text6.Text <> "" And Text7.Text <> "" And Text8.Text <> "" Then
        Me.Command2.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text8_GotFocus()
On Error GoTo ManejaError
    Text8.BackColor = &HFFE1E1
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text8_LostFocus()
      Text8.BackColor = &H80000005
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.Command2.SetFocus
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
Private Sub Text9_Change()
On Error GoTo ManejaError
    If Text9.Text <> "" Then
        Me.Command5.Enabled = True
    Else
        Me.Command5.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text9_GotFocus()
On Error GoTo ManejaError
    Text9.BackColor = &HFFE1E1
    Text9.SelStart = 0
    Text9.SelLength = Len(Text9.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text9_LostFocus()
      Text9.BackColor = &H80000005
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And Text9.Text <> "" Then
        BusProd
        ListView2.SetFocus
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


