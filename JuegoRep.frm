VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form JuegoRep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "JUEGOS DE REPARACION"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   7935
      Left            =   7560
      ScaleHeight     =   7875
      ScaleWidth      =   1155
      TabIndex        =   33
      Top             =   0
      Width           =   1215
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   34
         Top             =   6600
         Width           =   975
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "JuegoRep.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "JuegoRep.frx":030A
            Top             =   240
            Width           =   705
         End
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
            TabIndex        =   35
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   480
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   360
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   240
      MaxLength       =   8
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3960
      TabIndex        =   29
      Top             =   7320
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5880
      TabIndex        =   28
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   7320
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   6240
      Top             =   360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "JUEGO_REPARACION"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
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
      Height          =   375
      Left            =   6120
      Picture         =   "JuegoRep.frx":1DBC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6240
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ALMACEN3"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionados"
      Height          =   2655
      Left            =   120
      TabIndex        =   25
      Top             =   5160
      Width           =   7215
      Begin VB.CommandButton Command5 
         Caption         =   "Juego Nuevo"
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
         Left            =   4200
         Picture         =   "JuegoRep.frx":478E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
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
         Left            =   5760
         Picture         =   "JuegoRep.frx":7160
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2160
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1695
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
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
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   1
      Left            =   6000
      TabIndex        =   24
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccion"
      Height          =   3015
      Left            =   120
      TabIndex        =   21
      Top             =   2040
      Width           =   7215
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3825
         TabIndex        =   9
         Top             =   360
         Width           =   870
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         Enabled         =   0   'False
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
         Left            =   5760
         Picture         =   "JuegoRep.frx":9B32
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2520
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2778
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
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5520
         MaxLength       =   8
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         MaxLength       =   20
         TabIndex        =   8
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label8 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   4800
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Busqueda"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "JuegoRep.frx":C504
      Left            =   600
      List            =   "JuegoRep.frx":C517
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3480
      TabIndex        =   19
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   4440
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   100
      TabIndex        =   2
      Top             =   720
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label6 
      Caption         =   "Color"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Marca"
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Cantidad maxima"
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad minima"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Clave del Producto"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "JuegoRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Combo1_Change()
    Combo1.Clear
    Buscarcbo
End Sub
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HFFE1E1
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text1(2).SetFocus
    End If
    KeyAscii = 0
End Sub

Private Sub Combo1_LostFocus()
    Combo1.BackColor = &H80000005
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
    If Combo1.Text <> "" Then
        Text1(5).Text = Combo1.Text
    End If
End Sub
Private Sub Combo2_GotFocus()
    Combo2.BackColor = &HFFE1E1
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
    End If
    KeyAscii = 0
End Sub
Private Sub Combo2_LostFocus()
    Combo2.BackColor = &H80000005
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)
    Text1(4).Text = Combo2.Text
End Sub
Private Sub Command1_Click()
On Error GoTo ManejaError
    Text3(1).Text = Text4.Text
    If Text3(0).Text <> "" And Text3(1).Text <> "" And Text3(2).Text <> "" Then
        Adodc3.Recordset.AddNew
    Else
        MsgBox ("Falta Informacion...")
    End If
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    Dim Clave As String
    Clave = Text4.Text
    sBuscar = "SELECT * FROM JUEGO_REPARACION WHERE ID_REPARACION ='" & Clave & "'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
           Set tLi = ListView2.ListItems.Add(, , "VACIO")
                tLi.SubItems(1) = "VACIO"
                tLi.SubItems(2) = "VACIO"
        Else
            ListView2.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_REPARACION") & "")
                tLi.SubItems(1) = .Fields("ID_PRODUCTO") & ""
                tLi.SubItems(2) = .Fields("CANTIDAD") & ""
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
Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ListView2.SetFocus
    End If
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    Dim Clave As String
    Clave = Text4.Text
    sBuscar = "DELETE FROM JUEGO_REPARACION WHERE ID_PRODUCTO = '" & Text5.Text & "'"
    Set tRs = cnn.Execute(sBuscar)
    sBuscar = "SELECT * FROM JUEGO_REPARACION WHERE ID_REPARACION ='" & Clave & "'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            ListView2.ListItems.Clear
            Set tLi = ListView2.ListItems.Add(, , "VACIO")
                tLi.SubItems(1) = "VACIO"
                tLi.SubItems(2) = "VACIO"
        Else
            ListView2.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_REPARACION") & "")
                tLi.SubItems(1) = .Fields("ID_PRODUCTO") & ""
                tLi.SubItems(2) = .Fields("CANTIDAD") & ""
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
Private Sub Command3_Click()
On Error GoTo ManejaError
    Text3(1).Text = Text1(0).Text
    Text4.Text = Text1(0).Text
    Text3(1).Text = Text1(0).Text
    Text1(6).Text = "COMPUESTO"
    Text1(7).Text = "0"
    Text1(8).Text = "0"
    If Text1(0) <> "" And Text1(1) <> "" And Text1(2) <> "" And Text1(3) <> "" And Text1(5) <> "" Then
        Adodc1.Recordset.AddNew
        Text1(0).Text = Text4.Text
        Command1.Enabled = True
        Command2.Enabled = True
        Command3.Enabled = False
        Text3(0).Enabled = True
        Text2.Enabled = True
        Text1(0).Enabled = False
        Text1(1).Enabled = False
        Text1(2).Enabled = False
        Text1(3).Enabled = False
        Text1(5).Enabled = False
        Combo1.Enabled = False
        Combo2.Enabled = False
    Else
        MsgBox ("Falta Informacion Necesaria")
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

Private Sub Command5_Click()
        Command1.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = True
        Text3(0).Enabled = False
        Text2.Enabled = False
        Text1(0).Enabled = True
        Text1(1).Enabled = True
        Text1(2).Enabled = True
        Text1(3).Enabled = True
        Text1(5).Enabled = True
        Combo1.Enabled = True
        Combo2.Enabled = True
End Sub
Private Sub Form_Load()
On Error GoTo ManejaError
    Const sPathBase As String = "LINUX"
    Dim i As Long
    For i = 0 To 8
        Set Text1(i).DataSource = Adodc1
    Next

    Text1(0).DataField = "ID_PRODUCTO"
    Text1(1).DataField = "DESCRIPCION"
    Text1(2).DataField = "C_MINIMA"
    Text1(3).DataField = "C_MAXIMA"
    Text1(4).DataField = "COLOR"
    Text1(5).DataField = "MARCA"
    Text1(6).DataField = "TIPO"
    Text1(7).DataField = "GANANCIA"
    Text1(8).DataField = "PRECIO_COSTO"
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    rst.Open "SELECT * FROM ALMACEN1", cnn, adOpenDynamic, adLockOptimistic
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "CLAVE DEL PRODUCTO", 2500
        .ColumnHeaders.Add , , "DESCRIPCION", 5400
    End With
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "CLAVE DEL JUEGO DE REPARACION", 3100
        .ColumnHeaders.Add , , "CLAVE DEL PRODUCTO", 2600
        .ColumnHeaders.Add , , "CANTIDAD", 1300
    End With

    Dim X As Long
    For X = 0 To 2
        Set Text3(X).DataSource = Adodc3
    Next
    Text3(0).DataField = "CANTIDAD"
    Text3(2).DataField = "ID_PRODUCTO"
    Text3(1).DataField = "ID_REPARACION"
    
    Adodc1.Recordset.AddNew
    Adodc3.Recordset.AddNew
    Buscarcbo
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
Private Sub Buscarcbo()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    sBuscar = "SELECT * FROM MARCA ORDER BY MARCA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        tRs.MoveFirst
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("MARCA")
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
    Text3(2).Text = Item
    Text2.Text = Item
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command1.SetFocus
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text5.Text = Me.ListView2.SelectedItem.SubItems(1)
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    If Index = 5 Then
        Combo1.Visible = True
        Text1(5).Visible = False
        Combo1.Text = Text1(5).Text
    End If
    Text1(Index).BackColor = &HFFE1E1
    Text1(Index).SetFocus
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 0 Then
            Text1(1).SetFocus
        End If
        If Index = 1 Then
            Combo2.SetFocus
        End If
        If Index = 2 Then
            Text1(3).SetFocus
        End If
        If Index = 3 Then
            Text2.SetFocus
        End If
    End If
    Dim Valido As String
    Valido = "1234567890.abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = &H80000005
End Sub

Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Dim tRs As Recordset
        Dim tLi As ListItem
        Dim sBus As String
        sBus = "SELECT * FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%'"
        Set tRs = cnn.Execute(sBus)
        With tRs
            ListView1.ListItems.Clear
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                If Not IsNull(.Fields("DESCRIPCION")) Then tLi.SubItems(1) = .Fields("DESCRIPCION") & ""
                .MoveNext
            Loop
        End With
        sBus = "SELECT * FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%'"
        Set tRs = cnn.Execute(sBus)
        With tRs
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                If Not IsNull(.Fields("DESCRIPCION")) Then tLi.SubItems(1) = .Fields("DESCRIPCION") & ""
                .MoveNext
            Loop
        End With
        ListView1.SetFocus
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

Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    Text3(Index).BackColor = &HFFE1E1
    Text3(Index).SetFocus
    Text3(Index).SelStart = 0
    Text3(Index).SelLength = Len(Text3(Index).Text)
End Sub
Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
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

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).BackColor = &H80000005
End Sub
