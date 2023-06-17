VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXISTENCIAS"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ELIMINAR"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INSERTAR"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "..."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "..."
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim sqlQuery As String
Dim tRs As Recordset
Dim sqlQuery2 As String
Dim tRs2 As Recordset
Dim cId_Producto As String
Dim Cont As Integer

Private Sub Command1_Click()

    If Puede_Insertar Then
        sqlQuery = "SELECT COUNT(*) ID FROM ALMACEN3"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            Me.Label2.Caption = "SE INSERTARAN " & .Fields("ID") & " PRODUCTOS"
            Cont = .Fields("ID")
        End With
         
        With Me.ProgressBar1
            .Max = Cont
            .Value = 0
        End With
        
        sqlQuery = "SELECT * FROM ALMACEN3"
        Set tRs = cnn.Execute(sqlQuery)
        With tRs
            Do While Not .EOF
                cId_Producto = Trim(.Fields("ID_PRODUCTO"))
                sqlQuery2 = "INSERT INTO EXISTENCIAS(ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES('" & cId_Producto & "', 1000, '" & Trim(Me.Text1.Text) & "')"
                Set tRs2 = cnn.Execute(sqlQuery2)
                .MoveNext
                Cont = Cont - 1
                Me.Label1.Caption = "FALTAN " & Cont & " PRODUCTOS"
                Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
            Loop
            Me.Label1.Caption = "SE CAGARON CORRECTAMENTE TODOS LOS ARTICULOS"
        End With
    End If
    
End Sub

Private Sub Command2_Click()
    
    If Puede_Eliminar Then
        sqlQuery = "DELETE EXISTENCIAS WHERE SUCURSAL = '" & Trim(Me.Text1.Text) & "'"
        Set tRs = cnn.Execute(sqlQuery)
        'With tRs
            'Me.Label1.Caption = "SE ELIMINARON " & .RecordCount & " REGISTROS DE EXISTENCIAS"
        'End With
    End If
    
End Sub

Private Sub Form_Load()

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

End Sub

Function Puede_Eliminar() As Boolean

    If Trim(Me.Text1.Text) = "" Then
        MsgBox "ESCRIBA EL NOMBRE DE LA SUCURSAL", vbCritical, "MENSAJE DEL SISTEMA"
        Puede_Eliminar = False
        Exit Function
    End If
    
    Puede_Eliminar = True
        
End Function

Function Puede_Insertar() As Boolean

    If Trim(Me.Text1.Text) = "" Then
        MsgBox "ESCRIBA EL NOMBRE DE LA SUCURSAL", vbCritical, "MENSAJE DEL SISTEMA"
        Puede_Insertar = False
        Exit Function
    End If
    
    Puede_Insertar = True
        
End Function
