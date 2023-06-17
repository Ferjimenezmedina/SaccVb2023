VERSION 5.00
Begin VB.Form FrmRespaldoBD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Respaldar la Base de Datos"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   3720
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7095
   End
End
Attribute VB_Name = "FrmRespaldoBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1

Private Sub Command2_Click()
    Dim sBuscar As String
    Dim tRs As Recordset
    sBuscar = "SELECT  * From PARACHECARCREDITOSDEVENTAS WHERE     (TOTAL_COMPRA <> Tot_Venta)"
    Set tRs = cnn.Execute(sBuscar)
    Do While Not (tRs.EOF)
        sBuscar = "UPDATE CUENTAS SET TOTAL_COMPRA = " & Replace(tRs.Fields("TOT_VENTA"), ",", ".") & " WHERE ID_CUENTA = " & tRs.Fields("ID_CUENTA")
        cnn.Execute (sBuscar)
        tRs.MoveNext
    Loop
    MsgBox "YA! TATA!"
End Sub

Private Sub Form_Load()
    Dim Guarda As String
    Dim sBuscar As String
    Const sPathBase As String = "LINUX"
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    On Error GoTo HasRespaldo
    Guarda = "C:\RespaldoSACC" & Date & ".Bak"
    Guarda = Replace(Guarda, "/", "-")
    MsgBox "" & GetAttr(Guarda)
    GetAttr (Guarda)
    sBuscar = "BACKUP DATABASE APTONER TO DISK = '" & Guarda & "' WITH FORMAT,NAME = 'res'"
    cnn.Execute (sBuscar)
    Exit Sub
HasRespaldo:
    sBuscar = "BACKUP DATABASE APTONER TO DISK = '" & Guarda & "' WITH FORMAT,NAME = 'res'"
    cnn.Execute (sBuscar)
End Sub
