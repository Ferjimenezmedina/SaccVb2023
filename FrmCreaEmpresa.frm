VERSION 5.00
Begin VB.Form FrmNuevaEmpresa 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crear Empresa"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Height          =   375
      Left            =   2640
      Picture         =   "FrmCreaEmpresa.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
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
      Left            =   1080
      Picture         =   "FrmCreaEmpresa.frx":29D2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   2
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Servidor :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Base de Datos :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Empresa :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmNuevaEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private cnn1 As ADODB.Connection
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    If MsgBox("ESTÁ SEGURO QUE DESEA CEAR LA EMPRESA " & Text1.Text & "?", vbYesNo) = vbYes Then
        Dim sBaseDatos As String
        Dim sBuscar As String
        Dim s As String
        Dim tRs As ADODB.Recordset
        If Text2.Text = "" Then
            CreaNombreBD
        End If
        sBaseDatos = "EMPRESAS_SACC"
        Set cnn1 = New ADODB.Connection
        With cnn1
            .ConnectionString = _
                "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & sBaseDatos & ";Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDORPPAL", "") & ";"  '& NvoMen.txtServidor.Text & ";"
            .Open
        End With
        sBuscar = "SELECT ID From EMPRESAS WHERE EMPRESA = '" & Text1.Text & "' OR BASE_DATOS = '" & Text2.Text & "'"
        Set tRs = cnn1.Execute(sBuscar)
        If (tRs.EOF And tRs.BOF) Then
            s = mdlCreaDataBase.CreaBd(Text1.Text, Text3.Text, Text2.Text, NvoMen.TxtUsuario.Text, NvoMen.TxtContrasena.Text, NvoMen.TxtProvider.Text)
            MsgBox "EMPRESA CREADA CON EXITO", vbInformation
        Else
            MsgBox "LA EMPRESA O BASE DE DATOS YA EXISTEN ACTUALMENTE", vbExclamation
        End If
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Text3.Text = GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "")
End Sub
Private Sub CreaNombreBD()
    Dim sNombre As String
    sNombre = Text1.Text
    sNombre = Replace(sNombre, " ", "_")
    sNombre = Replace(sNombre, "!", "")
    sNombre = Replace(sNombre, ".", "")
    sNombre = Replace(sNombre, "#", "")
    sNombre = Replace(sNombre, "$", "")
    sNombre = Replace(sNombre, "&", "")
    sNombre = Replace(sNombre, "/", "")
    sNombre = Replace(sNombre, "(", "")
    sNombre = Replace(sNombre, ")", "")
    sNombre = Replace(sNombre, "=", "")
    sNombre = Replace(sNombre, "?", "")
    sNombre = Replace(sNombre, "'", "")
    sNombre = Replace(sNombre, "¡", "")
    sNombre = Replace(sNombre, "¿", "")
    sNombre = Replace(sNombre, "¨", "")
    sNombre = Replace(sNombre, "*", "")
    sNombre = Replace(sNombre, "+", "")
    sNombre = Replace(sNombre, "´", "")
    sNombre = Replace(sNombre, "[", "")
    sNombre = Replace(sNombre, "]", "")
    sNombre = Replace(sNombre, "{", "")
    sNombre = Replace(sNombre, "}", "")
    sNombre = Replace(sNombre, "\", "")
    sNombre = Replace(sNombre, "~", "")
    sNombre = Replace(sNombre, "¬", "")
    sNombre = Replace(sNombre, ";", "")
    sNombre = Replace(sNombre, ",", "")
    sNombre = Replace(sNombre, ".", "")
    sNombre = Replace(sNombre, ":", "")
    Text2.Text = sNombre
End Sub
Private Sub Text1_Change()
    CreaNombreBD
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ ._'+&$#"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
