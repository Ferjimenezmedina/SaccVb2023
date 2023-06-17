VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form MsgAPToner 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mensajero"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6375
   Icon            =   "MsgAPToner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "MsgAPToner.frx":0442
   Picture         =   "MsgAPToner.frx":B7BD4
   ScaleHeight     =   7080
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prefechados"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3720
      TabIndex        =   18
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   360
      Top             =   6600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Zumbido"
      Height          =   375
      Left            =   2520
      Picture         =   "MsgAPToner.frx":16F366
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   6720
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comunicado !!"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   4680
      TabIndex        =   12
      Top             =   6780
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   0
      Top             =   6600
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000013&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Picture         =   "MsgAPToner.frx":171D38
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000013&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Picture         =   "MsgAPToner.frx":17470A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000013&
      Caption         =   "B"
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
      Left            =   1440
      Picture         =   "MsgAPToner.frx":1770DC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000013&
      Caption         =   "Borrar"
      Height          =   375
      Left            =   5400
      Picture         =   "MsgAPToner.frx":179AAE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   16761024
      CalendarTitleForeColor=   12582912
      Format          =   50528257
      CurrentDate     =   38834
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   5400
      Width           =   5175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000013&
      Caption         =   "Color de Letra"
      Height          =   375
      Left            =   120
      Picture         =   "MsgAPToner.frx":17C480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   4815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "Enviar"
      Height          =   375
      Left            =   5400
      Picture         =   "MsgAPToner.frx":17EE52
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1080
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   5760
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hora :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Enviar a :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   855
   End
End
Attribute VB_Name = "MsgAPToner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private cnn As ADODB.Connection
Dim X As Integer
Dim Posi As Integer
Dim ID(50) As String
Dim CUENCOMBO As Integer
Dim ZumBa As Integer
Private Sub Check2_Click()
    If Check2.Value = 0 Then
        DTPicker1.Value = Format(Date, "dd/mm/yyyy")
        DTPicker1.Visible = False
    Else
        DTPicker1.Visible = True
    End If
End Sub
Private Sub Combo1_DropDown()
On Error GoTo ManejaError
    Combo1.Clear
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim Cont As Integer
    Cont = 0
    sBuscar = "SELECT ID_USUARIO, NOMBRE, APELLIDOS, PUESTO FROM USUARIOS WHERE ESTADO = 'A' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Combo1.AddItem tRs.Fields("NOMBRE") & " " & tRs.Fields("APELLIDOS") & "   (" & tRs.Fields("PUESTO") & ")"
            ID(Cont) = tRs.Fields("ID_USUARIO")
            Cont = Cont + 1
            tRs.MoveNext
        Loop
    End If
    CUENCOMBO = Combo1.ListCount
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
    Posi = Combo1.ListIndex
End Sub
Private Sub Command1_Click()
    If Text1.Text <> "" And Combo1.Text <> "" Then
        Enviar
    End If
End Sub
Private Sub Command2_Click()
    On Error GoTo ManejaError
    CommonDialog1.CancelError = True
    Me.CommonDialog1.ShowColor
    Text2.ForeColor = Me.CommonDialog1.Color
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub Command3_Click()
    If ZumBa = 0 Then
        Text1.Text = "---Enviaste ZuMbiDo---"
        ZumBa = 1
        Timer2.Enabled = True
        Enviar
    Else
        Text1.Text = "---No puedes enviar ZuMbiDos tan seguidos---"
    End If
End Sub
Private Sub Command5_Click()
    If Text1.Text <> "" And Combo1.Text <> "" Then
        Text1.Text = ""
    End If
End Sub
Private Sub Command6_Click()
    If Text2.FontBold = True Then
        Text2.FontBold = False
    Else
        Text2.FontBold = True
    End If
End Sub
Private Sub Command7_Click()
    If Text2.FontItalic = True Then
        Text2.FontItalic = False
    Else
        Text2.FontItalic = True
    End If
End Sub
Private Sub Command8_Click()
    X = X + 2
    Text2.FontSize = X
    If X = 14 Then
        X = 6
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    X = 8
    DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Text3.Text = Format(Time, "hh")
    Text4.Text = Format(Time, "nn")
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_GotFocus()
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 43 Or KeyAscii = 38 Or KeyAscii = 35 Or KeyAscii = 42 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 And Text1.Text <> "" And Combo1.Text <> "" Then
        KeyAscii = 0
        Enviar
    End If
End Sub
Private Sub Timer1_Timer()
On Error GoTo ManejaError
    Text3.Text = Format(Time, "hh")
    Text4.Text = Format(Time, "nn")
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    sBuscar = "SELECT ID_MENSAJE, MENSAJE, ID_USUARIO_DE FROM MSGAPTONER WHERE ID_USUARIO_PARA = '" & VarMen.Text1(0).Text & "' AND FECHA <= '" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            SetForegroundWindow (Me.hWnd)
            If tRs.Fields("MENSAJE") = "---Enviaste ZuMbiDo---" Then
                zumbido
                Text2.Text = Text2.Text & vbCrLf & "--- ZuMbiDo ---"
                sBuscar = "DELETE FROM MSGAPTONER WHERE ID_MENSAJE = " & tRs.Fields("ID_MENSAJE")
                cnn.Execute (sBuscar)
                tRs.MoveNext
            Else
                sBuscar = "SELECT NOMBRE, APELLIDOS, PUESTO FROM USUARIOS WHERE ID_USUARIO = '" & tRs.Fields("ID_USUARIO_DE") & "'"
                Set tRs2 = cnn.Execute(sBuscar)
                Text2.Text = Text2.Text & vbCrLf & tRs2.Fields("NOMBRE") & " " & tRs2.Fields("APELLIDOS") & "    (" & tRs2.Fields("PUESTO") & ") DICE :" & vbCrLf & "     " & tRs.Fields("MENSAJE")
                sBuscar = "DELETE FROM MSGAPTONER WHERE ID_MENSAJE = " & tRs.Fields("ID_MENSAJE")
                cnn.Execute (sBuscar)
                tRs.MoveNext
            End If
        Loop
    End If
    Text2.SelStart = Len(Text2.Text)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Enviar()
On Error GoTo ManejaError
    Dim sBuscar As String
    If Check1.Value = 0 Then
        Text2.Text = Text2.Text & vbCrLf & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text & "(" & VarMen.Text1(3).Text & ") DICE: " & vbCrLf & "     " & Text1.Text
        sBuscar = "INSERT INTO MSGAPTONER (ID_USUARIO_PARA, ID_USUARIO_DE, MENSAJE, FECHA, HORA, MINUTOS) VALUES ('" & ID(Posi) & "', '" & VarMen.Text1(0).Text & "', '" & Text1.Text & "', '" & Format(DTPicker1.Value, "dd/mm/yyyy") & "', " & Text3.Text & ", " & Text4.Text & ");"
        cnn.Execute (sBuscar)
    Else
        Text2.Text = Text2.Text & vbCrLf & "Para : TODOS LOS USUARIOS " & vbCrLf & "     " & Text1.Text
        For X = 0 To CUENCOMBO - 1
            sBuscar = "INSERT INTO MSGAPTONER (ID_USUARIO_PARA, ID_USUARIO_DE, MENSAJE, FECHA, HORA, MINUTOS) VALUES ('" & ID(X) & "', '" & VarMen.Text1(0).Text & "', '" & Text1.Text & "', '" & Format(DTPicker1.Value, "dd/mm/yyyy") & "', " & Text3.Text & ", " & Text4.Text & ");"
            cnn.Execute (sBuscar)
        Next X
    End If
    Text1.Text = ""
    Text2.SelStart = Len(Text2.Text)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub zumbido()
    Dim Cont As Integer
    For Cont = 1 To 50
        MsgAPToner.Top = MsgAPToner.Top - 150
        MsgAPToner.Left = MsgAPToner.Left - 150
        MsgAPToner.Top = MsgAPToner.Top + 150
        MsgAPToner.Left = MsgAPToner.Left + 150
    Next
End Sub
Private Sub Timer2_Timer()
    ZumBa = 0
    Timer2.Enabled = False
End Sub
