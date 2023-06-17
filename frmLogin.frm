VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2565
   ClientLeft      =   15
   ClientTop       =   -15
   ClientWidth     =   4575
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      Picture         =   "frmLogin.frx":1601A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      Picture         =   "frmLogin.frx":169FC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
      Begin VB.TextBox txtPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   0
         Picture         =   "frmLogin.frx":1763E
         Top             =   0
         Width           =   450
      End
      Begin VB.Label lblPass 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
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
         Left            =   -480
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   4095
      Begin VB.TextBox txtUser 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.Image Image4 
         Height          =   435
         Left            =   0
         Picture         =   "frmLogin.frx":18200
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblUser 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIO"
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
         Left            =   -480
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   480
      Picture         =   "frmLogin.frx":18D22
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   1080
      Picture         =   "frmLogin.frx":198A4
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label Label4 
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
      Left            =   720
      TabIndex        =   9
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Aceptar"
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
      Left            =   3000
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim cPassword As String
Private Sub CmdAceptar_Click()
On Error GoTo ManejaError
    If Password(Trim(Me.txtUser.Text)) = True Then
        If Trim(Me.txtPass.Text) = cPassword Then
            SaveSetting "APTONER", "ConfigSACC", "ULTIMOUSUARIO", txtUser.Text
            Unload Me
        Else
            Me.txtPass.SetFocus
            Me.txtPass.SelStart = 0
            Me.txtPass.SelLength = Len(Me.txtPass.Text)
            Me.lblPass.ForeColor = vbRed
            Me.lblUser.ForeColor = vbBlack
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Password(Clave As String) As Boolean
On Error GoTo ManejaError
    Dim cnn1 As ADODB.Connection
    Dim tRs1 As ADODB.Recordset
    Dim sBuscar As String
    Set cnn1 = New ADODB.Connection
    With cnn1
        .ConnectionString = _
        "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    sBuscar = "SELECT U.ID_USUARIO, U.NOMBRE, U.APELLIDOS, U.PUESTO, U.PASSWORD, U.ID_SUCURSAL, U.DEPARTAMENTO, U.PE1, U.PE2, U.PE3, U.PE4, U.PE5, U.PE6, U.PE7, U.PE8, U.PE9, U.PE10, U.PE11, U.PE12, U.PE13, U.PE14, U.PE15, U.PE16, U.PE17, U.PE18, U.PE19, U.PE20, U.PE21, U.PE22, U.PE23, U.PE24, U.PE25, U.PE26, U.PE27, U.PE28, U.PE29, U.PE30, U.PE31, U.PE32, U.PE33, U.PE34, U.PE35, U.PE36, U.PE37, U.PE38, U.PE39, U.PE40, U.PE41, U.PE42, U.PE43, U.PE44, U.PE45, U.PE46, U.PE47, U.PE48, U.PE49, U.PE50, U.PE51, U.PE52, U.PE53, U.PE54, U.PE55, U.PE56, U.PE57, U.PE58, U.PE59, U.PE60, U.PE61, U.PE62, U.PE63, U.PE64, U.PE65, U.PE66, U.PE67, U.PE68, U.PE69, U.PE70, S.NOMBRE AS NombreSucursal, S.CALLE, S.COLONIA, S.CIUDAD, S.ESTADO, S.TELEFONO, S.IVA, S.APLICA_PROMO FROM Usuarios AS U JOIN Sucursales AS S ON U.Id_Sucursal = S.Id_Sucursal WHERE U.Nombre = '" & Clave & "' AND U.ESTADO = 'A'"
    Set tRs1 = cnn1.Execute(sBuscar)
    With tRs1
        If Not (.EOF And .BOF) Then
            Password = True
            If Not IsNull(.Fields("ID_USUARIO")) Then VarMen.Text1(0).Text = .Fields("ID_USUARIO")
            If Not IsNull(.Fields("NOMBRE")) Then VarMen.Text1(1).Text = Trim(.Fields("NOMBRE"))
            If Not IsNull(.Fields("APELLIDOS")) Then VarMen.Text1(2).Text = Trim(.Fields("APELLIDOS"))
            If Not IsNull(.Fields("PUESTO")) Then VarMen.Text1(3).Text = Trim(.Fields("PUESTO"))
            If Not IsNull(.Fields("PASSWORD")) Then cPassword = Trim(.Fields("PASSWORD"))
            If Not IsNull(.Fields("PASSWORD")) Then VarMen.Text1(4).Text = Trim(.Fields("PASSWORD"))
            If Not IsNull(.Fields("ID_SUCURSAL")) Then VarMen.Text1(5).Text = .Fields("ID_SUCURSAL")
            If Not IsNull(.Fields("PE1")) Then VarMen.Text1(6).Text = .Fields("PE1")
            If Not IsNull(.Fields("PE2")) Then VarMen.Text1(7).Text = .Fields("PE2")
            If Not IsNull(.Fields("PE3")) Then VarMen.Text1(8).Text = .Fields("PE3")
            If Not IsNull(.Fields("PE4")) Then VarMen.Text1(9).Text = .Fields("PE4")
            If Not IsNull(.Fields("PE5")) Then VarMen.Text1(10).Text = .Fields("PE5")
            If Not IsNull(.Fields("PE6")) Then VarMen.Text1(11).Text = .Fields("PE6")
            If Not IsNull(.Fields("PE7")) Then VarMen.Text1(12).Text = .Fields("PE7")
            If Not IsNull(.Fields("PE8")) Then VarMen.Text1(13).Text = .Fields("PE8")
            If Not IsNull(.Fields("PE9")) Then VarMen.Text1(14).Text = .Fields("PE9")
            If Not IsNull(.Fields("PE10")) Then VarMen.Text1(15).Text = .Fields("PE10")
            If Not IsNull(.Fields("PE11")) Then VarMen.Text1(16).Text = .Fields("PE11")
            If Not IsNull(.Fields("PE12")) Then VarMen.Text1(17).Text = .Fields("PE12")
            If Not IsNull(.Fields("PE13")) Then VarMen.Text1(18).Text = .Fields("PE13")
            If Not IsNull(.Fields("PE14")) Then VarMen.Text1(19).Text = .Fields("PE14")
            If Not IsNull(.Fields("PE15")) Then VarMen.Text1(20).Text = .Fields("PE15")
            If Not IsNull(.Fields("PE16")) Then VarMen.Text1(21).Text = .Fields("PE16")
            If Not IsNull(.Fields("PE17")) Then VarMen.Text1(22).Text = .Fields("PE17")
            If Not IsNull(.Fields("PE18")) Then VarMen.Text1(23).Text = .Fields("PE18")
            If Not IsNull(.Fields("PE19")) Then VarMen.Text1(24).Text = .Fields("PE19")
            If Not IsNull(.Fields("PE21")) Then VarMen.Text1(26).Text = .Fields("PE21")
            If Not IsNull(.Fields("PE56")) Then VarMen.Text1(25).Text = .Fields("PE56")
            If Not IsNull(.Fields("PE63")) Then VarMen.Text1(27).Text = .Fields("PE63")
            If Not IsNull(.Fields("PE23")) Then VarMen.Text1(28).Text = .Fields("PE23")
            If Not IsNull(.Fields("PE24")) Then VarMen.Text1(29).Text = .Fields("PE24")
            If Not IsNull(.Fields("PE25")) Then VarMen.Text1(30).Text = .Fields("PE25")
            If Not IsNull(.Fields("PE26")) Then VarMen.Text1(31).Text = .Fields("PE26")
            If Not IsNull(.Fields("PE27")) Then VarMen.Text1(32).Text = .Fields("PE27")
            If Not IsNull(.Fields("PE28")) Then VarMen.Text1(33).Text = .Fields("PE28")
            If Not IsNull(.Fields("PE29")) Then VarMen.Text1(34).Text = .Fields("PE29")
            If Not IsNull(.Fields("PE30")) Then VarMen.Text1(76).Text = .Fields("PE30")
            If Not IsNull(.Fields("PE31")) Then VarMen.Text1(36).Text = .Fields("PE31")
            If Not IsNull(.Fields("PE32")) Then VarMen.Text1(37).Text = .Fields("PE32")
            If Not IsNull(.Fields("PE33")) Then VarMen.Text1(38).Text = .Fields("PE33")
            If Not IsNull(.Fields("PE34")) Then VarMen.Text1(39).Text = .Fields("PE34")
            If Not IsNull(.Fields("PE35")) Then VarMen.Text1(40).Text = .Fields("PE35")
            If Not IsNull(.Fields("PE36")) Then VarMen.Text1(41).Text = .Fields("PE36")
            If Not IsNull(.Fields("PE37")) Then VarMen.Text1(42).Text = .Fields("PE37")
            If Not IsNull(.Fields("PE38")) Then VarMen.Text1(43).Text = .Fields("PE38")
            If Not IsNull(.Fields("PE39")) Then VarMen.Text1(44).Text = .Fields("PE39")
            If Not IsNull(.Fields("PE40")) Then VarMen.Text1(45).Text = .Fields("PE40")
            If Not IsNull(.Fields("PE41")) Then VarMen.Text1(46).Text = .Fields("PE41")
            If Not IsNull(.Fields("PE42")) Then VarMen.Text1(47).Text = .Fields("PE42")
            If Not IsNull(.Fields("PE43")) Then VarMen.Text1(48).Text = .Fields("PE43")
            If Not IsNull(.Fields("PE44")) Then VarMen.Text1(49).Text = .Fields("PE44")
            If Not IsNull(.Fields("PE45")) Then VarMen.Text1(50).Text = .Fields("PE45")
            If Not IsNull(.Fields("PE46")) Then VarMen.Text1(51).Text = .Fields("PE46")
            If Not IsNull(.Fields("PE47")) Then VarMen.Text1(52).Text = .Fields("PE47")
            If Not IsNull(.Fields("PE48")) Then VarMen.Text1(53).Text = .Fields("PE48")
            If Not IsNull(.Fields("PE49")) Then VarMen.Text1(54).Text = .Fields("PE49")
            If Not IsNull(.Fields("PE50")) Then VarMen.Text1(55).Text = .Fields("PE50")
            If Not IsNull(.Fields("PE51")) Then VarMen.Text1(56).Text = .Fields("PE51")
            If Not IsNull(.Fields("PE52")) Then VarMen.Text1(57).Text = .Fields("PE52")
            If Not IsNull(.Fields("PE53")) Then VarMen.Text1(58).Text = .Fields("PE53")
            If Not IsNull(.Fields("PE54")) Then VarMen.Text1(59).Text = .Fields("PE54")
            If Not IsNull(.Fields("PE55")) Then VarMen.Text1(60).Text = .Fields("PE55")
            If Not IsNull(.Fields("PE22")) Then VarMen.Text1(61).Text = .Fields("PE22")
            If Not IsNull(.Fields("PE57")) Then VarMen.Text1(62).Text = .Fields("PE57")
            If Not IsNull(.Fields("PE58")) Then VarMen.Text1(63).Text = .Fields("PE58")
            If Not IsNull(.Fields("PE59")) Then VarMen.Text1(64).Text = .Fields("PE59")
            If Not IsNull(.Fields("PE60")) Then VarMen.Text1(65).Text = .Fields("PE60")
            If Not IsNull(.Fields("PE61")) Then VarMen.Text1(66).Text = .Fields("PE61")
            If Not IsNull(.Fields("PE62")) Then VarMen.Text1(67).Text = .Fields("PE62")
            If Not IsNull(.Fields("PE68")) Then VarMen.Text1(68).Text = .Fields("PE68")
            If Not IsNull(.Fields("PE64")) Then VarMen.Text1(69).Text = .Fields("PE64")
            If Not IsNull(.Fields("PE65")) Then VarMen.Text1(70).Text = .Fields("PE65")
            If Not IsNull(.Fields("PE66")) Then VarMen.Text1(71).Text = .Fields("PE66")
            If Not IsNull(.Fields("PE67")) Then VarMen.Text1(72).Text = .Fields("PE67")
            'If Not ISNULL(.Fields("PE68")) Then VarMen.Text1(73).Text = .Fields("PE68")
            If Not IsNull(.Fields("PE20")) Then VarMen.Text1(74).Text = .Fields("PE20")
            If Not IsNull(.Fields("DEPARTAMENTO")) Then VarMen.Text1(75).Text = .Fields("DEPARTAMENTO")
            If Not IsNull(.Fields("PE69")) Then VarMen.Text1(77).Text = .Fields("PE69")
            If Not IsNull(.Fields("PE70")) Then VarMen.Text1(78).Text = .Fields("PE70")
            If Not IsNull(.Fields("NombreSucursal")) Then VarMen.Text4(0).Text = Trim(.Fields("NombreSucursal"))
            If Not IsNull(.Fields("CALLE")) Then VarMen.Text4(1).Text = Trim(.Fields("CALLE"))
            If Not IsNull(.Fields("COLONIA")) Then VarMen.Text4(2).Text = Trim(.Fields("COLONIA"))
            If Not IsNull(.Fields("CIUDAD")) Then VarMen.Text4(3).Text = Trim(.Fields("CIUDAD"))
            If Not IsNull(.Fields("ESTADO")) Then VarMen.Text4(4).Text = Trim(.Fields("ESTADO"))
            If Not IsNull(.Fields("TELEFONO")) Then VarMen.Text4(5).Text = Trim(.Fields("TELEFONO"))
            If Not IsNull(.Fields("ID_SUCURSAL")) Then VarMen.Text4(6).Text = Trim(.Fields("ID_SUCURSAL"))
            If Not IsNull(.Fields("IVA")) Then VarMen.Text4(7).Text = Trim(.Fields("IVA"))
            If Not IsNull(.Fields("APLICA_PROMO")) Then VarMen.Text4(8).Text = Trim(.Fields("APLICA_PROMO"))
            sBuscar = "SELECT * FROM EMPRESA"
            Set tRs1 = cnn1.Execute(sBuscar)
            If Not (tRs1.EOF And tRs1.BOF) Then
                If Not IsNull(tRs1.Fields("NOMBRE")) Then VarMen.TxtEmp(0).Text = tRs1.Fields("NOMBRE")
                If Not IsNull(tRs1.Fields("NOMBRE")) Then NvoMen.LblEmpresa.Caption = tRs1.Fields("NOMBRE")
                If Not IsNull(tRs1.Fields("DIRECCION")) Then VarMen.TxtEmp(1).Text = tRs1.Fields("DIRECCION")
                If Not IsNull(tRs1.Fields("TELEFONO")) Then VarMen.TxtEmp(2).Text = tRs1.Fields("TELEFONO")
                If Not IsNull(tRs1.Fields("FAX")) Then VarMen.TxtEmp(3).Text = tRs1.Fields("FAX")
                If Not IsNull(tRs1.Fields("COLONIA")) Then VarMen.TxtEmp(4).Text = tRs1.Fields("COLONIA")
                If Not IsNull(tRs1.Fields("CD")) Then VarMen.TxtEmp(5).Text = tRs1.Fields("CD")
                If Not IsNull(tRs1.Fields("ESTADO")) Then VarMen.TxtEmp(6).Text = tRs1.Fields("ESTADO")
                If Not IsNull(tRs1.Fields("PAIS")) Then VarMen.TxtEmp(7).Text = tRs1.Fields("PAIS")
                If Not IsNull(tRs1.Fields("RFC")) Then VarMen.TxtEmp(8).Text = tRs1.Fields("RFC")
                If Not IsNull(tRs1.Fields("CP")) Then VarMen.TxtEmp(9).Text = tRs1.Fields("CP")
                If Not IsNull(tRs1.Fields("FACTURA_ELECTRONICA")) Then VarMen.TxtEmp(10).Text = tRs1.Fields("FACTURA_ELECTRONICA")
                If Not IsNull(tRs1.Fields("REPRESENTANTE")) Then VarMen.TxtEmp(11).Text = tRs1.Fields("REPRESENTANTE")
                VarMen.TxtEmp(12).Text = GetSetting("APTONER", "ConfigSACC", "COMPRAS_NAC_INT", "EXTENDIDO")
            End If
        Else
            Password = False
            Me.txtUser.SetFocus
            Me.txtUser.SelStart = 0
            Me.txtUser.SelLength = Len(Me.txtUser.Text)
            Me.lblUser.ForeColor = vbRed
            Me.lblPass.ForeColor = vbBlack
        End If
        .Close
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub cmdCancelar_Click()
    End
End Sub
Private Sub Form_Load()
On Error GoTo ManejaError
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    If GetSetting("APTONER", "ConfigSACC", "ULTIMOUSUARIO", "0") <> "0" Then
        txtUser.Text = GetSetting("APTONER", "ConfigSACC", "ULTIMOUSUARIO", "0")
        txtUser.SelStart = 0
        txtUser.SelLength = Len(txtUser.Text)
        Label1.Caption = GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER")
        Set cnn = New ADODB.Connection
        With cnn
            .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
            .Open
        End With
    End If
    Set VarMen = NvoMen
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtPass_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then Me.cmdAceptar.Value = True
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtUser_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Me.txtPass.SetFocus
        KeyAscii = 0
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
