VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRegDomi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captura de Recoleccion de Cartuchos a Domicilio"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   7560
      ScaleHeight     =   4875
      ScaleWidth      =   1755
      TabIndex        =   21
      Top             =   0
      Width           =   1815
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   22
         Top             =   2400
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
            TabIndex        =   23
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image9 
            Height          =   870
            Left            =   120
            MouseIcon       =   "FrmRegDomi.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "FrmRegDomi.frx":030A
            Top             =   120
            Width           =   720
         End
      End
   End
   Begin VB.CommandButton BtnNueColonia 
      Caption         =   "Colonia"
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
      Left            =   4800
      Picture         =   "FrmRegDomi.frx":23EC
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ComboBox CmbColonia 
      Height          =   315
      Left            =   1800
      TabIndex        =   19
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox TxtNoArticulos 
      Height          =   285
      Left            =   4560
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "0"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox TxtTelefonoDomi 
      Height          =   285
      Left            =   1800
      MaxLength       =   9
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton BtnGuardaDomi 
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
      Left            =   6240
      Picture         =   "FrmRegDomi.frx":4DBE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox TxtNotaDomi 
      Height          =   855
      Left            =   3600
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2160
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTPFechaDomi 
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   46792705
      CurrentDate     =   38833
   End
   Begin VB.TextBox TxtDomiCleinte 
      Height          =   285
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   2
      Top             =   720
      Width           =   5655
   End
   Begin VB.TextBox TxtNomCliente 
      Height          =   285
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Favor de pasar en el horario"
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   2895
      Begin VB.TextBox TxtHoraDe 
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtHoraAl 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Entra la Hora :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Y la Hora :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Label Label9 
      Caption         =   "# Articulos :"
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Telefono :"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Nota :"
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha :"
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Colonia :"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Domicilio :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "* Nombre del Cliente :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "FrmRegDomi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Private Sub CmbColonia_DropDown()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    CmbColonia.Clear
    sBuscar = "SELECT NOMBRE FROM COLONIAS ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("NOMBRE")) Then
                CmbColonia.AddItem tRs.Fields("NOMBRE")
            End If
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
'Private Sub Command1_Click()
'    Unload Me
'End Sub
Private Sub BtnGuardaDomi_Click()
On Error GoTo ManejaError
    If TxtNomCliente.Text = "" Or TxtDomiCleinte.Text = "" Or CmbColonia.Text = "" Or TxtNoArticulos.Text = "" Then
        MsgBox "FALTA INFORMACIÓN NECESARIA!", vbInformation, "SACC"
    Else
        Dim sBuscar As String
        Dim tRs As Recordset
        sBuscar = "SELECT ZONA FROM COLONIAS WHERE NOMBRE = '" & CmbColonia.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        sBuscar = "INSERT INTO DOMICILIOS (NUM_ARTICULOS, NOM_CLIENTE, DOMICILIO, COLONIA, TELEFONO, FECHA, DE_HORA, A_HORA, NOTA, ZONA, ESTADO) VALUES ('" & TxtNoArticulos.Text & "', '" & TxtNomCliente.Text & "', '" & TxtDomiCleinte.Text & "', '" & CmbColonia.Text & "','" & TxtTelefonoDomi.Text & "', " & DTPFechaDomi.Value & ", '" & TxtHoraDe.Text & "', '" & TxtHoraAl.Text & "', '" & TxtNotaDomi.Text & "', '" & tRs.Fields("ZONA") & "', 'P');"
        cnn.Execute (sBuscar)
        sBuscar = "SELECT ID_DOMICILIO FROM DOMICILIOS ORDER BY ID_DOMICILIO DESC"
        Set tRs = cnn.Execute(sBuscar)
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
        Printer.Print "SUCURSAL : " & VarMen.Text4(0).Text
        Printer.Print "No. DE COMANDA : " & tRs.Fields("ID_DOMICILIO") & "-DOMI"
        Printer.Print "ATENDIDO POR : " & VarMen.Text1(1).Text & VarMen.Text1(1).Text
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "                       RECOLECCION A DOMICILIO"
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print ""
        Printer.Print "CLIENTE :"
        Printer.Print Mid(TxtNomCliente.Text, 1, 25) & "-"
        If Len(TxtNomCliente.Text) > 26 Then
            Printer.Print Mid(TxtNomCliente.Text, 26, 50) & "-"
        End If
        If Len(TxtNomCliente.Text) > 51 Then
            Printer.Print Mid(TxtNomCliente.Text, 51, 75) & "-"
        End If
        If Len(TxtNomCliente.Text) > 76 Then
            Printer.Print Mid(TxtNomCliente.Text, 76, 100) & "-"
        End If
        Printer.Print "DOMICILIO :"
        Printer.Print Mid(TxtDomiCleinte.Text, 1, 25) & "-"
        If Len(TxtDomiCleinte.Text) > 26 Then
            Printer.Print Mid(TxtDomiCleinte.Text, 26, 50) & "-"
        End If
        If Len(TxtDomiCleinte.Text) > 51 Then
            Printer.Print Mid(TxtDomiCleinte.Text, 51, 75) & "-"
        End If
        If Len(TxtDomiCleinte.Text) > 76 Then
            Printer.Print Mid(TxtDomiCleinte.Text, 76, 100) & "-"
        End If
        Printer.Print "COLONIA : " & CmbColonia.Text
        Printer.Print "TELEFONO : " & TxtTelefonoDomi.Text
        Printer.Print "FECHA : " & DTPFechaDomi.Value
        Printer.Print "ENTRE LAS " & TxtHoraDe.Text & " Y LAS " & TxtHoraAl.Text
        Printer.Print "RECOGER " & TxtNoArticulos.Text & " ARTICULOS"
        Printer.Print "NOTAS :"
        Printer.Print Mid(TxtNotaDomi.Text, 1, 25) & "-"
        If Len(TxtNotaDomi.Text) > 26 Then
            Printer.Print Mid(TxtNotaDomi.Text, 26, 50) & "-"
        End If
        If Len(TxtNotaDomi.Text) > 51 Then
            Printer.Print Mid(TxtNotaDomi.Text, 51, 75) & "-"
        End If
        If Len(TxtNotaDomi.Text) > 76 Then
            Printer.Print Mid(TxtNotaDomi.Text, 76, 100) & "-"
        End If
        Printer.Print ""
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.Print "               GRACIAS POR SU COMPRA"
        Printer.Print "           PRODUCTO 100% GARANTIZADO"
        Printer.Print "--------------------------------------------------------------------------------"
        Printer.EndDoc
    End If
    TxtNomCliente.Text = ""
    TxtDomiCleinte.Text = ""
    TxtHoraDe.Text = ""
    TxtHoraAl.Text = ""
    TxtNotaDomi.Text = ""
    TxtTelefonoDomi.Text = ""
    TxtNoArticulos.Text = ""
    CmbColonia.Text = ""
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub

Private Sub BtnNueColonia_Click()
    FrmAgrColonia.Show vbModal
End Sub
'Private Sub Form_Load()
'    Me.BtnGuardaDomi.Enabled = False
'    DTPFechaDomi.Value = Format(Date, "dd/mm/yyyy")
'    Set cnn = New ADODB.Connection
'    With cnn
'        .ConnectionString = _
'            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
'            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
Private Sub CmbColonia_LostFocus()
    CmbColonia.BackColor = &H80000005
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub Image9_Click()
    Unload Me
End Sub

Private Sub TxtDomiCleinte_LostFocus()
    TxtDomiCleinte.BackColor = &H80000005
End Sub
Private Sub TxtNoArticulos_LostFocus()
    TxtNoArticulos.BackColor = &H80000005
End Sub
'        .Open
'    End With
'End Sub
Private Sub TxtNomCliente_Change()
On Error GoTo ManejaError
    If TxtNomCliente.Text <> "" And TxtDomiCleinte.Text <> "" And CmbColonia.Text <> "" And TxtTelefonoDomi.Text <> "" And TxtNoArticulos.Text <> "" Then
        Me.BtnGuardaDomi.Enabled = True
    Else
        Me.BtnGuardaDomi.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub TxtNomCliente_GotFocus()
    Me.TxtNomCliente.BackColor = &HFFE1E1
    TxtNomCliente.SelStart = 0
    TxtNomCliente.SelLength = Len(TxtNomCliente.Text)
End Sub
Private Sub TxtNomCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtDomiCleinte.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub TxtDomiCleinte_Change()
On Error GoTo ManejaError
    If TxtNomCliente.Text <> "" And TxtDomiCleinte.Text <> "" And CmbColonia.Text <> "" And TxtTelefonoDomi.Text <> "" And TxtNoArticulos.Text <> "" Then
        Me.BtnGuardaDomi.Enabled = True
    Else
        Me.BtnGuardaDomi.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub TxtDomiCleinte_GotFocus()
    Me.TxtDomiCleinte.BackColor = &HFFE1E1
    TxtDomiCleinte.SelStart = 0
    TxtDomiCleinte.SelLength = Len(TxtDomiCleinte.Text)
End Sub
Private Sub TxtDomiCleinte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbColonia.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub CmbColonia_GotFocus()
    Me.CmbColonia.BackColor = &HFFE1E1
    CmbColonia.SelStart = 0
    CmbColonia.SelLength = Len(CmbColonia.Text)
End Sub
Private Sub TxtHoraDe_GotFocus()
    TxtHoraDe.BackColor = &HFFE1E1
    TxtHoraDe.SelStart = 0
    TxtHoraDe.SelLength = Len(TxtHoraDe.Text)
End Sub
Private Sub TxtHoraDe_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        TxtHoraAl.SetFocus
    End If
    If KeyAscii = 58 And Len(TxtHoraDe.Text) = 1 Then
        TxtHoraDe.Text = "0" & TxtHoraDe.Text
        TxtHoraDe.SelStart = Len(TxtHoraDe.Text)
    End If
    If KeyAscii <> 8 Then
        If Len(TxtHoraDe.Text) = 1 And Val(TxtHoraDe.Text) > 2 Then
                TxtHoraDe.Text = "0" & TxtHoraDe.Text
                TxtHoraDe.SelStart = Len(TxtHoraDe.Text)
        Else
            If Len(TxtHoraDe.Text) = 2 Then
                TxtHoraDe.Text = TxtHoraDe.Text & ":"
                TxtHoraDe.SelStart = Len(TxtHoraDe.Text)
            End If
        End If
    End If
    Dim Valido As String
    If Len(TxtHoraDe.Text) = 1 Then
        If TxtHoraDe.Text = "2" Then
            Valido = "12340"
        Else
            Valido = "1234567890"
        End If
    End If
    If Len(TxtHoraDe.Text) = 3 Then
        Valido = "123450"
    End If
    If Len(TxtHoraDe.Text) = 4 Or Len(TxtHoraDe.Text) = 0 Then
        Valido = "1234567890"
    End If
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
Private Sub TxtHoraDe_LostFocus()
On Error GoTo ManejaError
    TxtHoraDe.BackColor = &H80000005
    If Len(TxtHoraDe.Text) = 1 Then
        TxtHoraDe.Text = "0" & TxtHoraDe.Text & ":00"
    End If
    If Len(TxtHoraDe.Text) = 2 Then
        TxtHoraDe.Text = TxtHoraDe.Text & ":00"
    End If
    If Len(TxtHoraDe.Text) = 3 Then
        TxtHoraDe.Text = TxtHoraDe.Text & "00"
    End If
    If Len(TxtHoraDe.Text) = 4 Then
        TxtHoraDe.Text = TxtHoraDe.Text & "0"
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub TxtHoraAl_GotFocus()
    TxtHoraAl.BackColor = &HFFE1E1
    TxtHoraAl.SelStart = 0
    TxtHoraAl.SelLength = Len(TxtHoraAl.Text)
End Sub
Private Sub TxtHoraAl_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        TxtNotaDomi.SetFocus
    End If
    If KeyAscii = 58 And Len(TxtHoraAl.Text) = 1 Then
        TxtHoraAl.Text = "0" & TxtHoraAl.Text
        TxtHoraAl.SelStart = Len(TxtHoraAl.Text)
    End If
    If KeyAscii <> 8 Then
        If Len(TxtHoraAl.Text) = 1 And Val(TxtHoraAl.Text) > 2 Then
                TxtHoraAl.Text = "0" & TxtHoraAl.Text
                TxtHoraAl.SelStart = Len(TxtHoraAl.Text)
        Else
            If Len(TxtHoraAl.Text) = 2 Then
                TxtHoraAl.Text = TxtHoraAl.Text & ":"
                TxtHoraAl.SelStart = Len(TxtHoraAl.Text)
            End If
        End If
    End If
    Dim Valido As String
    If Len(TxtHoraAl.Text) = 1 Then
        If TxtHoraAl.Text = "2" Then
            Valido = "12340"
        Else
            Valido = "1234567890"
        End If
    End If
    If Len(TxtHoraAl.Text) = 3 Then
        Valido = "123450"
    End If
    If Len(TxtHoraAl.Text) = 4 Or Len(TxtHoraAl.Text) = 0 Then
        Valido = "1234567890"
    End If
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
Private Sub TxtHoraAl_LostFocus()
On Error GoTo ManejaError
    TxtHoraAl.BackColor = &H80000005
    If Len(TxtHoraAl.Text) = 1 Then
        TxtHoraAl.Text = "0" & TxtHoraAl.Text & ":00"
    End If
    If Len(TxtHoraAl.Text) = 2 Then
        TxtHoraAl.Text = TxtHoraAl.Text & ":00"
    End If
    If Len(TxtHoraAl.Text) = 3 Then
        TxtHoraAl.Text = TxtHoraAl.Text & "00"
    End If
    If Len(TxtHoraAl.Text) = 4 Then
        TxtHoraAl.Text = TxtHoraAl.Text & "0"
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub TxtNomCliente_LostFocus()
    TxtNomCliente.BackColor = &H80000005
End Sub
Private Sub TxtNotaDomi_GotFocus()
    TxtNotaDomi.BackColor = &HFFE1E1
    TxtNotaDomi.SetFocus
    TxtNotaDomi.SelStart = 0
    TxtNotaDomi.SelLength = Len(TxtNotaDomi.Text)
End Sub
Private Sub TxtNotaDomi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.BtnGuardaDomi.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub TxtNotaDomi_LostFocus()
    TxtNotaDomi.BackColor = &H80000005
End Sub
Private Sub TxtTelefonoDomi_Change()
On Error GoTo ManejaError
    If TxtNomCliente.Text <> "" And TxtDomiCleinte.Text <> "" And CmbColonia.Text <> "" And TxtTelefonoDomi.Text <> "" And TxtNoArticulos.Text <> "" Then
        Me.BtnGuardaDomi.Enabled = True
    Else
        Me.BtnGuardaDomi.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub TxtTelefonoDomi_GotFocus()
    Me.TxtTelefonoDomi.BackColor = &HFFE1E1
    TxtTelefonoDomi.SelStart = 0
    TxtTelefonoDomi.SelLength = Len(TxtTelefonoDomi.Text)
End Sub
Private Sub TxtTelefonoDomi_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 47 And Len(TxtTelefonoDomi.Text) = 1 Then
        TxtTelefonoDomi.Text = "0" & TxtTelefonoDomi.Text
        TxtTelefonoDomi.SelStart = Len(TxtTelefonoDomi.Text)
    End If
    If KeyAscii <> 8 Then
        If Len(TxtTelefonoDomi.Text) = 3 Or Len(TxtTelefonoDomi.Text) = 6 Then
            TxtTelefonoDomi.Text = TxtTelefonoDomi.Text & "-"
            TxtTelefonoDomi.SelStart = Len(TxtTelefonoDomi.Text)
        End If
    End If
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then
        TxtNoArticulos.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub TxtNoArticulos_Change()
On Error GoTo ManejaError
    If TxtNomCliente.Text <> "" And TxtDomiCleinte.Text <> "" And CmbColonia.Text <> "" And TxtTelefonoDomi.Text <> "" And TxtNoArticulos.Text <> "" Then
        Me.BtnGuardaDomi.Enabled = True
    Else
        Me.BtnGuardaDomi.Enabled = False
    End If
    If TxtNoArticulos.Text = "" Then
        TxtNoArticulos.Text = 0
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub TxtNoArticulos_GotFocus()
    TxtNoArticulos.BackColor = &HFFE1E1
    TxtNoArticulos.SelStart = 0
    TxtNoArticulos.SelLength = Len(TxtNoArticulos.Text)
End Sub
Private Sub TxtNoArticulos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtHoraDe.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
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
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
        MsgBox "LA CONEXION SE RESABLECIO CON EXITO. PUEDE CONTINUAR CON SU TRABAJO.", vbInformation, "SACC"
    End With
Exit Sub
ManejaError:
    MsgBox Err.Number & Err.Description
    Err.Clear
    If MsgBox("NO PUDIMOS RESTABLECER LA CONEXIÓN, ¿DESEA REINTENTARLO?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
End Sub
Private Sub TxtTelefonoDomi_LostFocus()
    TxtTelefonoDomi.BackColor = &H80000005
End Sub
