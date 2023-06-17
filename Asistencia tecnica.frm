VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form AsisTec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capturar Asistencia Tecnica"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10350
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   9120
      ScaleHeight     =   6075
      ScaleWidth      =   1275
      TabIndex        =   29
      Top             =   0
      Width           =   1335
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   30
         Top             =   4800
         Width           =   975
         Begin VB.Label Label1 
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
            TabIndex        =   31
            Top             =   960
            Width           =   975
         End
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "Asistencia tecnica.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "Asistencia tecnica.frx":030A
            Top             =   240
            Width           =   705
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cliente Nuevo"
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
      Picture         =   "Asistencia tecnica.frx":1DBC
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox TxtTipoArt 
      DataField       =   "TIPO_ARTICULO"
      DataSource      =   "Adodc1"
      Height          =   1005
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos del Cliente"
      Height          =   2295
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   5175
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
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
   End
   Begin VB.TextBox TxtComTec 
      DataField       =   "COMENTARIOS_TECNICOS"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   6120
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   3975
      Begin VB.CheckBox Chk2 
         Caption         =   "Garantia"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox Chk1 
         Caption         =   "A domicilio"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label LblMenu2 
         Caption         =   "Label12"
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
         Left            =   960
         TabIndex        =   26
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por Nombre"
      Height          =   255
      Left            =   6120
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Clave"
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TxtDesPiez 
      DataField       =   "COMENTARIOS_COTIZACION"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   3120
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.CommandButton btnBuscar 
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
      Left            =   7440
      Picture         =   "Asistencia tecnica.frx":478E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Articulo"
      Height          =   1215
      Left            =   4200
      TabIndex        =   11
      Top             =   3000
      Width           =   4815
      Begin VB.TextBox TxtMarcaArt 
         Height          =   285
         Left            =   840
         TabIndex        =   25
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox TxtModelo 
         Height          =   285
         Left            =   840
         TabIndex        =   24
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Modelo"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdRegis 
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
      Left            =   3840
      Picture         =   "Asistencia tecnica.frx":7160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DtPFechAsi 
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   49938433
      CurrentDate     =   38678
   End
   Begin VB.Label LblMenu 
      Alignment       =   2  'Center
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   27
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de articulo"
      Height          =   195
      Left            =   960
      TabIndex        =   22
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Buscar Cliente"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   240
      Width           =   1020
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Comentarios para los Tecnicos"
      Height          =   195
      Left            =   6480
      TabIndex        =   17
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Atendió"
      Height          =   195
      Left            =   5400
      TabIndex        =   16
      Top             =   2160
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha en que se debe realizar el domicilio"
      Height          =   195
      Index           =   0
      Left            =   5400
      TabIndex        =   15
      Top             =   1320
      Width           =   2955
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Descipcion de piezas recibidas"
      Height          =   195
      Left            =   3360
      TabIndex        =   14
      Top             =   4320
      Width           =   2190
   End
End
Attribute VB_Name = "AsisTec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Dim DOMI As String
Dim GTIA As String
Dim IdClien As String
Dim NoAsTec As String
Dim NomClien As String
Dim TelCasa As String
Dim TelTrabajo As String
Dim Direc As String
Dim NoExte As String
Dim NoInte As String
Dim COLONIA As String
Private Sub btnBuscar_Click()
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
Private Sub Chk1_Click()
On Error GoTo ManejaError
    If Chk1.Value = 1 Then
        DOMI = "1"
    Else
        DOMI = "0"
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Chk1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Chk1.Value = 1
        Chk2.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Chk2_Click()
On Error GoTo ManejaError
    If Chk2.Value = 1 Then
        GTIA = "1"
    Else
        GTIA = "0"
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Chk2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Chk2.Value = 1
        TxtModelo.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cmdRegis_Click()
On Error GoTo ManejaError
    Dim sqlComanda As String
    sqlComanda = "INSERT INTO ASISTENCIA_TECNICA (SUCURSAL, GARANTIA, DESCRIPCION_PIEZAS, FECHA_DEBE_ATENDER, A_DOMICILIO, ATENDIDO, ID_USUARIO, ID_CLIENTE, FECHA_CAPTURA, TIPO_ARTICULO, MODELO, MARCA, COMENTARIOS_TECNICOS) VALUES ('" & LblMenu2.Caption & "', '" & GTIA & "', '" & TxtDesPiez.Text & "', '" & DtPFechAsi.Value & "', '" & DOMI & "', 0, '" & Menu.Text1(0).Text & "', '" & IdClien & "', '" & Date & "', '" & TxtTipoArt.Text & "', '" & TxtModelo.Text & "', '" & TxtMarcaArt.Text & "', '" & TxtComTec.Text & "');"
    cnn.Execute (sqlComanda)
    Dim tRs As Recordset
    sqlComanda = "SELECT ID_AS_TEC FROM ASISTENCIA_TECNICA ORDER BY ID_AS_TEC DESC"
    Set tRs = cnn.Execute(sqlComanda)
    NoAsTec = tRs.Fields("ID_AS_TEC")
    Imprimir
    Me.DtPFechAsi.Value = Date
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
Private Sub Command1_Click()
On Error GoTo ManejaError
    AltaClien.Show vbModal
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub

Private Sub DtPFechAsi_GotFocus()
    DtPFechAsi.BackColor = &HFFE1E1
End Sub
Private Sub DtPFechAsi_LostFocus()
    DtPFechAsi.BackColor = &H80000005
End Sub

Private Sub Form_Load()
On Error GoTo ManejaError
    LblMenu.Caption = Menu.Text1(1).Text
    LblMenu2.Caption = Menu.Text4(0).Text
    DtPFechAsi.Value = Date
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
        .ColumnHeaders.Add , , "Clave del Cliente", 0
        .ColumnHeaders.Add , , "Nombre", 2700
        .ColumnHeaders.Add , , "Telefono de casa", 2000
        .ColumnHeaders.Add , , "Telefono de oficina", 2000
        .ColumnHeaders.Add , , "Direccion", 2000
        .ColumnHeaders.Add , , "No. Exterior", 1500
        .ColumnHeaders.Add , , "No. Interior", 1500
        .ColumnHeaders.Add , , "Colonia", 2000
        .ColumnHeaders.Add , , "Ciudad", 2000
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
    If Text2.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As Recordset
        Dim tLi As ListItem
        If Option1.Value Then
            sBuscar = "SELECT ID_CLIENTE, NOMBRE, TELEFONO_CASA, TELEFONO_TRABAJO, DIRECCION, CIUDAD, COLONIA, NUMERO_EXTERIOR, NUMERO_INTERIOR FROM CLIENTE WHERE ID_CLIENTE LIKE " & Text2.Text
        Else
            sBuscar = "SELECT ID_CLIENTE, NOMBRE, TELEFONO_CASA, TELEFONO_TRABAJO, DIRECCION, CIUDAD, COLONIA, NUMERO_EXTERIOR, NUMERO_INTERIOR FROM CLIENTE WHERE NOMBRE LIKE '%" & Text2.Text & "%'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If Not (.BOF And .EOF) Then
                ListView1.ListItems.Clear
                .MoveFirst
                Do While Not .EOF
                    Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                    If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & ""
                    If Not IsNull(.Fields("TELEFONO_CASA")) Then tLi.SubItems(2) = .Fields("TELEFONO_CASA") & ""
                    If Not IsNull(.Fields("TELEFONO_TRABAJO")) Then tLi.SubItems(3) = .Fields("TELEFONO_TRABAJO") & ""
                    If Not IsNull(.Fields("DIRECCION")) Then tLi.SubItems(4) = .Fields("DIRECCION") & ""
                    If Not IsNull(.Fields("NUMERO_EXTERIOR")) Then tLi.SubItems(5) = .Fields("NUMERO_EXTERIOR") & ""
                    If Not IsNull(.Fields("NUMERO_INTERIOR")) Then tLi.SubItems(6) = .Fields("NUMERO_INTERIOR") & ""
                    If Not IsNull(.Fields("COLONIA")) Then tLi.SubItems(7) = .Fields("COLONIA") & ""
                    If Not IsNull(.Fields("CIUDAD")) Then tLi.SubItems(8) = .Fields("CIUDAD") & ""
                    .MoveNext
                Loop
            End If
        End With
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
    IdClien = Item
    Text2.Text = Item.SubItems(1)
    Option2.Value = True
    NomClien = Item.SubItems(1)
    TelCasa = Item.SubItems(2)
    TelTrabajo = Item.SubItems(3)
    Direc = Item.SubItems(4)
    NoExte = Item.SubItems(5)
    NoInte = Item.SubItems(6)
    COLONIA = Item.SubItems(7)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Chk1.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Option1_Click()
On Error GoTo ManejaError
    Text2.Text = ""
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
'Private Sub Text1_Change(Index As Integer)
'On Error GoTo ManejaError
    'If Index = 12 Or Index = 6 Then
        'If Text1(12).Text = "1" Then
            'Chk1.Value = 1
        'Else
            'Chk1.Value = 0
        'End If
        'If Text1(6).Text = "1" Then
            'Chk2.Value = 1
        'Else
            'Chk2.Value = 0
        'End If
    'End If
'Exit Sub
'ManejaError:
        'MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        'Err.Clear
'End Sub
'Private Sub Text1_GotFocus(Index As Integer)
'On Error GoTo ManejaError
    'Text1(Index).SetFocus
    'Text1(Index).SelStart = 0
    'Text1(Index).SelLength = Len(Text1(Index).Text)
'Exit Sub
'ManejaError:
        'MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        'Err.Clear
'End Sub
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'On Error GoTo ManejaError
    'Dim Valido As String
    'If Index = 3 Then
        'Valido = "1234567890-()"
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'If KeyAscii > 26 Then
            'If InStr(Valido, Chr(KeyAscii)) = 0 Then
                'KeyAscii = 0
            'End If
        'End If
    'End If
    'If Index = 7 Or Index = 8 Or Index = 11 Then
        'Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))
        'If KeyAscii > 26 Then
            'If InStr(Valido, Chr(KeyAscii)) = 0 Then
                'KeyAscii = 0
            'End If
        'End If
    'End If
'Exit Sub
'ManejaError:
        'MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        'Err.Clear
'End Sub
Private Sub Imprimir()
On Error GoTo ManejaError
    Printer.Print "   ACTITUD POSITIVA EN TONER S DE RL MI"
    Printer.Print "                    R.F.C. APT- 040201-KA5"
    Printer.Print "ORTIZ DE CAMPOS No. 1308 COL. SAN FELIPE"
    Printer.Print "      CHIHUAHUA, CHIHUAHUA C.P. 31203"
    Printer.Print "FECHA : " & Date
    Printer.Print "SUCURSAL : " & LblMenu2.Caption
    Printer.Print "No. DE ASISTENCIA : " & NoAsTec
    Printer.Print "ATENDIDO POR : " & LblMenu.Caption
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                       ASISTENCIA TECNICA"
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "Cliente : " & NomClien
    Printer.Print "Telefono : " & TelCasa & " o " & TelTrabajo
    Printer.Print "Calle : " & Direc & " # " & NoExte & "-" & NoInte
    Printer.Print "Colonia : " & COLONIA
    Printer.Print "Fecha a atender : " & DtPFechAsi.Value
    Printer.Print ""
    Printer.Print "Marca : " & TxtModelo.Text
    Printer.Print "Modelo : " & TxtMarcaArt.Text
    Printer.Print "Decripcion : " & TxtDesPiez.Text
    Printer.Print "Comentarios : " & TxtComTec.Text
    Printer.Print "Articulo : " & TxtTipoArt.Text
    Printer.Print ""
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "               GRACIAS POR SU COMPRA"
    Printer.Print "           PRODUCTO 100% GARANTIZADO"
    Printer.Print "LA GARANTÍA SERÁ VÁLIDA HASTA 15 DIAS"
    Printer.Print "     DESPUES DE HABER EFECTUADO SU "
    Printer.Print "                                COMPRA"
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.EndDoc
    Text2.Text = ""
    TxtModelo.Text = ""
    TxtMarcaArt.Text = ""
    TxtDesPiez.Text = ""
    TxtTipoArt.Text = ""
    TxtComTec.Text = ""
    Chk1.Value = 0
    Chk2.Value = 0
    'ListView1.ListItems.Clear
    MsgBox "                        ASITENCIA REGISTRADA!                        "
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text2_GotFocus()
On Error GoTo ManejaError
    Text2.BackColor = &HFFE1E1
    Text2.SetFocus
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Buscar
        'Me.ListView1.SetFocus
    End If
    Dim Valido As String
    If Option1.Value = True Then
        Valido = "1234567890"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    Else
        Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub

Private Sub TxtComTec_GotFocus()
    TxtComTec.BackColor = &HFFE1E1
End Sub
Private Sub TxtComTec_LostFocus()
    TxtComTec.BackColor = &H80000005
End Sub
Private Sub TxtDesPiez_GotFocus()
    TxtDesPiez.BackColor = &HFFE1E1
End Sub
Private Sub TxtDesPiez_LostFocus()
    TxtDesPiez.BackColor = &H80000005
End Sub
Private Sub TxtModelo_GotFocus()
On Error GoTo ManejaError
    TxtModelo.BackColor = &HFFE1E1
    TxtModelo.SetFocus
    TxtModelo.SelStart = 0
    TxtModelo.SelLength = Len(TxtModelo.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub TxtModelo_LostFocus()
    TxtModelo.BackColor = &H80000005
End Sub
Private Sub TxtModelo_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        TxtMarcaArt.SetFocus
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
Private Sub TxtMarcaArt_GotFocus()
On Error GoTo ManejaError
    TxtMarcaArt.BackColor = &HFFE1E1
    TxtMarcaArt.SetFocus
    TxtMarcaArt.SelStart = 0
    TxtMarcaArt.SelLength = Len(TxtMarcaArt.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub TxtMarcaArt_LostFocus()
    TxtMarcaArt.BackColor = &H80000005
End Sub
Private Sub TxtMarcaArt_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        TxtDesPiez.SetFocus
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
        'If BanCnn = False Then
            '.Close
            'BanCnn = True
        'End If
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
        'BanCnn = False
        MsgBox "LA CONEXION SE RESABLECIO CON EXITO. PUEDE CONTINUAR CON SU TRABAJO.", vbInformation, "MENSAJE DEL SISTEMA"
    End With
Exit Sub
ManejaError:
    MsgBox Err.Number & Err.Description
    Err.Clear
    If MsgBox("NO PUDIMOS RESTABLECER LA CONEXIÓN, ¿DESEA REINTENTARLO?", vbYesNo + vbCritical + vbDefaultButton1, "MENSAJE DEL SISTEMA") = vbYes Then Reconexion
End Sub
Private Sub TxtTipoArt_GotFocus()
    TxtTipoArt.BackColor = &HFFE1E1
End Sub
Private Sub TxtTipoArt_LostFocus()
    TxtTipoArt.BackColor = &H80000005
End Sub
