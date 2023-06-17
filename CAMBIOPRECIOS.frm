VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form CambioPRe 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAMBIO DE PRECIO"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9360
      TabIndex        =   19
      Top             =   3960
      Width           =   975
      Begin VB.Label Label26 
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
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "CAMBIOPRECIOS.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "CAMBIOPRECIOS.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Cambiar"
      TabPicture(0)   =   "CAMBIOPRECIOS.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListView1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "btnCalcular"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Option1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Option2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "btnBuscar"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "btnGuardar"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   7920
         TabIndex        =   6
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton btnGuardar 
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
         Left            =   7800
         Picture         =   "CAMBIOPRECIOS.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4440
         Width           =   1215
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
         Left            =   6240
         Picture         =   "CAMBIOPRECIOS.frx":4DDA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripcion"
         Height          =   195
         Left            =   4680
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave "
         Height          =   195
         Left            =   4680
         TabIndex        =   1
         Top             =   600
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   4455
      End
      Begin VB.CommandButton btnCalcular 
         Caption         =   "Calcular"
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
         Left            =   6480
         Picture         =   "CAMBIOPRECIOS.frx":77AC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         DataField       =   "ID_DESCUENTO"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   3
         Left            =   6720
         TabIndex        =   5
         Text            =   "30"
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         DataField       =   "PRECIO_COSTO"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   2
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "500"
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         DataField       =   "DESCRIPCION"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   1
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3960
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         DataField       =   "ID_PRODUCTO"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3960
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4260
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   1320
         TabIndex        =   18
         Top             =   3720
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Clave"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   3720
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Buscar"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ganancia"
         Height          =   195
         Left            =   6720
         TabIndex        =   15
         Top             =   3720
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Precio de Venta"
         Height          =   195
         Left            =   7920
         TabIndex        =   14
         Top             =   3720
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Precio de costo"
         Height          =   195
         Left            =   5400
         TabIndex        =   13
         Top             =   3720
         Width           =   1110
      End
   End
End
Attribute VB_Name = "CambioPRe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cambio As String
Option Explicit
Private cnn As ADODB.Connection
Private Sub btnBuscar_Click()
    Buscar
End Sub
Private Sub btnCalcular_Click()
    If Cambio = "G" Then
        Text1(4).Text = Format((CDbl(Text1(2).Text) + (CDbl(Text1(3).Text) / 100) * CDbl(Text1(2).Text)), "0.00")
    End If
    If Cambio = "V" Then
        Text1(3).Text = Format(((CDbl(Text1(4).Text) * 100) / CDbl(Text1(2).Text)) - 100, "0.00")
    End If
End Sub
Private Sub btnGuardar_Click()
    If CDbl(Text1(4).Text) > CDbl(Text1(2).Text) Then
        Dim sqlComanda As String
        Dim tRs As Recordset
        Dim Gana As String
        Gana = CDbl(Text1(3).Text) / 100
        Gana = Replace(Gana, ",", ".")
        'MsgBox Gana
        sqlComanda = "UPDATE ALMACEN3 SET GANANCIA = " & Gana & " WHERE ID_PRODUCTO = '" & Text1(0).Text & "'"
        cnn.Execute (sqlComanda)
        Text1(0).Text = ""
        Text1(1).Text = ""
        Text1(2).Text = ""
        Text1(3).Text = ""
        Text1(4).Text = ""
        Buscar
    Else
        MsgBox "EL PRECIO DE VENTA NO PUEDE ESTAR POR DEBAJO DEL PRECIO DE COSTO!", vbInformation, "SACC"
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .CheckBoxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "CLAVE", 1400
        .ColumnHeaders.Add , , "DESCRIPCIÓN", 4400
        .ColumnHeaders.Add , , "PRECIO COSTO", 1000
        .ColumnHeaders.Add , , "% DE GANANCIA", 1000
        .ColumnHeaders.Add , , "PRECIO DE VENTA", 1000
    End With
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1(0).Text = Item
    Text1(1).Text = Item.SubItems(1)
    Text1(2).Text = Item.SubItems(2)
    Text1(3).Text = CDbl(Item.SubItems(3)) * 100
    Text1(4).Text = Item.SubItems(4)
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        Else
            If Index = 3 Then
                Cambio = "G"
            End If
            If Index = 4 Then
                Cambio = "V"
            End If
        End If
    End If
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Buscar
    End If
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    If Index = 3 Or Index = 4 Then
        Text1(Index).BackColor = &HFFE1E1
    End If
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    If Index = 3 Or Index = 4 Then
        Text1(Index).BackColor = &H80000005
    End If
End Sub
Private Sub Buscar()
    Dim tLi As ListItem
    Dim tRs As Recordset
    Dim sBus As String
    Dim Remp As String
    Dim Remp2 As String
    If Option1.Value = True Then
        sBus = "SELECT * FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%'"
    Else
        sBus = "SELECT * FROM ALMACEN3 WHERE DESCRIPCION LIKE '%" & Text2.Text & "%'"
    End If
    Set tRs = cnn.Execute(sBus)
    With tRs
        ListView1.ListItems.Clear
        Do While Not .EOF
            Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                tLi.SubItems(1) = .Fields("DESCRIPCION") & ""
                If Not IsNull(.Fields("PRECIO_COSTO")) Then
                    Remp = Replace(.Fields("PRECIO_COSTO"), ",", ".")
                    tLi.SubItems(2) = Remp
                End If
                If Not IsNull(.Fields("GANANCIA")) Then
                    Remp2 = Replace(.Fields("GANANCIA"), ",", ".")
                    tLi.SubItems(3) = CDbl(Remp2) '* 100
                End If
                If Not IsNull(.Fields("GANANCIA")) Or Not IsNull(.Fields("PRECIO_COSTO")) Then
                    tLi.SubItems(4) = Format(((CDbl(Remp) * CDbl(Remp2))) + CDbl(Remp), "0.00")
                End If
                .MoveNext
        Loop
    End With
End Sub
