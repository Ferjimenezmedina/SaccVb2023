VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form BuscaEntrada 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Buscar entrada del articulo"
   ClientHeight    =   6855
   ClientLeft      =   4860
   ClientTop       =   2385
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   10455
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9360
      TabIndex        =   18
      Top             =   5520
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "BuscaEntrada.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "BuscaEntrada.frx":030A
         Top             =   120
         Width           =   720
      End
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
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "BuscaEntrada.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ListView1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Option3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Option2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Option1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text6"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text5"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text4"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Option4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "DTPicker1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "DTPicker2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52297729
         CurrentDate     =   40514
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52297729
         CurrentDate     =   40514
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Marca"
         Height          =   195
         Left            =   5160
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   450
         Width           =   4935
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   5760
         Width           =   6015
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   6120
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   6120
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Codigo de barras"
         Height          =   195
         Left            =   6240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Clave"
         Height          =   195
         Left            =   5160
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   6240
         TabIndex        =   3
         Top             =   600
         Width           =   1455
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
         Left            =   7800
         Picture         =   "BuscaEntrada.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4335
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7646
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label8 
         Caption         =   "A :"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "De :"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Entrada"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   5760
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   6120
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Total"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   6120
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Numero de factura"
         Height          =   255
         Left            =   4680
         TabIndex        =   13
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "BuscaEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Command1_Click()
    Busca
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    DTPicker1.Value = Date - 30
    DTPicker2.Value = Date
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 1500
        .ColumnHeaders.Add , , "Descripcion", 3900
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "FECHA", 1200
        .ColumnHeaders.Add , , "#ENTRADA", 1200
        .ColumnHeaders.Add , , "MARCA", 1200
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim sBuscar2 As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    If Item.SubItems(4) <> "" Then
        Text2.Text = Item.SubItems(4)
        sBuscar = "SELECT ID_ENTRADA,ID_PROVEEDOR,TOTAL,FECHA,FACTURA FROM ENTRADAS WHERE ID_ENTRADA =" & Item.SubItems(4) & ""
        Set tRs = cnn.Execute(sBuscar)
        Text4.Text = tRs.Fields("FECHA")
        Text5.Text = tRs.Fields("TOTAL")
        Text6.Text = tRs.Fields("FACTURA")
        sBuscar2 = "SELECT ID_PROVEEDOR,NOMBRE FROM PROVEEDOR WHERE ID_PROVEEDOR =" & tRs.Fields("ID_PROVEEDOR") & ""
        Set tRs2 = cnn.Execute(sBuscar2)
        Text3.Text = tRs2.Fields("NOMBRE")
    End If
End Sub
Private Sub Option1_Click()
    If Option1.Value = True Then
        Text1.SetFocus
    End If
End Sub
Private Sub Option2_Click()
    If Option2.Value = True Then
        Text1.SetFocus
    End If
End Sub
Private Sub Option3_Click()
    If Option3.Value = True Then
        Text1.SetFocus
    End If
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Busca
    Else
        Dim Valido As String
        Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Public Sub Limpiar()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
End Sub
Private Sub Busca()
    On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim i As Integer
    If Option1.Value = True Then
       sBuscar = "SELECT CODIGO_BARAS,ID_ENTRADA,ID_PRODUCTO,CANTIDAD,FECHA,Descripcion, MARCA FROM vsENTRADAS WHERE CODIGO_BARAS LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY ID_ENTRADA"
    End If
    If Option2.Value = True Then
        sBuscar = "SELECT ID_ENTRADA,ID_PRODUCTO,CANTIDAD,FECHA,Descripcion, MARCA  FROM vsENTRADAS WHERE ID_PRODUCTO LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY ID_ENTRADA"
    End If
    If Option3.Value = True Then
        sBuscar = "SELECT ID_ENTRADA,ID_PRODUCTO,CANTIDAD,FECHA,Descripcion, MARCA FROM vsENTRADAS WHERE Descripcion LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY ID_ENTRADA"
    End If
    If Option4.Value = True Then
        sBuscar = "SELECT ID_ENTRADA,ID_PRODUCTO,CANTIDAD,FECHA,Descripcion, MARCA FROM vsENTRADAS WHERE MARCA LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY ID_PRODUCTO"
    End If
    If sBuscar <> "" Then
        Set tRs = cnn.Execute(sBuscar)
        ListView1.ListItems.Clear
            If Not (tRs.BOF And tRs.EOF) Then
                Do While Not (tRs.EOF)
                        Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                        If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                        If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(3) = tRs.Fields("FECHA")
                        If Not IsNull(tRs.Fields("ID_ENTRADA")) Then tLi.SubItems(4) = tRs.Fields("ID_ENTRADA")
                        If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
                        If Not IsNull(tRs.Fields("MARCA")) Then tLi.SubItems(5) = tRs.Fields("MARCA")
                        tRs.MoveNext
                Loop
            Else
                Set tLi = ListView1.ListItems.Add(, , "")
                    tLi.SubItems(1) = "NO EXISTEN RESULTADOS"
            End If
                Limpiar
                ListView1.SetFocus
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
