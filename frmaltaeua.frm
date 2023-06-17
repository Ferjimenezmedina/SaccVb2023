VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmaltaeua 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Bodegas de envío"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   2
      Top             =   1680
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmaltaeua.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmaltaeua.frx":030A
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   0
      Top             =   2880
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmaltaeua.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "frmaltaeua.frx":1FD6
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Datos Personales"
      TabPicture(0)   =   "frmaltaeua.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListView1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text7"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Option1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Option2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      Begin VB.OptionButton Option2 
         Caption         =   "Internacional"
         Height          =   255
         Left            =   6360
         TabIndex        =   22
         Top             =   2400
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nacional"
         Height          =   255
         Left            =   6360
         TabIndex        =   21
         Top             =   2160
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   195
         Left            =   3000
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   5880
         TabIndex        =   19
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3360
         TabIndex        =   18
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3360
         TabIndex        =   17
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3360
         TabIndex        =   16
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3360
         TabIndex        =   15
         Top             =   1440
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Desactivar"
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
         Left            =   1680
         Picture         =   "frmaltaeua.frx":40D4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Activar"
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
         Left            =   480
         Picture         =   "frmaltaeua.frx":6AA6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3135
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4683
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* Nombre "
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         Height          =   195
         Left            =   3360
         TabIndex        =   11
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "* Nombre "
         Height          =   195
         Left            =   3360
         TabIndex        =   10
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
         Height          =   195
         Left            =   5880
         TabIndex        =   9
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Telefono Trabajo"
         Height          =   195
         Left            =   3360
         TabIndex        =   8
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad :"
         Height          =   195
         Left            =   3360
         TabIndex        =   7
         Top             =   1920
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmaltaeua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim ValDes As String
Dim Guar As String
Private Sub Command1_Click()
    Dim sBuscar As String
    sBuscar = "UPDATE DIREIMPOR SET STATUS ='F' WHERE ID = '" & Text7.Text & "'"
    cnn.Execute (sBuscar)
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    sBuscar = "UPDATE DIREIMPOR SET STATUS ='A' WHERE ID = '" & Text7.Text & "'"
    cnn.Execute (sBuscar)
End Sub
Private Sub Image8_Click()
    Dim sBuscar As String
    Dim sTipo As String
On Error GoTo ManejaError
    If Option1.Value = True Then
        sTipo = "N"
    Else
        sTipo = "I"
    End If
    If Text4.Text = "" Then
        Text4.Text = 0
    End If
    If Text3.Text = "" Then
        Text3.Text = "0"
    End If
    If Text5.Text = "" Then
        Text5.Text = "0"
    End If
    If Text1.Text = "" Then
        Text1.Text = "0"
    End If
    If Text6.Text = "" Then
        Text6.Text = "0"
    End If
    If Text4.Text <> "" Then
        sBuscar = "INSERT INTO DIREIMPOR (NOMBRE, DIRECCION, CD, TEL1, TEL2, STATUS, TIPO) VALUES ('" & Text4.Text & "', '" & Text3.Text & "', '" & Text5.Text & "', '" & Text1.Text & "', '" & Text6.Text & "', 'A', '" & sTipo & "');"
        cnn.Execute (sBuscar)
        Text4.Text = ""
        Text1.Text = ""
        Text5.Text = ""
        Text6.Text = ""
        Text3.Text = ""
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
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
        .ColumnHeaders.Add , , "ID", 500
        .ColumnHeaders.Add , , "Nombre", 3000
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text2.Text = ListView1.SelectedItem.SubItems(1)
    Text2.SetFocus
    Text7.Text = Item
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    If KeyAscii = 13 Then
        sBuscar = "SELECT *  FROM DIREIMPOR WHERE NOMBRE LIKE '%" & Text2.Text & "%'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID"))
                tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
