VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCamClienVent 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar Cliente a Venta"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6240
      TabIndex        =   10
      Top             =   2040
      Width           =   975
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Modificar"
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image7 
         Height          =   810
         Left            =   120
         MouseIcon       =   "FrmCamClienVent.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCamClienVent.frx":030A
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6240
      TabIndex        =   7
      Top             =   3240
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCamClienVent.frx":2434
         MousePointer    =   99  'Custom
         Picture         =   "FrmCamClienVent.frx":273E
         Top             =   120
         Width           =   720
      End
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
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Cambio"
      TabPicture(0)   =   "FrmCamClienVent.frx":4820
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   1680
         Width           =   4935
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3836
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame1 
         Caption         =   "Numero de Venta"
         Height          =   1095
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3495
         Begin VB.OptionButton Option2 
            Caption         =   "a Comanda"
            Height          =   255
            Left            =   2040
            TabIndex        =   2
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "a Venta"
            Height          =   195
            Left            =   2040
            TabIndex        =   1
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   0
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmCamClienVent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Attribute cnn.VB_VarHelpID = -1
Dim ClvCliente As String
Dim NomCli As String
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No. Cliente", 1000
        .ColumnHeaders.Add , , "Nombre", 5700
    End With
End Sub
Private Sub Image7_Click()
    If NomCli <> "" And ClvCliente <> "" And Text1.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        If Option1.Value = True Then
            sBuscar = "SELECT ID_VENTA FROM VENTAS WHERE ID_VENTA = " & Text1.Text & " AND FACTURADO = 0"
            Set tRs = cnn.Execute(sBuscar)
            If tRs.EOF And tRs.BOF Then
                MsgBox "LA VENTA NO EXISTE O YA FUE FACTURADA!", vbInformation, "SACC"
            Else
                If MsgBox("ESTA SEGURO QUE DESEA HACER EL CAMBIO DE CLIENTE A LA VENTA " & Text1.Text, vbYesNo + vbInformation + vbDefaultButton1, "SACC") = vbYes Then
                    sBuscar = "UPDATE VENTAS SET ID_CLIENTE = " & ClvCliente & ", NOMBRE = '" & NomCli & "' WHERE ID_VENTA = " & Text1.Text
                    cnn.Execute (sBuscar)
                    sBuscar = "SELECT ID_CUENTA FROM CUENTA_VENTA WHERE ID_VENTA = " & Text1.Text
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        sBuscar = "UPDATE CUENTAS SET ID_CLIENTE = " & ClvCliente & " WHERE ID_CUENTA = " & tRs.Fields("ID_CUENTA")
                        cnn.Execute (sBuscar)
                    End If
                    ClvCliente = ""
                    NomCli = ""
                    Text1.Text = ""
                Else
                    MsgBox "NO SE REALIZO EL CAMBIO!", vbInformation, "SACC"
                End If
            End If
        Else
            sBuscar = "SELECT ID_COMANDA FROM COMANDAS_2 WHERE ID_COMANDA = " & Text1.Text
            Set tRs = cnn.Execute(sBuscar)
            If tRs.EOF And tRs.BOF Then
                MsgBox "LA COMANDA NO EXISTE!", vbInformation, "SACC"
            Else
                If MsgBox("ESTA SEGURO QUE DESEA HACER EL CAMBIO DE CLIENTE A LA COMANDA " & Text1.Text, vbYesNo + vbInformation + vbDefaultButton1, "SACC") = vbYes Then
                    sBuscar = "UPDATE COMANDAS_2 SET ID_CLIENTE = " & ClvCliente & " WHERE ID_COMANDA = " & Text1.Text
                    cnn.Execute (sBuscar)
                    ClvCliente = ""
                    NomCli = ""
                    Text1.Text = ""
                Else
                    MsgBox "NO SE REALIZO EL CAMBIO!", vbInformation, "SACC"
                End If
            End If
        End If
    Else
        MsgBox "FALTA INFORMACIÓN NECESARIA!", vbInformation, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ClvCliente = Item
    NomCli = Item.SubItems(1)
    Text2.Text = Item.SubItems(1)
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
    End If
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
    Text2.SetFocus
    Text2.SelStart = 0
    Text2.SelLength = Len(Text1.Text)
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        ListView1.ListItems.Clear
        sBuscar = "SELECT NOMBRE, ID_CLIENTE FROM CLIENTE WHERE NOMBRE LIKE '%" & Text2.Text & "%' AND VALORACION = 'A'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tRs.MoveNext
            Loop
            ListView1.SetFocus
        End If
    End If
End Sub
