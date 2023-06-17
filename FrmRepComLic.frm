VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRepComLic 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Comparacion de Licitaciónes"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Estilo del Reporte"
      Height          =   1095
      Left            =   6120
      TabIndex        =   13
      Top             =   480
      Width           =   2055
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reporte Tabulado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reporte de Listado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7200
      TabIndex        =   3
      Top             =   1800
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepComLic.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepComLic.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6120
      TabIndex        =   1
      Top             =   1800
      Width           =   975
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reporte"
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image11 
         Height          =   675
         Left            =   120
         MouseIcon       =   "FrmRepComLic.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepComLic.frx":26F6
         Top             =   240
         Width           =   660
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5106
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Cliente"
      TabPicture(0)   =   "FrmRepComLic.frx":3E6C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Producto"
      TabPicture(1)   =   "FrmRepComLic.frx":3E88
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Text3"
      Tab(1).Control(2)=   "ListView2"
      Tab(1).Control(3)=   "Label3"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Competidores"
      TabPicture(2)   =   "FrmRepComLic.frx":3EA4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text5"
      Tab(2).Control(1)=   "ListView3"
      Tab(2).Control(2)=   "Label4"
      Tab(2).ControlCount=   3
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   -73920
         TabIndex        =   20
         Top             =   600
         Width           =   4695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Precio"
         Height          =   1215
         Left            =   -71280
         TabIndex        =   16
         Top             =   840
         Width           =   2055
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Mayor a :"
            Height          =   195
            Left            =   480
            TabIndex        =   18
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Menor a :"
            Height          =   255
            Left            =   480
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -74040
         TabIndex        =   10
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3720
         TabIndex        =   9
         Top             =   1440
         Width           =   2055
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1815
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   600
         Width           =   2895
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   12
         Top             =   960
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   22
         Top             =   960
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
         Caption         =   "Competidor :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "No. de Licitación :"
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmRepComLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave", 300
        .ColumnHeaders.Add , , "Cliente", 3420
    End With
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 3420
    End With
    With ListView3
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Competidor", 300
        .ColumnHeaders.Add , , "Nombre", 5560
    End With
End Sub
Private Sub Image11_Click()
    Dim sBuscar As String
    Dim VarWhere As String
    Dim cont As Integer
    Dim NoReg As Integer
    Dim RegCont As Integer
    Dim Path As String
    Path = App.Path
    sBuscar = "SELECT * FROM VsComparativoLicitacion "
    NoReg = ListView1.ListItems.Count
    For cont = 1 To NoReg
        If ListView1.ListItems.Item(cont).Checked = True Then
            If VarWhere = "" Then
                VarWhere = " WHERE (ID_CLIENTE = " & ListView1.ListItems.Item(cont)
            Else
                VarWhere = VarWhere & " OR ID_CLIENTE = " & ListView1.ListItems.Item(cont)
            End If
        End If
    Next cont
    If VarWhere <> "" And Text2.Text = "" Then
        VarWhere = VarWhere & ") "
    End If
    If Text2.Text <> "" Then
        If VarWhere = "" Then
            VarWhere = " WHERE NO_LICITACION = '" & Text2.Text & "'"
        Else
            VarWhere = VarWhere & " AND NO_LICITACION = '" & Text2.Text & "')"
        End If
    End If
    NoReg = ListView2.ListItems.Count
    For cont = 1 To NoReg
        If ListView2.ListItems.Item(cont).Checked = True Then
            If VarWhere = "" Then
                VarWhere = " WHERE (ID_PRODUCTO = '" & Trim(ListView2.ListItems.Item(cont)) & "'"
                RegCont = 1
            Else
                If RegCont = 0 Then
                    VarWhere = VarWhere & " AND (ID_PRODUCTO = '" & Trim(ListView2.ListItems.Item(cont)) & "'"
                    RegCont = 1
                Else
                    VarWhere = VarWhere & " OR ID_PRODUCTO = '" & Trim(ListView2.ListItems.Item(cont)) & "'"
                End If
            End If
        End If
    Next cont
    If RegCont = 1 And Text4.Text = "" Then
        VarWhere = VarWhere & ") "
    End If
    If Text4.Text <> "" Then
        If VarWhere = "" Then
            If Option3.Value = True Then
                VarWhere = " WHERE PRECIO < " & Text4.Text & ""
            Else
                VarWhere = " WHERE PRECIO > " & Text4.Text & ""
            End If
        Else
            If Option3.Value = True Then
                VarWhere = VarWhere & " AND PRECIO < " & Text4.Text & ""
            Else
                VarWhere = VarWhere & " AND PRECIO > " & Text4.Text & ""
            End If
        End If
        If RegCont = 1 Then
            VarWhere = VarWhere & ") "
        End If
    End If
    RegCont = 0
    NoReg = ListView3.ListItems.Count
    For cont = 1 To NoReg
        If ListView3.ListItems.Item(cont).Checked = True Then
            If VarWhere = "" Then
                VarWhere = " WHERE (ID_COMPETIDOR = " & ListView3.ListItems.Item(cont)
                RegCont = 1
            Else
                If RegCont = 0 Then
                    VarWhere = VarWhere & " AND (ID_COMPETIDOR = " & ListView3.ListItems.Item(cont)
                Else
                    VarWhere = VarWhere & " OR ID_COMPETIDOR = " & ListView3.ListItems.Item(cont)
                End If
            End If
        End If
    Next cont
    If RegCont = 1 Then
        VarWhere = VarWhere & ")"
    End If
    sBuscar = sBuscar & VarWhere
    If Option1.Value = True Then
        Set crReport = crApplication.OpenReport(Path & "\REPORTES\RepCompLicTabla.rpt")
        crReport.Database.LogOnServer "p2ssql.dll", GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "1"), GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER"), GetSetting("APTONER", "ConfigSACC", "USUARIO", "1"), GetSetting("APTONER", "ConfigSACC", "PASSWORD", "1")
        crReport.SQLQueryString = sBuscar
        crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        frmRep.Show vbModal, Me
    Else
        Set crReport = crApplication.OpenReport(Path & "\REPORTES\RepCompLic.rpt")
        crReport.Database.LogOnServer "p2ssql.dll", GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "1"), GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER"), GetSetting("APTONER", "ConfigSACC", "USUARIO", "1"), GetSetting("APTONER", "ConfigSACC", "PASSWORD", "1")
        crReport.SQLQueryString = sBuscar
        crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        frmRep.Show vbModal, Me
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As Recordset
        Dim tLi As ListItem
        sBuscar = "SELECT ID_CLIENTE, NOMBRE FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND VALORACION = 'A'"
        Set tRs = cnn.Execute(sBuscar)
        ListView1.ListItems.Clear
        If Not (tRs.EOF And tRs.EOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tRs.MoveNext
            Loop
        Else
            MsgBox "NO SE ENCONTRARON REGISTROS SIMILARES!", vbInformation, "SACC"
        End If
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As Recordset
        Dim tLi As ListItem
        sBuscar = "SELECT ID_PRODUCTO FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text3.Text & "%'"
        Set tRs = cnn.Execute(sBuscar)
        ListView2.ListItems.Clear
        If Not (tRs.EOF And tRs.EOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                tRs.MoveNext
            Loop
        Else
            MsgBox "NO SE ENCONTRARON REGISTROS SIMILARES!", vbInformation, "SACC"
        End If
    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As Recordset
        Dim tLi As ListItem
        sBuscar = "SELECT ID_COMPETIDOR, NOMBRE FROM COMPETIDOR_LICITACION WHERE NOMBRE LIKE '%" & Text5.Text & "%'"
        Set tRs = cnn.Execute(sBuscar)
        ListView3.ListItems.Clear
        If Not (tRs.EOF And tRs.EOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_COMPETIDOR"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tRs.MoveNext
            Loop
        Else
            MsgBox "NO SE ENCONTRARON REGISTROS SIMILARES!", vbInformation, "SACC"
        End If
    End If
End Sub
