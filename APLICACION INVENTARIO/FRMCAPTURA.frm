VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMCAPTURA 
   Caption         =   "CAPTURA INVENTARIO"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   Icon            =   "FRMCAPTURA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8160
      TabIndex        =   12
      Top             =   120
      Width           =   2295
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2535
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   4471
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.OptionButton Option2 
      Caption         =   "POR DESCRIPCION"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   360
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "POR CLAVE"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GUARDAR"
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5760
      TabIndex        =   4
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4680
      Width           =   3615
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6800
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label5 
      Caption         =   "SUCURSAL"
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "BUSCAR"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "CANTIDAD"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "CLAVE"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   4680
      Width           =   615
   End
End
Attribute VB_Name = "FRMCAPTURA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Combo1_DropDown()
    Dim sBuscar As String
    Dim tRs As Recordset
    sBuscar = "SELECT NOMBRE FROM SUCURSALES ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    Combo1.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command1_Click()
    GUARDAEXISTENCIA
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Const sPathBase As String = "LINUX"
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "CLAVE", 3400
        .ColumnHeaders.Add , , "DESCRIPCIÓN", 7400
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "CLAVE", 3400
        .ColumnHeaders.Add , , "CANTIDAD", 2400
    End With
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = Item
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        GUARDAEXISTENCIA
        Text3.SetFocus
    End If
    Dim Valido As String
    Valido = "1234567890.-"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text3_GotFocus()
    Me.Text3.SelStart = 0
    Me.Text3.SelLength = Len(Me.Text3.Text)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As Recordset
        Dim TLI As ListItem
        If Option1.Value = True Then
            sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text3.Text & "%' ORDER BY ID_PRODUCTO"
        Else
            sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text3.Text & "%' ORDER BY ID_PRODUCTO"
        End If
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            ListView1.ListItems.Clear
            If Not (.BOF And .EOF) Then
                Do While Not .EOF
                    Set TLI = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    If Not IsNull(.Fields("DESCRIPCION")) Then TLI.SubItems(1) = .Fields("DESCRIPCION") & ""
                    .MoveNext
                Loop
            End If
        End With
    End If
    If KeyAscii = 13 Then
        'Dim sBuscar As String
        'Dim tRs As Recordset
        'Dim TLI As ListItem
        sBuscar = "SELECT ID_PRODUCTO, CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO LIKE '%" & Text3.Text & "%' AND SUCURSAL = '" & Combo1.Text & "' ORDER BY ID_PRODUCTO"
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            ListView2.ListItems.Clear
            If Not (.BOF And .EOF) Then
                Do While Not .EOF
                    Set TLI = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    If Not IsNull(.Fields("CANTIDAD")) Then TLI.SubItems(1) = .Fields("CANTIDAD") & ""
                    .MoveNext
                Loop
            End If
        End With
    End If
    If ListView1.ListItems.Count > 0 And KeyAscii = 13 Then
        ListView1.SetFocus
    End If
End Sub
Private Sub GUARDAEXISTENCIA()
    If Text2.Text <> "" And Text1.Text <> "" And Combo1.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As Recordset
        Text2.Text = Replace(Text2.Text, ",", ".")
        sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Text1.Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If tRs.EOF And tRs.BOF Then
            sBuscar = "INSERT INTO EXISTENCIAS(ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & Text1.Text & "', " & Text2.Text & ", '" & Combo1.Text & "');"
            cnn.Execute (sBuscar)
        Else
            If MsgBox("EL PRODUCTO YA TIENE LA CANTIDAD DE " & tRs.Fields("CANTIDAD") & ", ¿DESEA SUMAR LA CANTIDAD ANOTADA A ESTA CANTIDAD? SI CLIKEA SI, SE SUMARAN, SI CLIKEA NO, SE SUSTITUIRA LA CANTIDAD POR LA ANOTADA", vbYesNo + vbCritical + vbDefaultButton1, "INVENTARIO") = vbYes Then
                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(Text2.Text) + CDbl(tRs.Fields("CANTIDAD")) & " WHERE ID_PRODUCTO = '" & Text1.Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
            Else
                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & Text2.Text & " WHERE ID_PRODUCTO = '" & Text1.Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
            End If
            cnn.Execute (sBuscar)
        End If
        Text1.Text = ""
        Text2.Text = ""
    Else
        MsgBox "FALTA INFORMACION NECESARIA", vbInformation, "INVENTARIOS"
    End If
End Sub
