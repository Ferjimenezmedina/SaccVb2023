VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmlocalizaalm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Localizacion de Productos"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8160
      TabIndex        =   2
      Top             =   3960
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmlocalizaalm.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmlocalizaalm.frx":030A
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8280
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmlocalizaalm.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdBuscar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame15"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.Frame Frame1 
         Caption         =   "ALMACENES"
         Height          =   1575
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   4695
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   3345
            TabIndex        =   19
            Top             =   240
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Almacen 1"
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Almacen 2"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Almacen 3"
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   1920
            TabIndex        =   15
            Top             =   240
            Width           =   2055
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   1920
            TabIndex        =   14
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.CommandButton Command2 
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
         Left            =   2880
         Picture         =   "frmlocalizaalm.frx":2408
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   2535
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4048
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame15 
         Caption         =   "SUCURSALES"
         Height          =   855
         Left            =   5400
         TabIndex        =   7
         Top             =   120
         Width           =   2175
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdBuscar 
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
         Left            =   6480
         Picture         =   "frmlocalizaalm.frx":4DDA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4320
         TabIndex        =   4
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Producto :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Ubicacion :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmlocalizaalm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
 Dim StrRep As String

Private Sub Check1_Click()
Combo1.Visible = True
Combo3.Visible = False
Combo4.Visible = False
Dim sBuscar As String

End Sub



Private Sub Check2_Click()

 Dim sBuscar As String
   sBuscar = "SELECT ID_PRODUCTO FROM VSINVALM2  ORDER BY ID_PRODUCTO "
            Set tRs = cnn.Execute(sBuscar)
            Combo3.AddItem "<TODAS>"
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
            Combo3.AddItem tRs.Fields("ID_PRODUCTO")
            tRs.MoveNext
            Loop
            End If

End Sub

Private Sub Check3_Click()

 Dim sBuscar As String
   sBuscar = "SELECT ID_PRODUCTO FROM vsloc5  ORDER BY ID_PRODUCTO "
            Set tRs = cnn.Execute(sBuscar)
            Combo4.AddItem "<TODAS>"
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
            Combo4.AddItem tRs.Fields("ID_PRODUCTO")
            'If Not IsNull(tRs.Fields(Combo4.) Then Combo4 tRs.Fields("ID_PRODUCTO")
            tRs.MoveNext
            Loop
            End If
End Sub

Private Sub cmdBuscar_Click()
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
     
   If Check3.Value = 1 Or Check2.Value = 1 Then
      
         If Check1.Value = 1 Then
   
            sBuscar = "UPDATE vsloc SET LOCALIZACION= '" & Text3.Text & "' WHERE (ID_PRODUCTO='" & Text1.Text & "' OR ID_PRODUCTO='" & Combo1.Text & "') AND  SUCURSAL = '" & Combo2.Text & "' "
         End If
         If Check2.Value = 1 Then
            sBuscar = "UPDATE ALMACEN2 SET LOCALIZACION= '" & Text3.Text & "' WHERE ID_PRODUCTO='" & Text1.Text & "' "
         End If
         If Check3.Value = 1 Then
            sBuscar = "  UPDATE  ALMACEN3 SET LOCALIZACION= '" & Text3.Text & "' WHERE  ID_PRODUCTO='" & Text1.Text & "' "
         End If
     Set tRs = cnn.Execute(sBuscar)
     MsgBox ("LA INFORMACION FUE PROCESADA")
     
  Else
  MsgBox ("FALTA  OPCION DE FILTRAR PARA  PODER  PROCESAR")
  
  End If
  

   Text3.Text = ""
   Text1.Text = ""
   'Combo2.Text = ""
   'Combo1.Text = ""
   ListView1.ListItems.Clear
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    
    ListView1.ListItems.Clear
   If Check2.Value = 1 Then
    sBuscar = "SELECT * FROM VSINVALM2 WHERE SUCURSAL = '" & Combo2.Text & "'  AND ID_PRODUCTO LIKE '%" & Text1.Text & "%' ORDER BY ID_PRODUCTO"
   End If
   If Check3.Value = 1 Then
    sBuscar = "SELECT * FROM VSINVALM3 WHERE SUCURSAL = '" & Combo2.Text & "'  AND ID_PRODUCTO LIKE '%" & Text1.Text & "%' ORDER BY ID_PRODUCTO"
   End If
 'sBuscar = "SELECT ID_VENTA,FOLIO,FECHA,NOMBRE,BANCO,FECHAABONO,TOTAL,CANT_ABONO,NO_CHEQUE,DEUDA,PAGADA,LIMITE_CREDITO,TOTAL_COMPRA FROM temporal_abonos WHERE FOlIO NOT IN ('CANCELADO') AND NOMBRE LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " ' ORDER BY NOMBRE,FECHA ASC "
    
 StrRep = sBuscar
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            tLi.SubItems(1) = tRs.Fields("DESCRIPCION")
            If Not IsNull(tRs.Fields("LOCA")) Then tLi.SubItems(2) = tRs.Fields("LOCA")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(3) = tRs.Fields("CANTIDAD")
             tRs.MoveNext
        Loop
      
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
    
   '   If Check1.Value = 1 Then
     '    sBuscar = "SELECT ID_PRODUCTO FROM vsinvalm1 ORDER BY ID_PRODUCTO"
    '        Set tRs = cnn.Execute(sBuscar)
      '      Combo1.AddItem "<TODAS>"
       ' If Not (tRs.EOF And tRs.BOF) Then
        '    Do While Not tRs.EOF
         '   Combo1.AddItem tRs.Fields("ID_PRODUCTO")
          '  tRs.MoveNext
           ' Loop
         ' End If
       ' End If
     sBuscar = "SELECT NOMBRE FROM SUCURSALES ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    Combo2.AddItem "<TODAS>"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo2.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
    
   ' sBuscar = "SELECT ID_PRODUCTO FROM EXISTENCIAS ORDER BY ID_PRODUCTO"
    'Set tRs = cnn.Execute(sBuscar)
    'Combo1.AddItem "<TODAS>"
    'If Not (tRs.EOF And tRs.BOF) Then
     '   Do While Not tRs.EOF
      '      Combo1.AddItem tRs.Fields("ID_PRODUCTO")
       '     tRs.MoveNext
       ' Loop
'    End If
  
          sBuscar = "SELECT ID_PRODUCTO FROM VSINVALM1  ORDER BY ID_PRODUCTO "
            Set tRs = cnn.Execute(sBuscar)
            Combo1.AddItem "<TODAS>"
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("ID_PRODUCTO")
            tRs.MoveNext
            Loop
End If

   
        'If Check2.Value = 1 Then
            'sBuscar = "SELECT ID_PRODUCTO FROM vsinvalm2 ORDER BY ID_PRODUCTO"
           ' Set tRs = cnn.Execute(sBuscar)
          '  Combo1.AddItem "<TODAS>"
         '   If Not (tRs.EOF And tRs.BOF) Then
        '    Do While Not tRs.EOF
       '     Combo1.AddItem tRs.Fields("ID_PRODUCTO")
      '      tRs.MoveNext
     '       Loop
    '    End If
   ' End If
    
       ' If Check3.Value = 1 Then
        '    sBuscar = "SELECT ID_PRODUCTO FROM vsinvalm3 ORDER BY ID_PRODUCTO"
         '     Set tRs = cnn.Execute(sBuscar)
         ' Combo1.AddItem "<TODAS>"
          '   If Not (tRs.EOF And tRs.BOF) Then
           '    Do While Not tRs.EOF
           ' Combo1.AddItem tRs.Fields("ID_PRODUCTO")
           ' tRs.MoveNext
           'Loop
         'End If
           With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id_Producto", 1000
        .ColumnHeaders.Add , , "Descripcion", 1200
        .ColumnHeaders.Add , , "Localizacion", 1200
        .ColumnHeaders.Add , , "Cantidad", 1200
     
        End With
       
End Sub

Private Sub Image9_Click()
Unload Me
End Sub




'Private Sub Image1_Click()

   ' Dim Path As String
    'Dim SelectionFormula As Date
     'Path = App.Path
     'sBuscar = "SELECT NOMBRE,ID_VENTA,FOLIO,FECHA,NOMBRE,BANCO,CANT_ABONO,NO_CHEQUE,DEUDA,LIMITE_CREDITO FROM VsRepAbonos WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY FECHA"
    'StrRep = sBuscar
      '  Set crReport = crApplication.OpenReport(Path & "\REPORTES\cuentasco.rpt")
       ' crReport.Database.LogOnServer "p2ssql.dll", GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "1"), GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER"), GetSetting("APTONER", "ConfigSACC", "USUARIO", "1"), GetSetting("APTONER", "ConfigSACC", "PASSWORD", "1")
        'crReport.SelectionFormula = "{VSABONOS.FECHA}>=Date (Year (" & DTPicker1.Value & "),Month (" & DTPicker1.Value & ") , Day (" & DTPicker1.Value & ")) and {VSABONOS.FECHA}<=Date (Year (" & DTPicker2.Value & "),Month (" & DTPicker2.Value & ") , Day ( " & DTPicker2.Value & "))"
        'crReport.Action = 0
       'crReport.Destination = crptToPrinter
      'crReport.Destination = crptToWindow
       ' crReport.SQLQueryString = StrRep
        'crReport.SQLQueryString = "SELECT NOMBRE,ID_VENTA,FOLIO,FECHA,NOMBRE,BANCO,CANT_ABONO,NO_CHEQUE,DEUDA,LIMITE_CREDITO FROM VsRepAbonos WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "' ORDER BY FECHA"
       ' crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        'frmRep.Show vbModal, Me
        
 '    End Sub



'Private Sub Image10_Click()
'    If ListView1.ListItems.Count > 0 Then
 '       Dim StrCopi As String
  '      Dim Con As Integer
   '     Dim Con2 As Integer
    '    Dim NumColum As Integer
     '   Dim Ruta As String
      '  Me.CommonDialog1.FileName = ""
       ' CommonDialog1.DialogTitle = "Guardar como"
  '      CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
   '     Me.CommonDialog1.ShowSave
    '    Ruta = Me.CommonDialog1.FileName
        'StrCopi = "Nota" & Chr(9) & "Factura" & Chr(9) & "Fecha" & Chr(9) & "Nombre" & Chr(9) & "Banco" & Chr(9) & "Monto" & Chr(9) & "DEUDA" & Chr(9) & Chr(9) & "LIMITE_CREDITO" & Chr(13)
     '   If Ruta <> "" Then
      '      NumColum = ListView1.ColumnHeaders.Count
       '     For Con = 1 To ListView1.ListItems.Count
        '        StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
         '       For Con2 = 1 To NumColum - 1
          '          StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
           '     Next
            '    StrCopi = StrCopi & Chr(13)
            'n ext
            'Te 'xt2.Text = StrCopi
            ''ar'chivo TXT
           ' Dim foo As Integer
           ' foo = FreeFile
           ' Open Ruta For Output As #foo
            '    Print #foo, Text2.Text
            'Close #foo
       ' End If
    'End If
'End Sub
'Private Sub Image9_Click()
 '   Unload Me
'End Sub

'Private Sub Option1_Click()

'End Sub





Private Sub TabStrip1_Click()

End Sub

Private Sub ListView1_DblClick()
    Text1.SetFocus
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Text1.Text = Item
   

End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
  If Combo2.Text <> "" Then
    
    If KeyAscii = 13 Then
        Me.cmdBuscar.Visible = True
    End If
  Else
  MsgBox ("FALTA  OPCION DE FILTRAR PARA  PODER  PROCESAR")
  End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command2.Value = True
    End If
End Sub




