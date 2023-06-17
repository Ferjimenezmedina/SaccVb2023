VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form listadodecomandas 
   BackColor       =   &H80000009&
   Caption         =   "Listado de Comandas con su estado actual"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10560
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10440
      TabIndex        =   0
      Top             =   6000
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "listadodecomandas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "listadodecomandas.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "listadodecomandas.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDesc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.Frame Frame2 
         Caption         =   "Material Nuevo"
         Enabled         =   0   'False
         Height          =   3615
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Width           =   5295
         Begin VB.TextBox txtID 
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Text            =   "Text4"
            Top             =   3120
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.CommandButton Command6 
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
            Left            =   3960
            Picture         =   "listadodecomandas.frx":2408
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Insertar"
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
            Left            =   2040
            Picture         =   "listadodecomandas.frx":4DDA
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   3120
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   3840
            TabIndex        =   13
            Top             =   2760
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   840
            TabIndex        =   12
            Top             =   360
            Width           =   3015
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1935
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   4815
            _ExtentX        =   8493
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
         Begin VB.Label Label5 
            Caption         =   "------------------------------------------------"
            Height          =   255
            Left            =   720
            TabIndex        =   21
            Top             =   2760
            Width           =   2295
         End
         Begin VB.Label Label4 
            Caption         =   "Clave:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Cantidad :"
            Height          =   255
            Left            =   3120
            TabIndex        =   19
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Clave"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Materiales"
         Height          =   6615
         Left            =   5520
         TabIndex        =   5
         Top             =   480
         Width           =   4455
         Begin VB.CommandButton Command5 
            Caption         =   "Quitar"
            Enabled         =   0   'False
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
            Left            =   3480
            Picture         =   "listadodecomandas.frx":77AC
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   5040
            Width           =   975
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   4335
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   7646
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
         Begin VB.Label lblID3 
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   6000
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label7 
            Caption         =   "Clave:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   4680
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "------------------------------------------------"
            Height          =   255
            Left            =   1080
            TabIndex        =   8
            Top             =   4680
            Width           =   2295
         End
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4320
         Picture         =   "listadodecomandas.frx":A17E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1815
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3201
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
      Begin VB.Label lblDesc 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2880
         Width           =   5295
      End
      Begin VB.Label Label1 
         Caption         =   "Clave del Juego de Reparacion"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "listadodecomandas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim StrRep As String
Dim StrRep2 As String
Dim StrRep3 As String
Dim StrRep4 As String
Private Sub Command1_Click()
    On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    
    If Text3.Text <> "" Then
        sBuscar = "INSERT INTO JUEGO_REPARACION (ID_REPARACION, ID_PRODUCTO, CANTIDAD) VALUES ('" & txtID.Text & "', '" & Label5.Caption & "', " & Text3.Text & ");"
        cnn.Execute (sBuscar)
        sBuscar = "SELECT * FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & txtID.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        ListView3.ListItems.Clear
        With tRs
            If Not (.BOF And .EOF) Then
                ListView3.ListItems.Clear
                .MoveFirst
                Do While Not .EOF
                    Set tLi = ListView3.ListItems.Add(, , .Fields("ID_REPARACION") & "")
                    tLi.SubItems(1) = .Fields("ID_PRODUCTO") & ""
                    tLi.SubItems(2) = .Fields("CANTIDAD") & ""
                    .MoveNext
                Loop
            End If
        End With
        Label5.Caption = "------------------------------------------------"
        Text3.Text = ""
    Else
        MsgBox "LA CANTIDAD DEBE SER UN NUMERO", vbInformation, "SACC"
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





Private Sub Command3_Click()
    Buscar
End Sub

Private Sub Command5_Click()
    Dim sBuscar As String
    Dim tLi As ListItem
    Dim tRs As Recordset
    
    If lblID3.Caption <> "" Then
        sBuscar = "DELETE FROM JUEGO_REPARACION WHERE ID_PRODUCTO = '" & Label6.Caption & "' AND ID_REPARACION = '" & txtID.Text & "'"
        cnn.Execute (sBuscar)
        
        sBuscar = "SELECT * FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & txtID.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            ListView3.ListItems.Clear
            If Not (.BOF And .EOF) Then
                .MoveFirst
                Do While Not .EOF
                    Set tLi = ListView3.ListItems.Add(, , .Fields("ID_REPARACION") & "")
                    tLi.SubItems(1) = .Fields("ID_PRODUCTO") & ""
                    tLi.SubItems(2) = .Fields("CANTIDAD") & ""
                    .MoveNext
                Loop
            Else
                Frame2.Enabled = False
                Command5.Enabled = False
                Label6.Caption = "------------------------------------------------"
                txtID.Text = ""
                lblDesc.Caption = ""
            End If
        End With
    End If
End Sub

Private Sub Command6_Click()
    On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    
    sBuscar = Text2.Text
    sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & sBuscar & "%'"
    Set tRs = cnn.Execute(sBuscar)
    ListView2.ListItems.Clear
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    tLi.SubItems(1) = .Fields("DESCRIPCION") & ""
                .MoveNext
            Loop
        End If
    End With
    sBuscar = Text2.Text
    sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION FROM ALMACEN1 WHERE ID_REPARACION LIKE '%" & sBuscar & "%'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    tLi.SubItems(1) = .Fields("DESCRIPCION") & ""
                     tLi.SubItems() = .Fields("CANTIDAD") & ""
                .MoveNext
            Loop
        End If
    End With
    If ListView2.ListItems.Count = 0 Then
        MsgBox "El producto no existe en el almacen"
    End If
    '///////////////////////////////////////////////////////////////////////////////
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

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
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
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Juego", 3200
        .ColumnHeaders.Add , , "Descripcion", 3200
    End With
    
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 1200
        .ColumnHeaders.Add , , "Descripcion", 3200
    End With
    With ListView3
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Juego", 1200
        .ColumnHeaders.Add , , "Clave del Producto", 1200
        .ColumnHeaders.Add , , "Cantidad", 1000
    End With
    With ListView4
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_PRODU", 2000
        .ColumnHeaders.Add , , "Juego de Reparacion", 2000
       
    End With

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
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = Text1.Text
    Frame2.Enabled = False
    txtID.Text = ""
    lblDesc.Caption = ""
    sBuscar = "SELECT J.ID_REPARACION, A.DESCRIPCION FROM JUEGO_REPARACION AS J JOIN ALMACEN3 AS A ON J.ID_REPARACION = A.ID_PRODUCTO WHERE ID_REPARACION LIKE '%" & sBuscar & "%' GROUP BY ID_REPARACION, DESCRIPCION"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_REPARACION") & "")
                    tLi.SubItems(1) = .Fields("DESCRIPCION")
                .MoveNext
            Loop
        Else
            ListView1.ListItems.Clear
            MsgBox "El producto buscado no es juego de reparacion o no existe"
        End If
    End With
    
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


Private Sub Image1_Click()

Dim FILE As String
On Error GoTo ManejaError
    CommonDialog1.DialogTitle = "Guardar Como"
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "[Archivo Excel (*.xls)] |*.xls|"
    CommonDialog1.ShowOpen
    FILE = CommonDialog1.FileName
    Dim ApExcel As Excel.Application
    Set ApExcel = CreateObject("Excel.application")
    ApExcel.Workbooks.Add
    Dim Cont As Integer
    Dim Col As Integer
       Dim Fila As Integer
    
     If StrRep2 <> "" Then
            For Cont = 1 To ListView4.ColumnHeaders.Count
                ApExcel.Cells(1, Cont) = ListView4.ColumnHeaders(Cont)
                ApExcel.Cells(1, Cont).Font.Bold = True
                ApExcel.Cells(1, Cont).Font.Color = vbRed
            Next Cont
        With ApExcel
            For Fila = 2 To ListView4.ListItems.Count + 1
                Col = 1
                .Cells(Fila, Col) = ListView4.ListItems.Item(Fila - 1)
                 For Col = 1 To ListView4.ColumnHeaders.Count - 1
                     .Cells(Fila, Col + 1) = _
                      ListView4.ListItems(Fila - 1).SubItems(Col)
                 Next
            Next
            End With
     End If
     
    ApExcel.AlertBeforeOverwriting = False
    ApExcel.ActiveWorkbook.SaveAs "" & FILE
    ApExcel.Visible = True
    Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub


Private Sub Image10_Click()
Dim FILE As String
On Error GoTo ManejaError
    CommonDialog1.DialogTitle = "Guardar Como"
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "[Archivo Excel (*.xls)] |*.xls|"
    CommonDialog1.ShowOpen
    FILE = CommonDialog1.FileName
    Dim ApExcel As Excel.Application
    Set ApExcel = CreateObject("Excel.application")
    ApExcel.Workbooks.Add
    Dim Cont As Integer
    Dim Col As Integer
       Dim Fila As Integer
    If StrRep <> "" Then
            For Cont = 1 To ListView3.ColumnHeaders.Count
                ApExcel.Cells(1, Cont) = ListView3.ColumnHeaders(Cont)
                ApExcel.Cells(1, Cont).Font.Bold = True
                ApExcel.Cells(1, Cont).Font.Color = vbRed
            Next Cont
        With ApExcel
            For Fila = 2 To ListView3.ListItems.Count + 1
                Col = 1
                .Cells(Fila, Col) = ListView3.ListItems.Item(Fila - 1)
                 For Col = 1 To ListView3.ColumnHeaders.Count - 1
                     .Cells(Fila, Col + 1) = _
                      ListView3.ListItems(Fila - 1).SubItems(Col)
                 Next
            Next
            End With
     End If
     
     
    ApExcel.AlertBeforeOverwriting = False
    ApExcel.ActiveWorkbook.SaveAs "" & FILE
    ApExcel.Visible = True
    Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
    


Private Sub Image3_Click()

Dim oExcel As Object
  Dim oBook As Object
  Dim oSheet As Object

   'Start a new workbook in Excel
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add


   'Add data to cells of the first worksheet in the new workbook
  Set oSheet = oBook.Worksheets(1)
  oSheet.Range("A1").Value = "Last Name"
   oSheet.Range("B1").Value = "First Name"
   oSheet.Range("A1:B1").Font.Bold = True
   oSheet.Range("A2").Value = "Doe"
   oSheet.Range("B2").Value = "John"

   'Save the Workbook and Quit Excel
   oBook.SaveAs "C:\Book1.xls"
   oExcel.Quit

      If ListView3.ListItems.Count > 0 Then
        Dim StrCopi As String
        Dim Con As Integer
        Dim Con2 As Integer
        Dim NumColum As Integer
        Dim Ruta As String
        Me.CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "Guardar como"
        CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
        Me.CommonDialog1.ShowSave
        Ruta = Me.CommonDialog1.FileName
        StrCopi = "Clave_producto" & Chr(9) & "Clave_del_juego de Reparacion" & Chr(13)
        If Ruta <> "" Then
           NumColum = ListView4.ColumnHeaders.Count
            For Con = 1 To ListView4.ListItems.Count
               StrCopi = StrCopi & ListView3.ListItems.Item(Con) & Chr(9)
               For Con2 = 1 To NumColum - 1
                   StrCopi = StrCopi & ListView4.ListItems.Item(Con).SubItems(Con2) & Chr(9)
               Next
                StrCopi = StrCopi & Chr(13)
            Next
            Text4.Text = StrCopi
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
           Open Ruta For Output As #foo
                Print #foo, Text4.Text
            Close #foo
        End If
    End If
    
End Sub

Private Sub Image9_Click()
    Unload Me
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    
    If ListView1.ListItems.Count > 0 Then
        txtID.Text = Item
        lblDesc.Caption = Item.SubItems(1)
        sBuscar = Item
        sBuscar = "SELECT * FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & sBuscar & "'"
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If Not (.BOF And .EOF) Then
                ListView3.ListItems.Clear
                .MoveFirst
                Do While Not .EOF
                    Set tLi = ListView3.ListItems.Add(, , .Fields("ID_REPARACION") & "")
                    tLi.SubItems(1) = .Fields("ID_PRODUCTO") & ""
                    tLi.SubItems(2) = .Fields("CANTIDAD") & ""
                    .MoveNext
                
                Loop
               
                Command5.Enabled = True
                StrRep = sBuscar
            End If
        End With
        Frame2.Enabled = True
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

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Label5.Caption = Item
End Sub

Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView3.ListItems.Count > 0 Then
        lblID3.Caption = Item.Index
        Label6.Caption = Item.ListSubItems(1)
        Frame11.Visible = True
         Frame4.Visible = False
        
    End If
End Sub
Private Sub ListView4_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView4.ListItems.Count > 0 Then
       
        Frame11.Visible = False
         Frame4.Visible = True
        
    End If
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error GoTo ManejaError
    If KeyAscii = 13 Then
        If Text1.Text <> "" Then
            Buscar
        End If
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ%1234567890- "
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

Private Sub Text2_KeyPress(KeyAscii As Integer)
    On Error GoTo ManejaError
    If KeyAscii = 13 Then
        If Text2.Text <> "" Then
            Command6.Value = True
        End If
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ%1234567890- "
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

Private Sub Text3_Change()
    Command1.Enabled = False
    If Label5.Caption <> "------------------------------------------------" And Text3.Text <> "" Then
        Command1.Enabled = True
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    On Error GoTo ManejaError
    If KeyAscii = 13 Then
        If Text3.Text <> "" Then
            'Buscar
        End If
    End If
    Dim Valido As String
    Valido = "1234567890."
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


