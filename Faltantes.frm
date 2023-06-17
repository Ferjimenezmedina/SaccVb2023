VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Faltantes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Faltantes"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   ControlBox      =   0   'False
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9720
      TabIndex        =   12
      Top             =   3720
      Width           =   975
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "Faltantes.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Faltantes.frx":030A
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EXCEL"
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
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9720
      TabIndex        =   4
      Top             =   2520
      Width           =   975
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   9
         Top             =   1320
         Width           =   975
         Begin VB.Image Image5 
            Height          =   720
            Left            =   120
            MouseIcon       =   "Faltantes.frx":1E4C
            MousePointer    =   99  'Custom
            Picture         =   "Faltantes.frx":2156
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label7 
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
            TabIndex        =   10
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   7
         Top             =   1320
         Width           =   975
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Aceptar"
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
         Begin VB.Image Image4 
            Height          =   675
            Left            =   120
            MouseIcon       =   "Faltantes.frx":3B18
            MousePointer    =   99  'Custom
            Picture         =   "Faltantes.frx":3E22
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   5
         Top             =   0
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
            TabIndex        =   6
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image1 
            Height          =   705
            Left            =   120
            MouseIcon       =   "Faltantes.frx":564C
            MousePointer    =   99  'Custom
            Picture         =   "Faltantes.frx":5956
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Eliminar"
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
      Begin VB.Image Image2 
         Height          =   780
         Left            =   120
         MouseIcon       =   "Faltantes.frx":7408
         MousePointer    =   99  'Custom
         Picture         =   "Faltantes.frx":7712
         Top             =   120
         Width           =   780
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9720
      TabIndex        =   2
      Top             =   4920
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Faltantes.frx":9704
         MousePointer    =   99  'Custom
         Picture         =   "Faltantes.frx":9A0E
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "Faltantes.frx":BAF0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.ListView ListView2 
         Height          =   5535
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   9763
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9960
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Faltantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Faltantes.Caption = "Faltantes al dia " & Format(Date, "dd/mm/yyyy")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    '**********************************LISTVIEW2**************************************
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 2300
        .ColumnHeaders.Add , , "Descripcion", 4700
        .ColumnHeaders.Add , , "CANT. MINIMA", 1500
        .ColumnHeaders.Add , , "CANT. MAXIMA", 1500
        .ColumnHeaders.Add , , "CANT. EN EXISTENCIA", 1500
        .ColumnHeaders.Add , , "SUCURSAL", 1500
        .ColumnHeaders.Add , , "ALMACEN", 1500
    End With
    Buscar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Buscar(Optional ByVal Siguiente As Boolean = False)
On Error GoTo ManejaError
    '*************************************ALMACEN 1*************************************
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT * FROM VsExisAlmacen1 ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            ListView2.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                If Not IsNull(.Fields("C_MINIMA")) Then tLi.SubItems(2) = .Fields("C_MINIMA") & ""
                If Not IsNull(.Fields("C_MAXIMA")) Then tLi.SubItems(3) = .Fields("C_MAXIMA") & ""
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(4) = .Fields("CANTIDAD") & ""
                If Not IsNull(.Fields("SUCURSAL")) Then tLi.SubItems(5) = .Fields("SUCURSAL") & ""
                tLi.SubItems(6) = "Almacen 1"
                If .Fields("C_MINIMA") > .Fields("CANTIDAD") Then
                    tLi.ForeColor = vbRed
                Else
                    tLi.ForeColor = vbBlack
                End If
                .MoveNext
            Loop
        End If
    End With
    '***************************************ALMACEN 2***********************************
    sBuscar = "SELECT * FROM VsExisAlmacen2 ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                If Not IsNull(.Fields("C_MINIMA")) Then tLi.SubItems(2) = .Fields("C_MINIMA") & ""
                If Not IsNull(.Fields("C_MAXIMA")) Then tLi.SubItems(3) = .Fields("C_MAXIMA") & ""
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(4) = .Fields("CANTIDAD") & ""
                If Not IsNull(.Fields("SUCURSAL")) Then tLi.SubItems(5) = .Fields("SUCURSAL") & ""
                tLi.SubItems(6) = "Almacen 2"
                If .Fields("C_MINIMA") > .Fields("CANTIDAD") Then
                    tLi.ForeColor = vbRed
                Else
                    tLi.ForeColor = vbBlack
                End If
                .MoveNext
            Loop
        End If
    End With
    '****************************************ALMACEN 3**********************************
    sBuscar = "SELECT * FROM VsExisAlmacen3 ORDER BY ID_PRODUCTO"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                If Not IsNull(.Fields("C_MINIMA")) Then tLi.SubItems(2) = .Fields("C_MINIMA") & ""
                If Not IsNull(.Fields("C_MAXIMA")) Then tLi.SubItems(3) = .Fields("C_MAXIMA") & ""
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(4) = .Fields("CANTIDAD") & ""
                If Not IsNull(.Fields("SUCURSAL")) Then tLi.SubItems(5) = .Fields("SUCURSAL") & ""
                tLi.SubItems(6) = "Almacen 3 (Ventas)"
                If .Fields("C_MINIMA") > .Fields("CANTIDAD") Then
                    tLi.ForeColor = vbRed
                Else
                    tLi.ForeColor = vbBlack
                End If
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Public Function EliminarItem(Lw As ListView)
On Error GoTo ManejaError
    With Lw
        If .ListItems.Count > 0 Then
            .ListItems(.SelectedItem.Index).Selected = True
            .ListItems.Remove .SelectedItem.Index
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub Image10_Click()
    On Error GoTo ManejaError
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
    If ListView2.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = ListView2.ColumnHeaders.Count
            For Con = 1 To ListView2.ColumnHeaders.Count
                StrCopi = StrCopi & ListView2.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView2.ListItems.Count
                StrCopi = StrCopi & ListView2.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView2.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub Image2_Click()
    On Error GoTo ManejaError
    With ListView2
        If .ListItems.Count > 0 Then
            .ListItems(.SelectedItem.Index).Selected = True
            .ListItems.Remove .SelectedItem.Index
        End If
    End With
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
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
End Sub
