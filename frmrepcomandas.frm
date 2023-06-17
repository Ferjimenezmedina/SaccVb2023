VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmrepcomandas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Status de Comandas"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9960
      TabIndex        =   14
      Top             =   6720
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmrepcomandas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmrepcomandas.frx":030A
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
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9960
      TabIndex        =   12
      Top             =   5520
      Width           =   975
      Begin VB.Label Label19 
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
      Begin VB.Image Image3 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmrepcomandas.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "frmrepcomandas.frx":26F6
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9960
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmrepcomandas.frx":4238
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.OptionButton Option3 
         Caption         =   "Sucursal"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Status"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5175
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   9128
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
         Caption         =   "Rango de Fechas"
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9495
         Begin VB.CommandButton cmdBuscar 
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
            Left            =   5040
            Picture         =   "frmrepcomandas.frx":4254
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1920
            TabIndex        =   8
            Top             =   480
            Width           =   2895
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   6960
            TabIndex        =   4
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   69664769
            CurrentDate     =   40071
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   6960
            TabIndex        =   3
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   69664769
            CurrentDate     =   40071
         End
         Begin VB.Label Label1 
            Caption         =   "Comanda :"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Al:"
            Height          =   255
            Left            =   6360
            TabIndex        =   7
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "De: "
            Height          =   255
            Left            =   6360
            TabIndex        =   6
            Top             =   360
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "frmrepcomandas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Sub cmdBuscar_Click()
Dim sBuscar As String
    Dim tRs As Recordset
    Dim tRs1 As Recordset
    Dim sNombre As String
    Dim tot As Double
    Dim abo   As Double
    Dim Pend  As Double
    ListView1.ListItems.Clear
    sBuscar = "SELECT * FROM vsstcomandas WHERE id_comanda  like '%" & Text1.Text & "%' AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " ' "
     Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_COMANDA"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(2) = tRs.Fields("ID_PRODUCTO")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(3) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("CANTIDAD_NO_SIRVIO")) Then tLi.SubItems(4) = tRs.Fields("CANTIDAD_NO_SIRVIO")
            tLi.SubItems(5) = CDbl(tLi.SubItems(3)) - CDbl(tLi.SubItems(4))
            If Not IsNull(tRs.Fields("ESTADO_ACTUAL")) Then
                If tRs.Fields("ESTADO_ACTUAL") = "A" Then
                    tLi.SubItems(6) = "Nueva"
                Else
                    If tRs.Fields("ESTADO_ACTUAL") = "R" Or tRs.Fields("ESTADO_ACTUAL") = "S" Then
                        tLi.SubItems(6) = "En Producción"
                    Else
                        If tRs.Fields("ESTADO_ACTUAL") = "P" Then
                            tLi.SubItems(6) = "Probando en Calidad"
                        Else
                            If tRs.Fields("ESTADO_ACTUAL") = "N" Or tRs.Fields("ESTADO_ACTUAL") = "M" Then
                                tLi.SubItems(6) = "Cartuchos Dañados"
                            Else
                                If tRs.Fields("ESTADO_ACTUAL") = "L" Then
                                    tLi.SubItems(6) = "Terminado"
                                Else
                                    If tRs.Fields("ESTADO_ACTUAL") = "Z" Then
                                        tLi.SubItems(6) = "Aprovar Rema"
                                    Else
                                        If tRs.Fields("ESTADO_ACTUAL") = "C" Or tRs.Fields("ESTADO_ACTUAL") = "0" Then
                                            tLi.SubItems(6) = "CANCELADA"
                                           Else
                                             If tRs.Fields("ESTADO_ACTUAL") = "I" Or tRs.Fields("ESTADO_ACTUAL") = "0" Then
                                            tLi.SubItems(6) = "COBRADO"
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If Not IsNull(tRs.Fields("FECHA_FIN")) Then tLi.SubItems(7) = tRs.Fields("FECHA_FIN")
            If Not IsNull(tRs.Fields("SUCURSAL")) Then tLi.SubItems(8) = tRs.Fields("SUCURSAL")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Form_Load()
    'VsRepCXC
    DTPicker1.Value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
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
        .FullRowSelect = True
        .ColumnHeaders.Add , , "No COMANDA", 1200
        .ColumnHeaders.Add , , "NOMBRE", 4000
        .ColumnHeaders.Add , , "ID PRODUCTO", 1200
        .ColumnHeaders.Add , , "CANTIDAD", 1300
        .ColumnHeaders.Add , , "CANTIDAD NO FUNCIONO", 1300
        .ColumnHeaders.Add , , "CANTIDAD FUNCIONO", 1300
        .ColumnHeaders.Add , , "ESTADO ACTUAL", 1300
        .ColumnHeaders.Add , , "FECHA DE INICIO", 1300
        .ColumnHeaders.Add , , "FECHA DE TERMINO", 1300
        .ColumnHeaders.Add , , "SUCURSAL", 1300
    End With
End Sub
Private Sub Image3_Click()
On Error GoTo ManejaError
    CommonDialog1.DialogTitle = "Guardar Como"
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "[Archivo Excel (*.xls)] |*.xls|"
    CommonDialog1.ShowOpen
    FILE = CommonDialog1.FileName
    Dim ApExcel As Excel.Application
    Set ApExcel = CreateObject("Excel.application")
    ApExcel.Workbooks.Add
    Dim Conta As Integer
    Dim Col As Integer
    For cont = 1 To ListView1.ColumnHeaders.Count
        ApExcel.Cells(1, cont) = ListView1.ColumnHeaders(cont)
        ApExcel.Cells(1, cont).Font.Bold = True
        ApExcel.Cells(1, cont).Font.Color = vbRed
    Next cont
    With ApExcel
        For Fila = 2 To ListView1.ListItems.Count + 1
            Col = 1
            .Cells(Fila, Col) = ListView1.ListItems.Item(Fila - 1)
            For Col = 1 To ListView1.ColumnHeaders.Count - 1
                .Cells(Fila, Col + 1) = _
                ListView1.ListItems(Fila - 1).SubItems(Col)
            Next
        Next
    End With
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
Private Sub Image9_Click()
    Unload Me
End Sub

