VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmImportarPrecios 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Lista de Precios"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView5 
      Height          =   255
      Left            =   7560
      TabIndex        =   44
      Top             =   840
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView4 
      Height          =   375
      Left            =   7440
      TabIndex        =   35
      Top             =   360
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   32
      Top             =   5880
      Width           =   975
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmImportarPrecios.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmImportarPrecios.frx":030A
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label6 
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
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   7320
      TabIndex        =   24
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   7560
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
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
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmImportarPrecios.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmImportarPrecios.frx":2156
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2566
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   1
      Top             =   2400
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmImportarPrecios.frx":3C98
         MousePointer    =   99  'Custom
         Picture         =   "FrmImportarPrecios.frx":3FA2
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6165
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " Precios"
      TabPicture(0)   =   "FrmImportarPrecios.frx":6084
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Check3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Inventario"
      TabPicture(1)   =   "FrmImportarPrecios.frx":60A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "Command4"
      Tab(1).Control(4)=   "Text2"
      Tab(1).Control(5)=   "Combo1"
      Tab(1).Control(6)=   "Command3"
      Tab(1).Control(7)=   "Check1"
      Tab(1).Control(8)=   "Check2"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Catálogo SAT"
      TabPicture(2)   =   "FrmImportarPrecios.frx":60BC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(2)=   "Text3"
      Tab(2).Control(3)=   "Command5"
      Tab(2).Control(4)=   "Frame4"
      Tab(2).Control(5)=   "Command6"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton Command6 
         Caption         =   "Importar"
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
         Left            =   -69360
         Picture         =   "FrmImportarPrecios.frx":60D8
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Almacen a actualizar"
         Height          =   1215
         Left            =   -74760
         TabIndex        =   39
         Top             =   1200
         Width           =   1815
         Begin VB.OptionButton Option11 
            Caption         =   "Almacen 1"
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Almacen 2"
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Almacen 3"
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   255
         Left            =   -68640
         TabIndex        =   37
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   720
         Width           =   6015
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Solo actualizar precios mayores"
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   2160
         Width           =   2535
      End
      Begin VB.OptionButton Option8 
         Caption         =   "En Dolares"
         Height          =   195
         Left            =   2160
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         Caption         =   "En Pesos"
         Height          =   195
         Left            =   2160
         TabIndex        =   27
         Top             =   1440
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Capturar como existencias"
         Height          =   255
         Left            =   -72720
         TabIndex        =   26
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Marcar como inventario inicial"
         Height          =   255
         Left            =   -72720
         TabIndex        =   25
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Importar"
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
         Left            =   5640
         Picture         =   "FrmImportarPrecios.frx":8AAA
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Importar"
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
         Left            =   -69360
         Picture         =   "FrmImportarPrecios.frx":B47C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -72720
         TabIndex        =   16
         Text            =   "Sucursal"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   6015
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   255
         Left            =   -68640
         TabIndex        =   14
         Top             =   720
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Caption         =   "Almacen a actualizar"
         Height          =   1215
         Left            =   -74760
         TabIndex        =   10
         Top             =   1200
         Width           =   1815
         Begin VB.OptionButton Option6 
            Caption         =   "Almacen 3"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Almacen 2"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Almacen 1"
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   6015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   6360
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Almacen a actualizar"
         Height          =   1215
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
         Begin VB.OptionButton Option1 
            Caption         =   "Almacen 1"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Almacen 2"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Almacen 3"
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.Label Label9 
         Caption         =   $"FrmImportarPrecios.frx":DE4E
         Height          =   615
         Left            =   -74880
         TabIndex        =   47
         Top             =   2640
         Width           =   6855
      End
      Begin VB.Label Label8 
         Caption         =   $"FrmImportarPrecios.frx":DF43
         Height          =   495
         Left            =   -74880
         TabIndex        =   46
         Top             =   2640
         Width           =   6855
      End
      Begin VB.Label Label2 
         Caption         =   $"FrmImportarPrecios.frx":DFF9
         Height          =   495
         Left            =   120
         TabIndex        =   45
         Top             =   2640
         Width           =   6855
      End
      Begin VB.Label Label7 
         Caption         =   "Archivo :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   38
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Archivo :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Archivo :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1455
      Left            =   120
      TabIndex        =   29
      Top             =   5880
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2566
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
   Begin MSComctlLib.ListView ListView3 
      Height          =   1455
      Left            =   120
      TabIndex        =   31
      Top             =   4080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2566
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "Productos actualizados..."
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
      Left            =   120
      TabIndex        =   30
      Top             =   5640
      Width           =   5895
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Productos no encontrados durante el proceso..."
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
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   5895
   End
End
Attribute VB_Name = "FrmImportarPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private obj_Excel       As Object
Private obj_Workbook    As Object
Private obj_Worksheet   As Object
Private Sub Command1_Click()
On Error GoTo Maneja
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Abrir"
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "[Archivo Excel (*.xls)] |*.xls|"
    CommonDialog1.ShowOpen
    Text1.Text = CommonDialog1.FileName
    Call Excel_FlexGrid(CommonDialog1.FileName, ListView4, 20, 5, "Hoja1")
Maneja:
    Err.Clear
End Sub
Private Sub Command2_Click()
    If Text1.Text <> "" Then
        Dim sBuscar As String
        Dim iAfectados As Long
        Dim tRs1 As ADODB.Recordset
        Dim tLi1 As ListItem
        Dim tLi2 As ListItem
        Dim tRs As ADODB.Recordset
        Dim Col As Integer, Fila As Integer
        Dim Con As Double
        For Con = 1 To ListView4.ListItems.COUNT
            If Option1.value Then
                If Option7.value Then
                    sBuscar = "UPDATE ALMACEN1 SET PRECIO_COSTO = " & ListView4.ListItems(Con).SubItems(1) & ", PRECIO_EN = 'PESOS' WHERE ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                Else
                    sBuscar = "UPDATE ALMACEN1 SET PRECIO_COSTO = " & ListView4.ListItems(Con).SubItems(1) & ", PRECIO_EN = 'DOLARES' WHERE ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                End If
            End If
            If Option2.value Then
                If Option7.value Then
                    sBuscar = "UPDATE ALMACEN2 SET PRECIO_COSTO = " & ListView4.ListItems(Con).SubItems(1) & ", PRECIO_EN = 'PESOS' WHERE ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                Else
                    sBuscar = "UPDATE ALMACEN2 SET PRECIO_COSTO = " & ListView4.ListItems(Con).SubItems(1) & ", PRECIO_EN = 'DOLARES' WHERE ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                End If
            End If
            If Option3.value Then
                sBuscar = "SELECT PRECIO_COSTO FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                Set tRs1 = cnn.Execute(sBuscar)
                If Not (tRs1.EOF And tRs1.BOF) Then
                    If Check3.value = 1 Then
                        If (ListView4.ListItems(Con).SubItems(1) > tRs1.Fields("PRECIO_COSTO")) Then
                            If Option7.value Then
                                sBuscar = "UPDATE ALMACEN3 SET PRECIO_COSTO = " & ListView4.ListItems(Con).SubItems(1) & ", PRECIO_EN = 'PESOS' WHERE ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                            Else
                                sBuscar = "UPDATE ALMACEN3 SET PRECIO_COSTO = " & ListView4.ListItems(Con).SubItems(1) & ", PRECIO_EN = 'DOLARES' WHERE ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                            End If
                            cnn.Execute (sBuscar)
                        End If
                    Else
                        If Option7.value Then
                            sBuscar = "UPDATE ALMACEN3 SET PRECIO_COSTO = " & ListView4.ListItems(Con).SubItems(1) & ", PRECIO_EN = 'PESOS' WHERE ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                        Else
                            sBuscar = "UPDATE ALMACEN3 SET PRECIO_COSTO = " & ListView4.ListItems(Con).SubItems(1) & ", PRECIO_EN = 'DOLARES' WHERE ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                        End If
                        cnn.Execute (sBuscar)
                    End If
                End If
            End If
            Set tRs = cnn.Execute(sBuscar, iAfectados, adCmdText)
            If iAfectados >= 1 Then
                Set tLi1 = ListView2.ListItems.Add(, , ListView4.ListItems(Con))
                tLi1.SubItems(1) = ListView4.ListItems(Con).SubItems(1)
                'tLi1.SubItems(2) = tRs1.Fields("PRECIO_COSTO")
            Else
                Set tLi2 = ListView3.ListItems.Add(, , ListView4.ListItems(Con))
                tLi2.SubItems(1) = ListView4.ListItems(Con).SubItems(1)
            End If
            Fila = Fila + 1
        Next
        MsgBox "La importación ha finalizado exitosamente!", vbExclamation, "SACC"
        Me.Height = 7950
        Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Else
        MsgBox "No se ha seleccionado ningun archivo a importar!", vbExclamation, "SACC"
    End If
End Sub
Private Sub Command3_Click()
    Dim tRs As ADODB.Recordset
    If Text2.Text <> "" Then
        Dim sBuscar As String
        Dim tLi As ListItem
        Dim Con As Double
        For Con = 1 To ListView4.ListItems.COUNT
            If Option4.value Then
                sBuscar = "SELECT ID_PRODUCTO FROM ALMACEN1 WHERE ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    sBuscar = "select ID_PRODUCTO, CANTIDAD, SUCURSAL from vsinvalm1 WHERE SUCURSAL = '" & Combo1.Text & "' AND ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        If tRs.Fields("CANTIDAD") = 0 Then
                            If Check1.value = 1 Then
                                sBuscar = "INSERT INTO EXISTENCIASTEMPORALES (ID_PRODUCTO, CANTIDAD,SUCURSAL) VALUES ('" & ListView4.ListItems(Con) & "', '" & ListView4.ListItems(Con).SubItems(1) & "', '" & ListView4.ListItems(Con).SubItems(2) & "');"
                                cnn.Execute (sBuscar)
                            End If
                            If Check2.value = 1 Then
                                sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD,SUCURSAL) VALUES ('" & ListView4.ListItems(Con) & "', '" & ListView4.ListItems(Con).SubItems(1) & "', '" & Combo1.Text & "');"
                                cnn.Execute (sBuscar)
                            End If
                        Else
                            If Check1.value = 1 Then
                                sBuscar = "UPDATE EXISTENCIASTEMPORALES SET CANTIDAD = " & ListView4.ListItems(Con).SubItems(1) & " WHERE SUCURSAL = '" & ListView4.ListItems(Con).SubItems(2) & "' AND ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                                cnn.Execute (sBuscar)
                            End If
                            If Check2.value = 1 Then
                                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & ListView4.ListItems(Con).SubItems(1) & " WHERE SUCURSAL = '" & Combo1.Text & "' AND ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                                cnn.Execute (sBuscar)
                            End If
                        End If
                    Else
                        If Check1.value = 1 Then
                            sBuscar = "INSERT INTO EXISTENCIASTEMPORALES (ID_PRODUCTO, CANTIDAD,SUCURSAL) VALUES ('" & ListView4.ListItems(Con) & "', '" & ListView4.ListItems(Con).SubItems(1) & "', '" & ListView4.ListItems(Con).SubItems(2) & "');"
                            cnn.Execute (sBuscar)
                        End If
                        If Check2.value = 1 Then
                            sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD,SUCURSAL) VALUES ('" & ListView4.ListItems(Con) & "', '" & ListView4.ListItems(Con).SubItems(1) & "', '" & Combo1.Text & "');"
                            cnn.Execute (sBuscar)
                        End If
                    End If
                Else
                    Set tLi = ListView1.ListItems.Add(, , ListView4.ListItems(Con))
                    tLi.SubItems(1) = ListView4.ListItems(Con).SubItems(1)
                End If
            End If
            If Option5.value Then
                sBuscar = "SELECT ID_PRODUCTO FROM ALMACEN2 WHERE ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    sBuscar = "select ID_PRODUCTO, CANTIDAD,SUCURSAL from vsinvalm2 WHERE SUCURSAL = '" & Combo1.Text & "' AND ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        If tRs.Fields("CANTIDAD") = 0 Then
                            If Check1.value = 1 Then
                                sBuscar = "INSERT INTO EXISTENCIASTEMPORALES (ID_PRODUCTO, CANTIDAD,SUCURSAL) VALUES ('" & ListView4.ListItems(Con) & "', '" & ListView4.ListItems(Con).SubItems(1) & "', '" & ListView4.ListItems(Con).SubItems(2) & "');"
                                cnn.Execute (sBuscar)
                            End If
                            If Check2.value = 1 Then
                                sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD,SUCURSAL) VALUES ('" & ListView4.ListItems(Con) & "', '" & ListView4.ListItems(Con).SubItems(1) & "', '" & Combo1.Text & "');"
                                cnn.Execute (sBuscar)
                            End If
                        Else
                            If Check1.value = 1 Then
                                sBuscar = "UPDATE EXISTENCIASTEMPORALES SET CANTIDAD = " & ListView4.ListItems(Con).SubItems(1) & " WHERE SUCURSAL = '" & ListView4.ListItems(Con).SubItems(2) & "' AND ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                                cnn.Execute (sBuscar)
                            End If
                            If Check2.value = 1 Then
                                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & ListView4.ListItems(Con).SubItems(1) & " WHERE SUCURSAL = '" & Combo1.Text & "' AND ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                                cnn.Execute (sBuscar)
                            End If
                        End If
                    Else
                        If Check1.value = 1 Then
                            sBuscar = "INSERT INTO EXISTENCIASTEMPORALES (ID_PRODUCTO, CANTIDAD,SUCURSAL) VALUES ('" & ListView4.ListItems(Con) & "', '" & ListView4.ListItems(Con).SubItems(1) & "', '" & ListView4.ListItems(Con).SubItems(2) & "');"
                            cnn.Execute (sBuscar)
                        End If
                        If Check2.value = 1 Then
                            sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD,SUCURSAL) VALUES ('" & ListView4.ListItems(Con) & "', '" & ListView4.ListItems(Con).SubItems(1) & "', '" & Combo1.Text & "');"
                            cnn.Execute (sBuscar)
                        End If
                    End If
                Else
                    Set tLi = ListView1.ListItems.Add(, , ListView4.ListItems(Con))
                    tLi.SubItems(1) = ListView4.ListItems(Con).SubItems(1)
                End If
            End If
            If Option6.value Then
                sBuscar = "SELECT ID_PRODUCTO FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    sBuscar = "select ID_PRODUCTO, CANTIDAD,SUCURSAL from vsinvalm3 WHERE SUCURSAL = '" & Combo1.Text & "' AND ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                    Set tRs = cnn.Execute(sBuscar)
                    If Not (tRs.EOF And tRs.BOF) Then
                        If tRs.Fields("CANTIDAD") = 0 Then
                            If Check1.value = 1 Then
                                sBuscar = "INSERT INTO EXISTENCIASTEMPORALES (ID_PRODUCTO, CANTIDAD,SUCURSAL) VALUES ('" & ListView4.ListItems(Con) & "', '" & ListView4.ListItems(Con).SubItems(1) & "', '" & ListView4.ListItems(Con).SubItems(2) & "');"
                                cnn.Execute (sBuscar)
                            End If
                            If Check2.value = 1 Then
                                sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD,SUCURSAL) VALUES ('" & ListView4.ListItems(Con) & "', '" & ListView4.ListItems(Con).SubItems(1) & "', '" & Combo1.Text & "');"
                                cnn.Execute (sBuscar)
                            End If
                        Else
                            If Check1.value = 1 Then
                                sBuscar = "UPDATE EXISTENCIASTEMPORALES SET CANTIDAD = " & ListView4.ListItems(Con).SubItems(1) & " WHERE SUCURSAL = '" & ListView4.ListItems(Con).SubItems(2) & "' AND ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                                cnn.Execute (sBuscar)
                            End If
                            If Check2.value = 1 Then
                                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & ListView4.ListItems(Con).SubItems(1) & " WHERE SUCURSAL = '" & Combo1.Text & "' AND ID_PRODUCTO = '" & ListView4.ListItems(Con) & "'"
                                cnn.Execute (sBuscar)
                            End If
                        End If
                    Else
                        If Check1.value = 1 Then
                            sBuscar = "INSERT INTO EXISTENCIASTEMPORALES (ID_PRODUCTO, CANTIDAD,SUCURSAL) VALUES ('" & ListView4.ListItems(Con) & "', '" & ListView4.ListItems(Con).SubItems(1) & "', '" & ListView4.ListItems(Con).SubItems(2) & "');"
                            cnn.Execute (sBuscar)
                        End If
                        If Check2.value = 1 Then
                            sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD,SUCURSAL) VALUES ('" & ListView4.ListItems(Con) & "', '" & ListView4.ListItems(Con).SubItems(1) & "', '" & Combo1.Text & "');"
                            cnn.Execute (sBuscar)
                        End If
                    End If
                Else
                    Set tLi = ListView1.ListItems.Add(, , ListView4.ListItems(Con))
                    tLi.SubItems(1) = ListView4.ListItems(Con).SubItems(1)
                End If
            End If
        Next Con
        Me.Height = 7695
        Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
        MsgBox "La importación ha finalizado exitosamente!", vbExclamation, "SACC"
    Else
        MsgBox "No se ha seleccionado ningun archivo a importar!", vbExclamation, "SACC"
    End If
End Sub
Private Sub Command4_Click()
On Error GoTo Maneja
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Abrir"
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "[Archivo Excel (*.xls)] |*.xls|"
    CommonDialog1.ShowOpen
    Text2.Text = CommonDialog1.FileName
    Call Excel_FlexGrid(CommonDialog1.FileName, ListView4, 20, 5, "Hoja1")
Maneja:
    Err.Clear
End Sub
Private Sub Command5_Click()
    On Error GoTo Maneja
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Abrir"
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "[Archivo Excel (*.xls)] |*.xls|"
    CommonDialog1.ShowOpen
    Text3.Text = CommonDialog1.FileName
    Call Excel_FlexGrid2(CommonDialog1.FileName, ListView5, 20, 5, "Hoja1")
Maneja:
    Err.Clear
End Sub
Private Sub Command6_Click()
    Dim tRs As ADODB.Recordset
    If Text3.Text <> "" Then
        Dim sBuscar As String
        Dim tLi As ListItem
        Dim Con As Double
        For Con = 1 To ListView5.ListItems.COUNT
            If Option11.value Then
                sBuscar = "SELECT ID_PRODUCTO FROM ALMACEN1 WHERE ID_PRODUCTO = '" & ListView5.ListItems(Con) & "'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    sBuscar = "UPDATE ALMACEN1 SET CATEGORIA = '" & ListView5.ListItems(Con).SubItems(1) & "', PRESENTACION = '" & ListView5.ListItems(Con).SubItems(2) & "' WHERE ID_PRODUCTO = '" & ListView5.ListItems(Con) & "'"
                    cnn.Execute (sBuscar)
                End If
            End If
            If Option10.value Then
                sBuscar = "SELECT ID_PRODUCTO FROM ALMACEN2 WHERE ID_PRODUCTO = '" & ListView5.ListItems(Con) & "'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    sBuscar = "UPDATE ALMACEN2 SET CATEGORIA = '" & ListView5.ListItems(Con).SubItems(1) & "', PRESENTACION = '" & ListView5.ListItems(Con).SubItems(2) & "' WHERE ID_PRODUCTO = '" & ListView5.ListItems(Con) & "'"
                    cnn.Execute (sBuscar)
                End If
            End If
            If Option9.value Then
                sBuscar = "SELECT ID_PRODUCTO FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ListView5.ListItems(Con) & "'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    sBuscar = "UPDATE ALMACEN3 SET CATEGORIA = '" & ListView5.ListItems(Con).SubItems(1) & "', PRESENTACION = '" & ListView5.ListItems(Con).SubItems(2) & "' WHERE ID_PRODUCTO = '" & ListView5.ListItems(Con) & "'"
                    cnn.Execute (sBuscar)
                End If
            End If
        Next Con
        MsgBox "La importación ha finalizado exitosamente!", vbExclamation, "SACC"
    Else
        MsgBox "No se ha seleccionado ningun archivo a importar!", vbExclamation, "SACC"
    End If
End Sub
Private Sub Form_Load()
    Dim tRs As ADODB.Recordset
    Dim sBu As String
    Me.Height = 4230
    Set obj_Workbook = Nothing
    Set obj_Excel = Nothing
    Set obj_Worksheet = Nothing
    Me.MousePointer = vbDefault
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    sBu = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBu)
        If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("NOMBRE")
             tRs.MoveNext
        Loop
    End If
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 2000
        .ColumnHeaders.Add , , "CANTIDAD", 2000
        .ColumnHeaders.Add , , "SUCURSAL", 2000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 2000
        .ColumnHeaders.Add , , "PRECIO NUEVO", 2000
        .ColumnHeaders.Add , , "RECIO ANTERIOR", 2000
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 2000
        .ColumnHeaders.Add , , "PRECIO NUEVO", 2000
    End With
End Sub
Private Sub Image1_Click()
On Error GoTo Maneja
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Dim foo As Integer
    Dim Ruta1 As String
    Dim foo1 As Integer
    If ListView2.ListItems.COUNT > 0 Then
        Me.CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "Guardar como"
        CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
        Me.CommonDialog1.ShowSave
        Ruta = Me.CommonDialog1.FileName
        If Ruta <> "" Then
            NumColum = ListView2.ColumnHeaders.COUNT
            For Con = 1 To ListView2.ColumnHeaders.COUNT
                StrCopi = StrCopi & ListView2.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView2.ListItems.COUNT
                StrCopi = StrCopi & ListView2.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView2.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            'archivo TXT
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
    StrCopi = ""
    If ListView3.ListItems.COUNT > 0 Then
        Me.CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "Guardar como"
        CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
        Me.CommonDialog1.ShowSave
        Ruta1 = Me.CommonDialog1.FileName
        If Ruta1 <> "" Then
            NumColum = ListView3.ColumnHeaders.COUNT
            For Con = 1 To ListView3.ColumnHeaders.COUNT
                StrCopi = StrCopi & ListView3.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView3.ListItems.COUNT
                StrCopi = StrCopi & ListView3.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView3.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            'archivo TXt
            foo1 = FreeFile
            Open Ruta1 For Output As #foo1
                Print #foo1, StrCopi
            Close #foo1
        End If
        ShellExecute Me.hWnd, "open", Ruta1, "", "", 4
    End If
Maneja:
    Err.Clear
End Sub
Private Sub Image10_Click()
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
    If ListView1.ListItems.COUNT > 0 Then
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.COUNT
            For Con = 1 To ListView1.ColumnHeaders.COUNT
                StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
            Next
            ProgressBar1.value = 0
            ProgressBar1.Visible = True
            ProgressBar1.Min = 0
            ProgressBar1.Max = ListView1.ListItems.COUNT
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView1.ListItems.COUNT
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
                ProgressBar1.value = Con
            Next
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ProgressBar1.Visible = False
        ProgressBar1.value = 0
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        ListView1.Visible = False
        ListView2.Visible = True
        ListView3.Visible = True
        Frame11.Visible = False
        Frame3.Visible = True
        Label5.Visible = True
    Else
        ListView1.Visible = True
        ListView2.Visible = False
        ListView3.Visible = False
        Frame11.Visible = True
        Frame3.Visible = False
        Label5.Visible = False
    End If
End Sub
Private Sub Excel_FlexGrid(sPath As String, FlexGrid As Object, Filas As Integer, Columnas As Integer, Optional sSheetName As String = vbNullString)
    Dim i As Long
    Dim n As Long
    Dim r As Long
    Dim tLi As ListItem
    n = 1
    On Error GoTo error_sub
    If Len(Dir(sPath)) = 0 Then
       MsgBox "No se ha encontrado el archivo: " & sPath, vbCritical
       Exit Sub
    End If
    Me.MousePointer = vbHourglass
    Set obj_Excel = CreateObject("Excel.Application")
    Set obj_Workbook = obj_Excel.Workbooks.Open(sPath)
    'If sSheetName = vbNullString Then
        Set obj_Worksheet = obj_Workbook.ActiveSheet
    'Else
    '    Set obj_Worksheet = obj_Workbook.Sheets(sSheetName)
    'End If
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        Do While obj_Worksheet.Cells(1, n).value <> ""
            .ColumnHeaders.Add , , "No. " & n, 1000
            n = n + 1
        Loop
    End With
    n = n - 1
    i = 1
    ListView4.ListItems.Clear
    Do While obj_Worksheet.Cells(i, 1).value <> ""
        Set tLi = ListView4.ListItems.Add(, , obj_Worksheet.Cells(i, 1).value)
        If n > 1 Then
            For r = 2 To n
                tLi.SubItems(r - 1) = obj_Worksheet.Cells(i, r).value
            Next r
        End If
        i = i + 1
    Loop
    obj_Workbook.Close
    obj_Excel.Quit
    Set obj_Workbook = Nothing
    Set obj_Excel = Nothing
    Set obj_Worksheet = Nothing
    Me.MousePointer = vbDefault
Exit Sub
error_sub:
    MsgBox Err.Description
    Set obj_Workbook = Nothing
    Set obj_Excel = Nothing
    Set obj_Worksheet = Nothing
    Me.MousePointer = vbDefault
    Me.MousePointer = vbDefault
End Sub
Private Sub Excel_FlexGrid2(sPath As String, FlexGrid As Object, Filas As Integer, Columnas As Integer, Optional sSheetName As String = vbNullString)
    Dim i As Long
    Dim n As Long
    Dim r As Long
    Dim tLi As ListItem
    n = 1
    On Error GoTo error_sub
    If Len(Dir(sPath)) = 0 Then
       MsgBox "No se ha encontrado el archivo: " & sPath, vbCritical
       Exit Sub
    End If
    Me.MousePointer = vbHourglass
    Set obj_Excel = CreateObject("Excel.Application")
    Set obj_Workbook = obj_Excel.Workbooks.Open(sPath)
    'If sSheetName = vbNullString Then
        Set obj_Worksheet = obj_Workbook.ActiveSheet
    'Else
    '    Set obj_Worksheet = obj_Workbook.Sheets(sSheetName)
    'End If
    With ListView5
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        Do While obj_Worksheet.Cells(1, n).value <> ""
            .ColumnHeaders.Add , , "No. " & n, 1000
            n = n + 1
        Loop
    End With
    n = n - 1
    i = 1
    ListView5.ListItems.Clear
    Do While obj_Worksheet.Cells(i, 1).value <> ""
        Set tLi = ListView5.ListItems.Add(, , obj_Worksheet.Cells(i, 1).value)
        If n > 1 Then
            For r = 2 To n
                tLi.SubItems(r - 1) = obj_Worksheet.Cells(i, r).value
            Next r
        End If
        i = i + 1
    Loop
    obj_Workbook.Close
    obj_Excel.Quit
    Set obj_Workbook = Nothing
    Set obj_Excel = Nothing
    Set obj_Worksheet = Nothing
    Me.MousePointer = vbDefault
Exit Sub
error_sub:
    MsgBox Err.Description
    Set obj_Workbook = Nothing
    Set obj_Excel = Nothing
    Set obj_Worksheet = Nothing
    Me.MousePointer = vbDefault
    Me.MousePointer = vbDefault
End Sub
