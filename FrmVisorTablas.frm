VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmVisorTablas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visor de Tablas"
   ClientHeight    =   6855
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   12135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11280
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Filtro Escrito"
      Height          =   195
      Left            =   9360
      TabIndex        =   46
      ToolTipText     =   "Seleccione esta opción para escribir usted mismo sus instrucciones SQL a ejecutar"
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Filtro Ayuda"
      Height          =   195
      Left            =   10680
      TabIndex        =   45
      ToolTipText     =   "Seleccione esta opción para ejecutar las acciones del modo de asistente"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11040
      TabIndex        =   43
      Top             =   2880
      Width           =   975
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   44
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image14 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmVisorTablas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmVisorTablas.frx":030A
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11040
      TabIndex        =   4
      Top             =   4080
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmVisorTablas.frx":1F0C
         MousePointer    =   99  'Custom
         Picture         =   "FrmVisorTablas.frx":2216
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   6588
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4471
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Sentencia SQL"
      TabPicture(0)   =   "FrmVisorTablas.frx":3D58
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Text1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Opciones de Vista"
      TabPicture(1)   =   "FrmVisorTablas.frx":3D74
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5(2)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Filtro de Cosulta"
      TabPicture(2)   =   "FrmVisorTablas.frx":3D90
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5(1)"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   1815
         Index           =   1
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   11655
         Begin VB.Frame Frame4 
            Caption         =   "Opciones"
            Height          =   1695
            Left            =   9960
            TabIndex        =   38
            Top             =   120
            Width           =   1575
            Begin VB.CheckBox Check1 
               Caption         =   "Agrupar"
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   480
               Width           =   975
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Orden Descendiente"
               Height          =   435
               Left            =   120
               TabIndex        =   39
               Top             =   960
               Width           =   1335
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Campos en Vista"
            Height          =   1695
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   3255
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   2175
            End
            Begin VB.ListBox List1 
               Height          =   1035
               Left            =   120
               TabIndex        =   36
               Top             =   600
               Width           =   3015
            End
            Begin VB.CommandButton Command2 
               BackColor       =   &H00E0E0E0&
               Caption         =   "+"
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
               Left            =   2760
               TabIndex        =   35
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton Command3 
               BackColor       =   &H00E0E0E0&
               Caption         =   "-"
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
               Left            =   2400
               TabIndex        =   34
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Orden en Vista"
            Height          =   1695
            Left            =   6600
            TabIndex        =   28
            Top             =   120
            Width           =   3255
            Begin VB.CommandButton Command4 
               BackColor       =   &H00E0E0E0&
               Caption         =   "-"
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
               Left            =   2400
               TabIndex        =   32
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton Command5 
               BackColor       =   &H00E0E0E0&
               Caption         =   "+"
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
               Left            =   2760
               TabIndex        =   31
               Top             =   240
               Width           =   375
            End
            Begin VB.ListBox List2 
               Height          =   1035
               Left            =   120
               TabIndex        =   30
               Top             =   600
               Width           =   3015
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Sumar en Vista"
            Height          =   1695
            Left            =   3360
            TabIndex        =   23
            Top             =   120
            Width           =   3255
            Begin VB.CommandButton Command6 
               BackColor       =   &H00E0E0E0&
               Caption         =   "-"
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
               Left            =   2400
               TabIndex        =   27
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton Command7 
               BackColor       =   &H00E0E0E0&
               Caption         =   "+"
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
               Left            =   2760
               TabIndex        =   26
               Top             =   240
               Width           =   375
            End
            Begin VB.ListBox List3 
               Height          =   1035
               Left            =   120
               TabIndex        =   25
               Top             =   600
               Width           =   3015
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Width           =   2175
            End
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   1815
         Index           =   2
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   11655
         Begin VB.ComboBox Combo7 
            Height          =   315
            Left            =   1560
            TabIndex        =   41
            Top             =   120
            Width           =   3615
         End
         Begin VB.CheckBox Check4 
            Caption         =   "O"
            Height          =   255
            Left            =   5520
            TabIndex        =   16
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Y"
            Height          =   255
            Left            =   5520
            TabIndex        =   15
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton Command9 
            BackColor       =   &H00E0E0E0&
            Caption         =   "+"
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
            Left            =   5760
            TabIndex        =   14
            Top             =   1200
            Width           =   375
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "-"
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
            Left            =   5400
            TabIndex        =   13
            Top             =   1200
            Width           =   375
         End
         Begin VB.ListBox List4 
            Height          =   1425
            Left            =   6240
            TabIndex        =   12
            Top             =   240
            Width           =   5295
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   1560
            TabIndex        =   11
            Top             =   840
            Width           =   3615
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   1560
            TabIndex        =   10
            Top             =   480
            Width           =   3615
         End
         Begin VB.ComboBox Combo6 
            Height          =   315
            Left            =   1560
            TabIndex        =   9
            Top             =   1200
            Width           =   3615
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   3720
            TabIndex        =   17
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50462721
            CurrentDate     =   39531
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1560
            TabIndex        =   18
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50462721
            CurrentDate     =   39531
         End
         Begin VB.Label Label4 
            Caption         =   "Tablas de la BD"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Valor a comparar"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Validación"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Campo de la BD"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "FrmVisorTablas.frx":3DAC
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   1845
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         ToolTipText     =   "Escriba la sentencia SQL para ejecutar"
         Top             =   480
         Width           =   11655
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11040
      TabIndex        =   0
      Top             =   5280
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
         MouseIcon       =   "FrmVisorTablas.frx":3DB2
         MousePointer    =   99  'Custom
         Picture         =   "FrmVisorTablas.frx":40BC
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Menu MENUSOTE 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu PopMenCopiar 
         Caption         =   "Copiar"
      End
      Begin VB.Menu SubMenCorta 
         Caption         =   "Cortar"
      End
      Begin VB.Menu SubMenEliSel 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu a5 
         Caption         =   "-"
      End
      Begin VB.Menu SubMenSelTodo 
         Caption         =   "Seleccionar Todo"
      End
   End
   Begin VB.Menu MenArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu SbuMenExcel 
         Caption         =   "Enviar Selección a Excel"
      End
      Begin VB.Menu SubMenEnviTabExcel 
         Caption         =   "Enviar Tabla a Excel"
      End
      Begin VB.Menu SubMenSelArchTxt 
         Caption         =   "Enviar Seleccion a Archivo de Texto"
      End
      Begin VB.Menu SubMenTablaArchTxt 
         Caption         =   "Enviar Tabla a Archivo de Texto"
      End
   End
   Begin VB.Menu MenEdicion 
      Caption         =   "Edición"
      Begin VB.Menu SubMenCopiar 
         Caption         =   "Copiar"
      End
      Begin VB.Menu SubMenCortar 
         Caption         =   "Cortar"
      End
      Begin VB.Menu SubMenEliSel2 
         Caption         =   "Eliminar "
      End
      Begin VB.Menu a6 
         Caption         =   "-"
      End
      Begin VB.Menu SubMenSelecTodo 
         Caption         =   "Seleccionar Todo"
      End
   End
   Begin VB.Menu MenAvanzada 
      Caption         =   "Avanzada"
      Begin VB.Menu SubMenRestablecer 
         Caption         =   "Restablecer Base de Datos en Blanco"
      End
   End
End
Attribute VB_Name = "FrmVisorTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************************************************************
'                           Programado por :  Ing. Armando H Valdez Arras
'                           Fecha Inicio:     03 de Noviembre de 2011
'                           Version :         V 1.0
'****************************************************************************************************************************************
Option Explicit
Private cnn As ADODB.Connection
Dim VarEl1 As Integer
Dim VarEl2 As Integer
Dim VarEl3 As Integer
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo3_DropDown()
    Dim Con As Integer
    Combo3.Clear
    For Con = 0 To List1.ListCount - 1
        Combo3.AddItem List1.List(Con)
    Next
End Sub
Private Sub Combo3_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo5_DropDown()
    If Combo4.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim sBuscar As String
        sBuscar = ""
        sBuscar = Replace(Combo7.Text, "(BASE TABLE)", "")
        sBuscar = Replace(sBuscar, "(VIEW)", "")
        sBuscar = "SELECT " & Combo4.Text & " FROM " & sBuscar
        Set tRs = cnn.Execute(sBuscar)
        Combo5.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            'tRs.Fields(Combo4.Text).Type = adBigInt
            'tRs.Fields(Combo4.Text).Type = adBinary
            'tRs.Fields(Combo4.Text).Type = adBoolean
            'tRs.Fields(Combo4.Text).Type = adBSTR
            'tRs.Fields(Combo4.Text).Type = adChapter
            'tRs.Fields(Combo4.Text).Type = adChar
            'tRs.Fields(Combo4.Text).Type = adCurrency
            'tRs.Fields(Combo4.Text).Type = adDate
            'tRs.Fields(Combo4.Text).Type = adDBDate
            'tRs.Fields(Combo4.Text).Type = adDBFileTime
            'tRs.Fields(Combo4.Text).Type = adDBTime
            'tRs.Fields(Combo4.Text).Type = adDBTimeStamp
            'tRs.Fields(Combo4.Text).Type = adDecimal
            'tRs.Fields(Combo4.Text).Type = adDouble
            'tRs.Fields(Combo4.Text).Type = adEmpty
            'tRs.Fields(Combo4.Text).Type = adError
            'tRs.Fields(Combo4.Text).Type = adFileTime
            'tRs.Fields(Combo4.Text).Type = adGUID
            'tRs.Fields(Combo4.Text).Type = adIDispatch
            'tRs.Fields(Combo4.Text).Type = adInteger
            'tRs.Fields(Combo4.Text).Type = adIUnknown
            'tRs.Fields(Combo4.Text).Type = adLongVarBinary
            'tRs.Fields(Combo4.Text).Type = adLongVarChar
            'tRs.Fields(Combo4.Text).Type = adLongVarWChar
            'tRs.Fields(Combo4.Text).Type = adNumeric
            'tRs.Fields(Combo4.Text).Type = adPropVariant
            'tRs.Fields(Combo4.Text).Type = adSingle
            'tRs.Fields(Combo4.Text).Type = adSmallInt
            'tRs.Fields(Combo4.Text).Type = adTinyInt
            'tRs.Fields(Combo4.Text).Type = adUnsignedBigInt
            'tRs.Fields(Combo4.Text).Type = adUnsignedInt
            'tRs.Fields(Combo4.Text).Type = adUnsignedSmallInt
            'tRs.Fields(Combo4.Text).Type = adUnsignedTinyInt
            'tRs.Fields(Combo4.Text).Type = adUserDefined
            'tRs.Fields(Combo4.Text).Type = adVarBinary
            'tRs.Fields(Combo4.Text).Type = adVarChar
            'tRs.Fields(Combo4.Text).Type = adVariant
            'tRs.Fields(Combo4.Text).Type = adVarNumeric
            'tRs.Fields(Combo4.Text).Type = adVarWChar
            'tRs.Fields(Combo4.Text).Type = adWChar
            Combo5.AddItem "MAYOR QUE"
            Combo5.AddItem "MENOR QUE"
            Combo5.AddItem "IGUAL A (EXACTO)"
            Combo5.AddItem "DIFERENTE A (EXACTO)"
            If tRs.Fields(Combo4.Text).Type = adChar Or tRs.Fields(Combo4.Text).Type = adLongVarChar Or tRs.Fields(Combo4.Text).Type = adLongVarWChar Or tRs.Fields(Combo4.Text).Type = adVarChar Or tRs.Fields(Combo4.Text).Type = adVarWChar Or tRs.Fields(Combo4.Text).Type = adWChar Then
                Combo5.AddItem "QUE CONTENGA (PARTE DE CADENA)"
                Combo5.AddItem "QUE NO CONTENGA (PARTE DE CADENA)"
            End If
            If tRs.Fields(Combo4.Text).Type = adDate Or tRs.Fields(Combo4.Text).Type = adDBDate Then
                Combo5.AddItem "ENTRE FECHA Y FECHA"
            End If
        End If
    End If
End Sub
Private Sub Combo5_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo5_LostFocus()
    If Combo5.Text = "ENTRE FECHA Y FECHA" Then
        Combo6.Visible = False
        DTPicker1.Value = Date - 7
        DTPicker2.Value = Date
        DTPicker1.Visible = True
        DTPicker2.Visible = True
    Else
        Combo6.Visible = True
        DTPicker1.Visible = False
        DTPicker2.Visible = False
    End If
End Sub
Private Sub Combo6_DropDown()
    If Combo4.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim Con As Integer
        Dim sBuscar As String
        sBuscar = Replace(Combo7.Text, "(BASE TABLE)", "")
        sBuscar = Replace(sBuscar, "(VIEW)", "")
        sBuscar = "SELECT * FROM " & sBuscar
        Set tRs = cnn.Execute(sBuscar)
        Combo6.Clear
        For Con = 0 To tRs.Fields.Count - 1
            If tRs.Fields(tRs.Fields.Item(Con).Name).Type = tRs.Fields(Combo4.Text).Type Then
                Combo6.AddItem tRs.Fields.Item(Con).Name
            End If
        Next
        If tRs.Fields(Combo4.Text).Type = adBoolean Then
            Combo6.Clear
            Combo6.AddItem "VERDADERO"
            Combo6.AddItem "FALSO"
        End If
    End If
End Sub
Private Sub Combo7_LostFocus()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim Con As Integer
    If Combo7.Text <> "" Then
        sBuscar = Replace(Combo7.Text, "(BASE TABLE)", "")
        sBuscar = Replace(sBuscar, "(VIEW)", "")
        sBuscar = "SELECT * FROM " & sBuscar
        Set tRs = cnn.Execute(sBuscar)
        ListView1.ListItems.Clear
        ListView1.ColumnHeaders.Clear
        With ListView1
            .View = lvwReport
            .Gridlines = True
            .LabelEdit = lvwManual
            .HideSelection = False
            .HotTracking = False
            .FullRowSelect = True
            .HoverSelection = False
            For Con = 0 To tRs.Fields.Count - 1
                .ColumnHeaders.Add , , tRs.Fields.Item(Con).Name & " ( Tamaño :" & tRs.Fields.Item(Con).DefinedSize & ")", 1500
                Combo1.AddItem tRs.Fields.Item(Con).Name
                Combo2.AddItem tRs.Fields.Item(Con).Name
                Combo4.AddItem tRs.Fields.Item(Con).Name
            Next
        End With
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields(tRs.Fields.Item(0).Name))
            For Con = 1 To tRs.Fields.Count - 1
                If Not IsNull(tRs.Fields(tRs.Fields.Item(Con).Name)) Then tLi.SubItems(Con) = tRs.Fields(tRs.Fields.Item(Con).Name)
            Next
            tRs.MoveNext
        Loop
        Text1.Text = sBuscar
    End If
End Sub
Private Sub Combo7_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo4_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Command2_Click()
    Dim Con As Integer
    Dim Esta As Integer
    Esta = 0
    If Combo1.Text <> "" Then
        If Combo1.Text = "<TODOS>" Then
            List1.Clear
            List1.AddItem "<TODOS>"
        Else
            If List1.ListCount > 0 Then
                For Con = 0 To List1.ListCount
                    If List1.List(Con) = Combo1.Text Then
                        Esta = 1
                    End If
                    If List1.List(Con) = "<TODOS>" Then
                        List1.Clear
                        List1.AddItem "<TODOS>"
                        Esta = 1
                    End If
                Next
                For Con = 0 To List3.ListCount
                    If List3.List(Con) = Combo1.Text Then
                        Esta = 1
                    End If
                Next
                If Esta = 0 Then
                    List1.AddItem Combo1.Text
                End If
            Else
                For Con = 0 To List3.ListCount
                    If List3.List(Con) = Combo1.Text Then
                        Esta = 1
                    End If
                Next
                If Esta = 0 Then
                    List1.AddItem Combo1.Text
                End If
            End If
        End If
    End If
End Sub
Private Sub Command3_Click()
    If VarEl1 <> "" Then
        List1.RemoveItem (VarEl1)
        VarEl1 = ""
    End If
End Sub
Private Sub Command4_Click()
    If VarEl2 <> "" Then
        List2.RemoveItem (VarEl2)
        VarEl2 = ""
    End If
End Sub
Private Sub Command5_Click()
    Dim Con As Integer
    Dim Esta As Integer
    Esta = 0
    If Combo2.Text <> "" Then
        If List2.ListCount > 0 Then
            For Con = 0 To List2.ListCount
                If List2.List(Con) = Combo2.Text Then
                    Esta = 1
                End If
            Next
            If Esta = 0 Then
                List2.AddItem Combo2.Text
            End If
        Else
            List2.AddItem Combo2.Text
        End If
    End If
End Sub
Private Sub Command6_Click()
    If VarEl3 <> "" Then
        List3.RemoveItem (VarEl3)
        VarEl3 = ""
    End If
End Sub
Private Sub Command7_Click()
    Dim Con As Integer
    Dim Esta As Integer
    Esta = 0
    If Combo3.Text <> "" Then
        If List3.ListCount > 0 Then
            For Con = 0 To List3.ListCount
                If List3.List(Con) = Combo3.Text Then
                    Esta = 1
                End If
            Next
            For Con = 0 To List1.ListCount
                If List1.List(Con) = Combo3.Text Then
                    List1.RemoveItem (Con)
                End If
            Next
            If Esta = 0 Then
                List3.AddItem Combo3.Text
            End If
        Else
            List3.AddItem Combo3.Text
            For Con = 0 To List1.ListCount
                If List1.List(Con) = Combo3.Text Then
                    List1.RemoveItem (Con)
                End If
            Next
        End If
    End If
    If List3.ListCount > 0 Then
        Check1.Value = 1
    End If
End Sub
Private Sub Form_Load()
    Dim Con As Integer
    Dim PPosi As Integer
    Dim Tabla As String
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim NUMREG As Double
    Dim tLi As ListItem
    Dim i As Integer
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    DTPicker1.Visible = False
    DTPicker2.Visible = False
    SubMenCopiar.Enabled = False
    SubMenCortar.Enabled = False
    SubMenEliSel2.Enabled = False
    SubMenSelecTodo.Enabled = False
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    sBuscar = "SELECT * from Information_Schema.Tables ORDER BY TABLE_TYPE, TABLE_NAME"
    Set tRs = cnn.Execute(sBuscar)
    Frame5(1).Enabled = True
    Frame5(2).Enabled = True
    Combo1.Clear
    Combo2.Clear
    Combo3.Clear
    Combo4.Clear
    Combo5.Clear
    Combo6.Clear
    Combo1.AddItem "<TODOS>"
    ListView1.Visible = True
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    ListView1.ColumnHeaders.Clear
    SubMenCopiar.Enabled = True
    SubMenCortar.Enabled = True
    SubMenEliSel2.Enabled = True
    SubMenSelecTodo.Enabled = True
    Do While Not tRs.EOF
        Combo7.AddItem tRs.Fields("TABLE_NAME") & " (" & tRs.Fields("TABLE_TYPE") & ")"
        tRs.MoveNext
    Loop
Exit Sub
ManejaError:
    If Err.Number = -2147217887 Then
        Err.Clear
        MsgBox "ABIERTO CON ERRORES", vbCritical, "SACC"
    Else
        If Err.Number = -2147217865 Then
            MsgBox "EL NOMBRE DE LA TABLA NO EXISTE O EL ARCHIVO NO ES UN DBF", vbCritical, "SACC"
        Else
            If Err.Number <> -2147467259 Then
                If Err.Number = -2147352571 Then
                    MsgBox "OCURRIO UN ERROR AL ABRIR LA TABLA, ES POSIBLE QUE NO SE MUESTRE TODA LA INFORMACIÓN DE ESTA, PUEDE FILTRAR LA INFORMACIÓN PARA SECCIONARLA", vbCritical, "SACC"
                Else
                    If Err.Number <> 0 Then
                        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
                    End If
                End If
            End If
        End If
    End If
    Err.Clear
End Sub
Private Sub Image10_Click()
    If ListView1.ListItems.Count > 0 Then
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
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.Count
                For Con = 1 To ListView1.ColumnHeaders.Count - 1
                    StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
                Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView1.ListItems.Count
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            Text2.Text = StrCopi
            'archivo TXT
            Dim foo As Integer
        
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, Text2.Text
            Close #foo
        End If
    End If
End Sub
Private Sub Image14_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim OrdBy As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim Con As Integer
    Dim CampSum As String
    Dim CampAgrup As String
    Dim SeWhe As String
    Dim NoReg As Integer
    If Option1.Value Then
        If Text1.Text <> "" Then
            sBuscar = Text1.Text
        Else
            MsgBox "NO SE HA SELECCIONADO NINGUNA BUSQUEDA", vbExclamation, "SACC"
        End If
    Else
        If List1.ListCount = 0 And List3.ListCount = 0 Then
            MsgBox "NO SE HA SELECCIONADO NINGUNA BUSQUEDA", vbExclamation, "SACC"
        Else
            ' Sentencia de campos en el SUM()
            For Con = 0 To List3.ListCount - 1
                CampSum = CampSum & "SUM (" & List3.List(Con) & ") AS SUM_" & List3.List(Con)
                If Con < List3.ListCount - 1 Then
                    CampSum = CampSum & ", "
                End If
            Next
            ' Sentencia desde el SELECT hasta el nombre de los campos seleccionados
            sBuscar = "SELECT "
            If List1.ListCount > 0 Then
                For Con = 0 To List1.ListCount - 1
                    If List1.List(Con) = "<TODOS>" Then
                        sBuscar = "SELECT * "
                    Else
                        sBuscar = sBuscar & List1.List(Con)
                        If Con < List1.ListCount - 1 Then
                            sBuscar = sBuscar & ", "
                        End If
                    End If
                Next
                If CampSum <> "" Then
                    sBuscar = sBuscar & ", "
                End If
                sBuscar = sBuscar & CampSum & " FROM " & Combo7.Text
            Else
                If CampSum <> "" Then
                    If List1.List(1) = "<TODOS>" Then
                        MsgBox "LA SUMATORIA NO SE REALIZARA YA QUE AGREGO LA OPCION DE TODOS LOS CAMPOS, PARA REALIZAR UNA SUMATORIA DEBE SELECCIONAR CAMPO POR CAMPO", vbExclamation, "Hache's system"
                        Check1.Value = 0
                        List3.Clear
                    Else
                        sBuscar = sBuscar & CampSum & " FROM " & Combo7.Text
                    End If
                Else
                    sBuscar = sBuscar & CampSum & " FROM " & Combo7.Text
                End If
            End If
            'Sentencia del WHERE
            If List4.ListCount > 0 Then
                SeWhe = " WHERE "
                For Con = 0 To List4.ListCount - 1
                    SeWhe = SeWhe & List4.List(Con)
                    If Con < List4.ListCount - 1 Then
                        SeWhe = SeWhe & ", "
                    End If
                Next
            End If
            sBuscar = sBuscar & SeWhe
            'Sentencia del ORDER BY
            If List2.ListCount > 0 Then
                OrdBy = " ORDER BY "
                For Con = 0 To List2.ListCount - 1
                    OrdBy = OrdBy & List2.List(Con)
                    If Con < List2.ListCount - 1 Then
                        OrdBy = OrdBy & ", "
                    End If
                Next
                If Check2.Value = 1 Then
                    OrdBy = OrdBy & " DESC"
                End If
            End If
            sBuscar = sBuscar & OrdBy
            'Sentencia de agrupamiento GROUP BY
            If List1 <> "<TODOS>" Or List1.ListCount > 0 Then
                If Check1.Value = 1 Then
                    If List1.ListCount >= 1 Then
                        CampAgrup = " GROUP BY "
                        For Con = 0 To List1.ListCount - 1
                            CampAgrup = CampAgrup & List1.List(Con)
                            If Con < List1.ListCount - 1 Then
                                CampAgrup = CampAgrup & ", "
                            End If
                        Next
                    End If
                End If
            Else
                MsgBox "EL AGRUPAMIENTO NO SE REALIZARA YA QUE AGREGO LA OPCION DE TODOS LOS CAMPOS, PARA REALIZAR UNA SUMATORIA DEBE SELECCIONAR CAMPO POR CAMPO", vbExclamation, "SACC"
            End If
            sBuscar = sBuscar & CampAgrup
            Text1.Text = sBuscar
        End If
    End If
    If sBuscar <> "" Then
        ListView1.ListItems.Clear
        ListView1.ColumnHeaders.Clear
        If UCase(Mid(sBuscar, 1, 6)) = "UPDATE" Or UCase(Mid(sBuscar, 1, 6)) = "DELETE" Or UCase(Mid(sBuscar, 1, 6)) = "INSERT" Then
            Dim iAfectados As Long
            Set tRs = cnn.Execute(sBuscar, iAfectados, adCmdText)
            MsgBox "Se han afectado " & iAfectados & " registros en la ultima tarea", vbInformation, "SACC"
            Exit Sub
        Else
            Set tRs = cnn.Execute(sBuscar)
        End If
        With ListView1
            .View = lvwReport
            .Gridlines = True
            .LabelEdit = lvwManual
            .HideSelection = False
            .HotTracking = False
            .FullRowSelect = True
            .HoverSelection = False
            For Con = 0 To tRs.Fields.Count - 1
                .ColumnHeaders.Add , , tRs.Fields.Item(Con).Name & " ( Tamaño :" & tRs.Fields.Item(Con).DefinedSize & ")", 1500
            Next
        End With
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields(tRs.Fields.Item(0).Name))
            For Con = 1 To tRs.Fields.Count - 1
                If Not IsNull(tRs.Fields(tRs.Fields.Item(Con).Name)) Then tLi.SubItems(Con) = tRs.Fields(tRs.Fields.Item(Con).Name)
            Next
            tRs.MoveNext
            NoReg = NoReg + 1
        Loop
    End If
Exit Sub
ManejaError:
    If Err.Number = 91 Then
        MsgBox "NO SE HA ABIERTO UNA BASE DE DATOS", vbExclamation, "Hache's system"
    Else
        If Err.Number = -2147217900 Then
            MsgBox "LA SENTENCIA DE BUSQUEDA TIENE UN ERROR DE SINTAXIS", vbCritical, "Hache's system"
        Else
            If Err.Number = -2147467259 Then
                MsgBox "UNA O MAS EXPRECIONES DEL SUM NO SON TIPO DE DATOS NUMERICOS", vbCritical, "Hache's system"
            Else
                If Err.Number = -2147217887 Then
                    Err.Clear
                Else
                    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "Hache's system"
                End If
            End If
        End If
    End If
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Command8_Click()
    List4.RemoveItem (List4.ListCount - 1)
End Sub
Private Sub Command9_Click()
     If Combo4.Text <> "" And Combo5.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim sBuscar As String
        Dim CadAgre As String
        Dim Esta As Integer
        Dim Con As Integer
        sBuscar = Replace(Combo7.Text, "(BASE TABLE)", "")
        sBuscar = Replace(sBuscar, "(VIEW)", "")
        sBuscar = "SELECT " & Combo4.Text & " FROM " & sBuscar
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            If tRs.Fields(Combo4.Text).Type = adDate Or tRs.Fields(Combo4.Text).Type = adDBDate Then
                If Combo5.Text = "ENTRE FECHA Y FECHA" Then
                    CadAgre = "BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "'"
                Else
                    CadAgre = "'" & Combo6.Text & "'"
                End If
            Else
                If tRs.Fields(Combo4.Text).Type = adChar Or tRs.Fields(Combo4.Text).Type = adLongVarChar Or tRs.Fields(Combo4.Text).Type = adLongVarWChar Or tRs.Fields(Combo4.Text).Type = adVarChar Or tRs.Fields(Combo4.Text).Type = adVarWChar Or tRs.Fields(Combo4.Text).Type = adWChar Then
                    For Con = 0 To Combo6.ListCount
                        If Combo6.List(Con) = Combo6.Text Then
                            Esta = 1
                        End If
                    Next
                    If Esta = 1 Then
                        CadAgre = Combo6.Text
                    Else
                        If Combo5.Text = "QUE CONTENGA (PARTE DE CADENA)" Or Combo5.Text = "QUE NO CONTENGA (PARTE DE CADENA)" Then
                            CadAgre = "'%" & Combo6.Text & "%'"
                        Else
                            CadAgre = "'" & Combo6.Text & "'"
                        End If
                    End If
                Else
                    If tRs.Fields(Combo4.Text).Type = adBoolean Then
                        If Combo6.Text = "VERDADERO" Then
                            CadAgre = ".T."
                        Else
                            CadAgre = ".F."
                        End If
                    Else
                        CadAgre = Combo6.Text
                    End If
                End If
            End If
            If Combo5.Text = "QUE CONTENGA (PARTE DE CADENA)" Then
                CadAgre = "LIKE " & CadAgre
            End If
            If Combo5.Text = "QUE NO CONTENGA (PARTE DE CADENA)" Then
                CadAgre = "NOT LIKE " & CadAgre
            End If
            If Combo5.Text = "MAYOR QUE" Then
                CadAgre = "> " & CadAgre
            End If
            If Combo5.Text = "MENOR QUE" Then
                CadAgre = "< " & CadAgre
            End If
            If Combo5.Text = "IGUAL A (EXACTO)" Then
                CadAgre = "= " & CadAgre
            End If
            If Combo5.Text = "DIFERENTE A (EXACTO)" Then
                CadAgre = "<> " & CadAgre
            End If
            CadAgre = Combo4.Text & " " & CadAgre
            If List4.ListCount > 0 Then
                If Check3.Value = 1 Or Check4.Value = 1 Then
                    If Check3.Value = 1 Then
                        CadAgre = "AND " & CadAgre
                    End If
                    If Check4.Value = 1 Then
                        CadAgre = "OR " & CadAgre
                    End If
                     List4.AddItem CadAgre
                Else
                    MsgBox "DEBE SELECCIONAR UNA SENTENCIA Y U O PARA CONCATENAR CON LAS CONDICIONES ANTERIORES"
                End If
            Else
                 List4.AddItem CadAgre
                 Combo4.Text = ""
                 Combo5.Text = ""
                 Combo6.Text = ""
            End If
        End If
     End If
End Sub
Private Sub List1_Click()
    VarEl1 = List1.ListIndex
End Sub
Private Sub List1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub List3_Click()
    VarEl3 = List3.ListIndex
End Sub
Private Sub List3_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub List2_Click()
    VarEl2 = List2.ListIndex
End Sub
Private Sub List2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub SubMenRestablecer_Click()
    If MsgBox("ESTA OPCION ELIMINARA TODA LA INFORMACION SIN OPORTUNIDAD DE RECUPERARLA, ESTA SEGURO DE CONTINUAR?", vbYesNo, "SACC") = vbYes Then
        If InputBox("CONFIRME CONTRASEÑA", "SACC") = VarMen.Text1(4).Text Then
            Dim sBuscar As String
            Dim tRs As ADODB.Recordset
            Dim tRs1 As ADODB.Recordset
            sBuscar = "SELECT TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE From INFORMATION_SCHEMA.TABLES WHERE (TABLE_TYPE = 'BASE TABLE')"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not tRs.EOF
                    sBuscar = "DELETE FROM " & tRs.Fields("TABLE_NAME")
                    cnn.Execute (sBuscar)
                    sBuscar = "SELECT SCHEMA_NAME(OBJECTPROPERTY(object_id, 'SchemaId')) AS SchemaName, OBJECT_NAME(object_id) AS TableName, name AS ColumnName From sys.Columns Where (is_identity = 1) AND OBJECT_NAME(object_id) = '" & tRs.Fields("TABLE_NAME") & "' ORDER BY SchemaName, TableName, ColumnName"
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
                        sBuscar = "DBCC CHECKIDENT (" & tRs.Fields("TABLE_NAME") & ", RESEED,0)"
                        cnn.Execute (sBuscar)
                    End If
                    tRs.MoveNext
                Loop
                MsgBox "INFORMACION ELIMINADA", vbInformation, "SACC"
            Else
                MsgBox "LA BASE DE DATOS NO CUENTA CON TABLAS", vbInformation, "SACC"
            End If
        Else
            MsgBox "Conreaseña incorrecta ", vbExclamation, "SACC"
        End If
    End If
End Sub
