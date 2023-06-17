VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCapCompLic 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capturar Licitación de Competidores"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame20 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8280
      TabIndex        =   27
      Top             =   2880
      Width           =   975
      Begin VB.Image Image19 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmCapCompLic.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCapCompLic.frx":030A
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Compe."
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
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8280
      TabIndex        =   1
      Top             =   4080
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCapCompLic.frx":1DBC
         MousePointer    =   99  'Custom
         Picture         =   "FrmCapCompLic.frx":20C6
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Cliente"
      TabPicture(0)   =   "FrmCapCompLic.frx":41A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Competidores"
      TabPicture(1)   =   "FrmCapCompLic.frx":41C4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command10"
      Tab(1).Control(1)=   "Text8"
      Tab(1).Control(2)=   "Command7"
      Tab(1).Control(3)=   "Command1"
      Tab(1).Control(4)=   "ListView1"
      Tab(1).Control(5)=   "Label9"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Articulos"
      TabPicture(2)   =   "FrmCapCompLic.frx":41E0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Text11"
      Tab(2).Control(2)=   "Command11"
      Tab(2).Control(3)=   "Frame3"
      Tab(2).Control(4)=   "Text9"
      Tab(2).Control(5)=   "Command8"
      Tab(2).Control(6)=   "Command2"
      Tab(2).Control(7)=   "ListView2"
      Tab(2).Control(8)=   "Label12"
      Tab(2).Control(9)=   "Label10"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Generales"
      TabPicture(3)   =   "FrmCapCompLic.frx":41FC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command4"
      Tab(3).Control(1)=   "Command3"
      Tab(3).Control(2)=   "Frame1"
      Tab(3).Control(3)=   "ListView3"
      Tab(3).Control(4)=   "ListView4"
      Tab(3).Control(5)=   "Label2"
      Tab(3).Control(6)=   "Label1"
      Tab(3).ControlCount=   7
      Begin VB.Frame Frame4 
         Caption         =   "Cantidad"
         Height          =   855
         Left            =   -71640
         TabIndex        =   46
         Top             =   4200
         Width           =   2175
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   240
            TabIndex        =   47
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74880
         TabIndex        =   45
         Top             =   4560
         Width           =   3015
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Agregar"
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
         Left            =   6840
         Picture         =   "FrmCapCompLic.frx":4218
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command11 
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
         Left            =   -68160
         Picture         =   "FrmCapCompLic.frx":6BEA
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
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
         Left            =   -68520
         Picture         =   "FrmCapCompLic.frx":95BC
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
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
         Left            =   6480
         Picture         =   "FrmCapCompLic.frx":BF8E
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   4560
         Width           =   4335
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -69840
         TabIndex        =   35
         Top             =   480
         Width           =   1575
         Begin VB.OptionButton Option4 
            Caption         =   "Por Descripción"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Por Clave"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   120
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -73920
         TabIndex        =   34
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -73920
         TabIndex        =   32
         Top             =   720
         Width           =   5175
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Siguiente"
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
         Left            =   -68160
         Picture         =   "FrmCapCompLic.frx":E960
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Siguiente"
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
         Left            =   -68160
         Picture         =   "FrmCapCompLic.frx":11332
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Numero de Licitación"
         Height          =   855
         Left            =   4560
         TabIndex        =   25
         Top             =   4200
         Width           =   2175
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   240
            MaxLength       =   50
            TabIndex        =   26
            Top             =   360
            Width           =   1695
         End
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   3015
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5318
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Top             =   720
         Width           =   5175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
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
         Picture         =   "FrmCapCompLic.frx":13D04
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Agregar"
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
         Picture         =   "FrmCapCompLic.frx":166D6
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Quitar"
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
         Left            =   -73560
         Picture         =   "FrmCapCompLic.frx":190A8
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Quitar"
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
         Left            =   -71040
         Picture         =   "FrmCapCompLic.frx":1BA7A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Captura"
         Height          =   4575
         Left            =   -69840
         TabIndex        =   7
         Top             =   480
         Width           =   2775
         Begin VB.CommandButton Command12 
            Caption         =   "Cerrar"
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
            Left            =   1560
            Picture         =   "FrmCapCompLic.frx":1E44C
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   3960
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   2400
            Width           =   2535
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   3000
            Width           =   2535
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   1800
            Width           =   2535
         End
         Begin VB.Label Label6 
            Caption         =   "Cliente :"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Producto :"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Competidor :"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Numero de Licitación :"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Width           =   1815
         End
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   3
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   7011
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
         Height          =   3975
         Left            =   -72360
         TabIndex        =   5
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   7011
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   18
         Top             =   1080
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5318
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   20
         Top             =   1080
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   6165
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label12 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   44
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Buscar :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   33
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Buscar :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   31
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Buscar :"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Competidores :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Articulos :"
         Height          =   255
         Left            =   -72240
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmCapCompLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim IdClien As String
Dim IdClienxx As String
Dim IdCompe As String
Dim NomCompe As String
Dim IdProdu As String
Dim Eli1 As Integer
Dim Eli2 As Integer
Private Sub Command1_Click()
    If IdCompe <> "" Then
        Dim tLi As ListItem
        Set tLi = ListView3.ListItems.Add(, , IdCompe)
        tLi.SubItems(1) = NomCompe
        IdCompe = ""
        NomCompe = ""
    Else
        MsgBox "NO SE HA SELECCIONADO UN REGISTRO!", vbInformation, "SACC"
    End If
End Sub
Private Sub Command12_Click()
    If ListView3.ListItems.Count <> 0 And ListView4.ListItems.Count <> 0 And Text3.Text <> "" And Text4.Text <> "" Then
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        ListView3.Enabled = False
        ListView4.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        Command12.Enabled = False
        Dim Cont1 As Integer
        Dim CONT2 As Integer
        Dim Acum As Integer
        Dim ACUM2 As Integer
        Dim ClvCompe As String
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Cont1 = ListView3.ListItems.Count
        CONT2 = ListView4.ListItems.Count
        For Acum = 1 To Cont1
            ClvCompe = ListView3.ListItems(Acum)
            Text1.Text = ListView3.ListItems(Acum).SubItems(1)
            For ACUM2 = 1 To CONT2
                Text2.Text = ListView4.ListItems(ACUM2)
                sBuscar = "INSERT INTO COMPARATIVO_LICITACIONES(ID_CLIENTE, ID_COMPETIDOR, NO_LICITACION, ID_PRODUCTO, CANTIDAD, PRECIO) VALUES(" & IdClien & "," & ClvCompe & ",'" & Text3.Text & "','" & Text2.Text & "'," & ListView4.ListItems(ACUM2).SubItems(1) & ", " & InputBox("PRECIO DADO POR : " & Text1.Text & " EN EL ARTICULO : " & Text2.Text, "SACC") & ")"
                cnn.Execute (sBuscar)
            Next ACUM2
        Next Acum
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
        ListView3.Enabled = True
        ListView4.Enabled = True
        Command3.Enabled = True
        Command4.Enabled = True
        Command12.Enabled = True
    Else
        MsgBox "FALTA INFORMACIÓN NECESARIA", vbInformation, "SACC"
    End If
End Sub
Private Sub Command2_Click()
    If IdProdu <> "" Then
        If Text12.Text <> "" Then
            Dim tLi As ListItem
            Set tLi = ListView4.ListItems.Add(, , IdProdu)
            tLi.SubItems(1) = Text12.Text
            IdProdu = ""
            Text12.Text = ""
            Text11.Text = ""
        Else
            MsgBox "DEBE DAR UNA CANTIDAD!", vbInformation, "SACC"
        End If
    Else
        MsgBox "DEBE SELECCIONAR UN ARTICULO!", vbInformation, "SACC"
    End If
End Sub
Private Sub Command3_Click()
    If Eli2 <> "" Then
        ListView4.ListItems.Remove (Eli2)
        Eli2 = 0
    Else
        MsgBox "ES NECESARIO SELECCIONAR UN REGISTRO", vbInformation, "SACC"
    End If
End Sub
Private Sub Command4_Click()
    If Eli1 <> "" Then
        ListView3.ListItems.Remove (Eli1)
        Eli1 = 0
    Else
        MsgBox "ES NECESARIO SELECCIONAR UN REGISTRO", vbInformation, "SACC"
    End If
End Sub
Private Sub Command6_Click()
    If Text10.Text <> "" And Text7.Text <> "" Then
        Text4.Text = Text10.Text
        Text3.Text = Text7.Text
        IdClien = IdClienxx
        SSTab1.Tab = 1
        Text4.Text = ""
        Text3.Text = ""
    Else
        MsgBox "FALTA INFORMACIÓN NECESARIA!", vbInformation, "SACC"
    End If
End Sub
Private Sub Command7_Click()
    SSTab1.Tab = 2
End Sub
Private Sub Command8_Click()
    SSTab1.Tab = 3
End Sub
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
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Competidor", 0
        .ColumnHeaders.Add , , "Nombre", 7050
        .ColumnHeaders.Add , , "Notas", 7050
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 6050
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Competidor", 0
        .ColumnHeaders.Add , , "Nombre", 3050
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Producto", 2320
        .ColumnHeaders.Add , , "Cantidad", 0
    End With
    With ListView5
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave del Cliente", 0
        .ColumnHeaders.Add , , "Nombre", 7050
        .ColumnHeaders.Add , , "RFC", 7050
    End With
End Sub
Private Sub Image19_Click()
    FrmCompetenciaLic.Show vbModal
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdCompe = Item
    NomCompe = Item.SubItems(1)
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdProdu = Item
    Text11.Text = Item
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Eli1 = Item.Index
End Sub
Private Sub ListView4_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Eli2 = Item.Index
End Sub
Private Sub ListView5_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text10.Text = Item.SubItems(1)
    IdClienxx = Item
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        sBuscar = "SELECT ID_CLIENTE, NOMBRE, RFC FROM CLIENTE WHERE NOMBRE LIKE '%" & Replace(Text6.Text, " ", "%") & "%' AND VALORACION = 'A'"
        Set tRs = cnn.Execute(sBuscar)
        ListView5.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView5.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("RFC")) Then tLi.SubItems(2) = tRs.Fields("RFC")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        sBuscar = "SELECT ID_COMPETIDOR, NOMBRE, NOTAS FROM COMPETIDOR_LICITACION WHERE NOMBRE LIKE '%" & Replace(Text8.Text, " ", "%") & "%'"
        Set tRs = cnn.Execute(sBuscar)
        ListView1.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_COMPETIDOR"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("NOTAS")) Then tLi.SubItems(2) = tRs.Fields("NOTAS")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        If Option3.Value = True Then
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Replace(Text9.Text, " ", "%") & "%'"
        Else
            sBuscar = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE Descripcion LIKE '%" & Replace(Text9.Text, " ", "%") & "%'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        ListView2.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
