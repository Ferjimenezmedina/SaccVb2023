VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmOrdenRapida 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden de Compra Rapida"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   8520
      TabIndex        =   45
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   35
      Top             =   7680
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
         TabIndex        =   36
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmOrdenRapida.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmOrdenRapida.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   30
      Top             =   4080
      Width           =   975
      Begin VB.Image Image12 
         Height          =   780
         Left            =   120
         MouseIcon       =   "FrmOrdenRapida.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmOrdenRapida.frx":26F6
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Productos"
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
         TabIndex        =   31
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   28
      Top             =   5280
      Width           =   975
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
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
         TabIndex        =   29
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmOrdenRapida.frx":46E8
         MousePointer    =   99  'Custom
         Picture         =   "FrmOrdenRapida.frx":49F2
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   24
      Top             =   6480
      Width           =   975
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
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
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FrmOrdenRapida.frx":64A4
         MousePointer    =   99  'Custom
         Picture         =   "FrmOrdenRapida.frx":67AE
         Top             =   240
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   15478
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Orden de Compra"
      TabPicture(0)   =   "FrmOrdenRapida.frx":8380
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Option1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Option2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "FrmOrdenRapida.frx":839C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text16"
      Tab(1).Control(1)=   "Combo2"
      Tab(1).Control(2)=   "Text9"
      Tab(1).Control(3)=   "Command2"
      Tab(1).Control(4)=   "ListView3"
      Tab(1).Control(5)=   "Label25"
      Tab(1).Control(6)=   "Label12"
      Tab(1).Control(7)=   "Label10"
      Tab(1).ControlCount=   8
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   -68520
         TabIndex        =   57
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmOrdenRapida.frx":83B8
         Left            =   -72840
         List            =   "FrmOrdenRapida.frx":83BA
         TabIndex        =   54
         Top             =   7680
         Width           =   2775
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   -73560
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   8040
         Width           =   5295
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   6840
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Información para Agregar"
         Height          =   3495
         Left            =   120
         TabIndex        =   16
         Top             =   5160
         Width           =   8055
         Begin VB.CheckBox Check2 
            Caption         =   "$ Descuento"
            Height          =   195
            Left            =   240
            TabIndex        =   53
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   52
            Text            =   "0"
            Top             =   3120
            Width           =   975
         End
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   6960
            TabIndex        =   51
            Text            =   "0"
            Top             =   2640
            Width           =   975
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Otros"
            Height          =   195
            Left            =   5640
            TabIndex        =   50
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   4560
            TabIndex        =   49
            Text            =   "0"
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox Text13 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   4560
            TabIndex        =   48
            Text            =   "0"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox Text12 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   6960
            TabIndex        =   47
            Text            =   "0"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   46
            Text            =   "0"
            Top             =   2400
            Width           =   975
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Sumar  IVA 11%"
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   2640
            Width           =   1575
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Retencion ISR"
            Height          =   195
            Left            =   2880
            TabIndex        =   43
            Top             =   2640
            Width           =   1455
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Retencion IVA 10%"
            Height          =   195
            Left            =   2880
            TabIndex        =   42
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Retenciòn 4 %"
            Height          =   195
            Left            =   5640
            TabIndex        =   41
            Top             =   2160
            Width           =   1455
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Pagar Flete"
            Height          =   195
            Left            =   4440
            TabIndex        =   38
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Pagar Renta"
            Height          =   195
            Left            =   4440
            TabIndex        =   37
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "FrmOrdenRapida.frx":83BC
            Left            =   6480
            List            =   "FrmOrdenRapida.frx":83C6
            TabIndex        =   34
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   3120
            TabIndex        =   6
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
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
            Height          =   375
            Left            =   6840
            Picture         =   "FrmOrdenRapida.frx":83DA
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   3000
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Sumar IVA 16%"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   2160
            Width           =   1575
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   720
            Width           =   5055
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label8 
            Caption         =   "Proveedor :"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Cantidad :"
            Height          =   255
            Left            =   2280
            TabIndex        =   23
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Precio :"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción :"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Clave :"
            Height          =   255
            Left            =   4800
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   5160
         TabIndex        =   9
         Top             =   3240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   5160
         TabIndex        =   8
         Top             =   3000
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
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
         Left            =   7080
         Picture         =   "FrmOrdenRapida.frx":ADAC
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3000
         Width           =   1095
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
         Left            =   4800
         Picture         =   "FrmOrdenRapida.frx":D77E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   885
         Width           =   1050
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   3120
         Width           =   3975
      End
      Begin VB.CommandButton Command2 
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
         Left            =   -68160
         Picture         =   "FrmOrdenRapida.frx":10150
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7680
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   12
         Top             =   960
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   11668
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   1575
         Left            =   120
         TabIndex        =   4
         Top             =   3480
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2778
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
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   960
         Width           =   3615
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2778
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
      Begin VB.Label Label25 
         Caption         =   "Total : "
         Height          =   255
         Left            =   -69120
         TabIndex        =   56
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Departamento Solicitante"
         Height          =   255
         Left            =   -74760
         TabIndex        =   55
         Top             =   7680
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Comentarios :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   39
         Top             =   8040
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "No Orden :"
         Height          =   255
         Left            =   6000
         TabIndex        =   33
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Productos :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   8520
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FrmOrdenRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim IdProv As String
Dim ind As Integer
Dim ret1 As Double
Dim ret2 As Double
Dim retiva As Double
Dim isr As Double
Dim retisr As Double
Dim tipoimp As String 'tipo  10%
Dim tip As String 'tipo  de  4 %
Dim tipiva As String 'tipo  de  4 %
Dim retdiez   As Double
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Check8.Value = 0
        If Not IsNull(Text6.Text) And Not IsNull(Text5.Text) Then Text10.Text = Format((Val(Text6.Text) * Val(Text5.Text)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        Text10.Enabled = True
    Else
        Text10.Text = "0"
        Text10.Enabled = False
    End If
End Sub

Private Sub Check2_Click()
    If Me.Check2.Value = 1 Then
        Text15.Text = "0"
        Text15.Enabled = True
    Else
        Text15.Text = "0"
        Text15.Enabled = False
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Check4.Value = 0
        Check1.Value = 1
        Check6.Value = 1
        Check7.Value = 1
    Else
        Check1.Value = 0
        Check6.Value = 0
        Check7.Value = 0
    End If
End Sub
Private Sub Check4_Click()
    If Check4.Value = 1 Then
        Check3.Value = 0
        Check1.Value = 1
        Check5.Value = 1
    Else
        Check1.Value = 0
        Check5.Value = 0
    End If
End Sub
Private Sub Check5_Click()
    If Check5.Value = 1 Then
        If Text6.Text <> "" And Text5.Text <> "" Then
            Text12.Text = Format(CDbl(Text6.Text * Text5.Text) * 0.04, "###,###,##0.00")
        End If
        Text12.Enabled = True
    Else
        Text12.Enabled = False
        Text12.Text = "0"
    End If
End Sub
Private Sub Check6_Click()
    If Me.Check6.Value = 1 Then
        If Not IsNull(Text6.Text) And Not IsNull(Text5.Text) Then
            Text13.Text = Format((Val(Text6.Text) * Val(Text5.Text)) * CDbl(((CDbl(VarMen.Text4(7).Text) / 3) * 2) / 100), "###,###,##0.00")
            Text13.Enabled = True
        Else
            Text13.Text = "0"
            Text13.Enabled = False
        End If
    Else
        Text13.Text = "0"
        Text13.Enabled = False
    End If
End Sub
Private Sub Check7_Click()
    If Me.Check7.Value = 1 Then
        If Text5.Text <> "" Or Text6.Text <> "" Then
            Text14.Text = CDbl(Text5.Text * Text6.Text) * 0.1
        End If
        Text14.Enabled = True
    Else
        Text14.Text = "0"
        Text14.Enabled = False
    End If
End Sub
Private Sub Check8_Click()
    If Check8.Value = 1 Then
        Check1.Value = 0
        If Not IsNull(Text6.Text) And Not IsNull(Text5.Text) Then Text10.Text = Format((Val(Text6.Text) * Val(Text5.Text)) * 0.11, "###,###,##0.00")
        'If Not IsNull(Text6.Text) And Not IsNull(Text5.Text) Then Text10.Text = Format((Val(Text6.Text) * Val(Text5.Text)) * 0.11, "###,###,##0.00")
        Text10.Enabled = True
    Else
        Text10.Text = "0"
        Text10.Enabled = False
    End If
End Sub
Private Sub Check9_Click()
    If Check9.Value = 1 Then
        Text11.Text = "0"
        Text11.Enabled = True
    Else
        Text11.Enabled = False
        Text11.Text = "0"
    End If
End Sub
Private Sub Combo2_DropDown()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Combo2.Clear
    sBuscar = "SELECT DEPARTAMENTO FROM DEPARTAMENTOS WHERE (ESTATUS = 'A') AND PRESUPUESTO_MENSUAL > 0 ORDER BY DEPARTAMENTO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("DEPARTAMENTO")) Then
                Combo2.AddItem tRs.Fields("DEPARTAMENTO")
            End If
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command1_Click()
On Error GoTo ManejaError
    ret1 = 0
    ret2 = 0
    If IdProv <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" And Text6.Text <> "" And Combo1.Text <> "" Then
        Dim ItMx As ListItem
        Set ItMx = Me.ListView3.ListItems.Add(, , IdProv)
        If Not IsNull(Text3.Text) Then ItMx.SubItems(1) = Text3.Text
        If Not IsNull(Text4.Text) Then ItMx.SubItems(2) = Text4.Text
        If Not IsNull(Text5.Text) Then ItMx.SubItems(3) = Text5.Text
        If Not IsNull(Text6.Text) Then ItMx.SubItems(4) = Text6.Text
        If Not IsNull(Text6.Text) And Not IsNull(Text5.Text) Then ItMx.SubItems(5) = Format((CDbl(Text6.Text) * CDbl(Text5.Text)) - CDbl(Text15.Text), "0.00")
        If Not IsNull(Text10.Text) Then ItMx.SubItems(6) = Text10.Text
        ItMx.SubItems(7) = ((CDbl(Text5.Text) * CDbl(Text6.Text)) - CDbl(Text15.Text)) + CDbl(Text10.Text) - CDbl(Text13.Text) - CDbl(Text14.Text) - CDbl(Text12.Text) + CDbl(Text11.Text)
        ItMx.SubItems(8) = "S"
        If Not IsNull(Combo1.Text) Then ItMx.SubItems(9) = Combo1.Text
        If Not IsNull(Text13.Text) Then ItMx.SubItems(10) = Text13.Text
        ItMx.SubItems(11) = CDbl(Text12.Text) + CDbl(Text13.Text) + CDbl(Text14.Text)
        ItMx.SubItems(13) = Text15.Text
        If Not IsNull(Text14.Text) Then ItMx.SubItems(15) = Text14.Text
        If Not IsNull(Text11.Text) Then ItMx.SubItems(16) = Text11.Text
        'If Me.Check1.Value = 1 Then
        '    'IVA del 16%
        '    If Not IsNull(Text6.Text) And Not IsNull(Text5.Text) Then ItMx.SubItems(6) = Format((Val(Text6.Text) * Val(Text5.Text)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "0.00")
        '    If Not IsNull(Text6.Text) And Not IsNull(Text5.Text) Then ItMx.SubItems(7) = Format((Val(Text6.Text) * Val(Text5.Text)) * (CDbl(CDbl(VarMen.Text4(7).Text) / 100) + 1), "0.00")
        'Else
        '    ' IVA del 11%
        '    If Me.Check8.Value = 1 Then
        '        If Not IsNull(Text6.Text) And Not IsNull(Text5.Text) Then ItMx.SubItems(6) = Format((Val(Text6.Text) * Val(Text5.Text)) * 0.11, "0.00")
        '        If Not IsNull(Text6.Text) And Not IsNull(Text5.Text) Then ItMx.SubItems(7) = Format((Val(Text6.Text) * Val(Text5.Text)) * 1.11, "0.00")
        '    Else
        '        ItMx.SubItems(6)= "0.00"
        '        If Not IsNull(Text6.Text) And Not IsNull(Text5.Text) Then ItMx.SubItems(7) = Format(Val(Text6.Text) * Val(Text5.Text), "0.00")
        '    End If
        'End If
        ''Retencion de ISR
        'If Me.Check7.Value = 1 Then
        '    ItMx.SubItems(15) = Format((ItMx.SubItems(5)) * 0.1, "0.00")
        '    ItMx.SubItems(7) = CDbl(ItMx.SubItems(7)) - CDbl(ItMx.SubItems(15))
        '    retisr = ItMx.SubItems(15)
        'Else
        '    ItMx.SubItems(15)= "0.00"
        'End If
        ''Retencion del 4%
        'If Check5.Value = 1 Then
        '    ItMx.SubItems(11) = Format((ItMx.SubItems(5)) * 0.04, "0.00")
        '    ret1 = ItMx.SubItems(11)
        '    ItMx.SubItems(10) = CDbl(ItMx.SubItems(7)) - CDbl(Format(ItMx.SubItems(5)) * 0.04)
        '    ret2 = ItMx.SubItems(10)
        'End If
        ''Retencion del 10%
        'If Check6.Value = 1 Then
        '    ItMx.SubItems(13) = Format((ItMx.SubItems(5)) * 0.1066, "0.00")
        '    ItMx.SubItems(7) = CDbl(ItMx.SubItems(7)) - CDbl(ItMx.SubItems(13))
        'Else
        '    ItMx.SubItems(13)= "0.00"
        'End If
        Combo1.Enabled = False
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
    Else
        MsgBox "FALTA INFORMACIÓN NECESARIA PARA EL REGISTRO", vbInformation, "SACC"
    End If
    SumaOrden
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text1.SetFocus
    End If
End Sub
Private Sub Command2_Click()
    If ind <> 0 Then
        ListView3.ListItems.Remove (ind)
    End If
    SumaOrden
End Sub
Private Sub Command3_Click()
    If Text1.Text <> "" Then
        BUSCAPROVEEDOR
    End If
End Sub
Private Sub Command4_Click()
    If Text2.Text <> "" Then
        BUSCAPRODUCTO
    End If
End Sub
Private Sub Command5_Click()
    FunImprimeORCopia
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    IdProv = ""
    Check1.Caption = "Sumar IVA " & VarMen.Text4(7).Text & "%"
    Image3.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If NvoMen.Text1(22).Text = "N" Then
        Frame13.Visible = False
    End If
    If NvoMen.Text1(32).Text = "N" Then
        Frame3.Visible = False
    End If
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "# DEL PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE", 6100
        .ColumnHeaders.Add , , "DIRECCION", 1200
        .ColumnHeaders.Add , , "RFC", 1200
        .ColumnHeaders.Add , , "TELEFONO 1", 1200
        .ColumnHeaders.Add , , "TELEFONO 2", 1200
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE DEL PRODUCTO", 1600
        .ColumnHeaders.Add , , "DESCRIPCION", 6800
        .ColumnHeaders.Add , , "PRECIO", 1500
        .ColumnHeaders.Add , , "NOTAS", 1500
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE DEL PROVEEDOR", 0
        .ColumnHeaders.Add , , "CLAVE DEL PRODUCTO", 2000
        .ColumnHeaders.Add , , "DESCRIPCION", 6800
        .ColumnHeaders.Add , , "PRECIO COMPRA", 1500
        .ColumnHeaders.Add , , "CANTIDAD", 1500
        .ColumnHeaders.Add , , "SUBTOTAL", 1500
        .ColumnHeaders.Add , , "IVA", 1500
        .ColumnHeaders.Add , , "TOTAL", 1500
        .ColumnHeaders.Add , , "C.E.", 100
        .ColumnHeaders.Add , , "MONEDA", 100
        .ColumnHeaders.Add , , "RETENCION", 500
        .ColumnHeaders.Add , , "TOT-RETENCION", 1000
        .ColumnHeaders.Add , , "TIPO", 500
        .ColumnHeaders.Add , , "OTROS DESCUENTOS", 1500
        .ColumnHeaders.Add , , "TIPO DE  IVA", 1000
        .ColumnHeaders.Add , , "ISR", 1000
        .ColumnHeaders.Add , , "OTROS IMPUESTOS", 1000
    End With
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    Busca
End Sub
Private Sub SumaOrden()
    Dim Cont As Integer
    Dim SUMA As Double
    For Cont = 1 To ListView3.ListItems.Count
        SUMA = SUMA + CDbl(ListView3.ListItems(Cont).SubItems(7))
    Next Cont
    Text16.Text = Format(SUMA, "###,###,##0.00")
End Sub
Private Sub BUSCAPROVEEDOR()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim ItMx As ListItem
    sBuscar = "SELECT ID_PROVEEDOR, NOMBRE, DIRECCION, RFC, TELEFONO1, TELEFONO2 FROM PROVEEDOR_CONSUMO WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND (EMAIL IS NOT NULL) AND (EMAIL <> '')"
    Set tRs = cnn.Execute(sBuscar)
    Me.ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set ItMx = Me.ListView1.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then ItMx.SubItems(1) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("DIRECCION")) Then ItMx.SubItems(2) = tRs.Fields("DIRECCION")
                If Not IsNull(tRs.Fields("RFC")) Then ItMx.SubItems(3) = tRs.Fields("RFC")
                If Not IsNull(tRs.Fields("TELEFONO1")) Then ItMx.SubItems(4) = tRs.Fields("TELEFONO1")
                If Not IsNull(tRs.Fields("TELEFONO2")) Then ItMx.SubItems(5) = tRs.Fields("TELEFONO2")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Busca()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT ID_ORDEN_RAPIDA FROM ORDEN_RAPIDA ORDER BY ID_ORDEN_RAPIDA DESC"
    Set tRs = cnn.Execute(sBuscar)
    If tRs.EOF And tRs.BOF Then
        Text8.Text = 1
    Else
        Text8.Text = CDbl(tRs.Fields("ID_ORDEN_RAPIDA")) + 1
    End If
End Sub
Private Sub BUSCAPRODUCTO()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim ItMx As ListItem
    If Option1.Value = True Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO, NOTAS FROM PRODUCTOS_CONSUMIBLES WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%'"
    Else
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO, NOTAS FROM PRODUCTOS_CONSUMIBLES WHERE Descripcion LIKE '%" & Text2.Text & "%'"
    End If
    Set tRs = cnn.Execute(sBuscar)
    Me.ListView2.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set ItMx = Me.ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                If Not IsNull(tRs.Fields("Descripcion")) Then ItMx.SubItems(1) = tRs.Fields("Descripcion")
                If Not IsNull(tRs.Fields("PRECIO")) Then ItMx.SubItems(2) = tRs.Fields("PRECIO")
                If Not IsNull(tRs.Fields("NOTAS")) Then ItMx.SubItems(3) = tRs.Fields("NOTAS")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image1_Click()
On Error GoTo ManejaError
    If ListView3.ListItems.Count <> 0 And IdProv <> 0 And Combo2.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim Conta As Integer
        Dim Folio As String
        Dim Path As String
        Dim VarRent As String
        Dim TotCompras As String
        Dim TotPresupuesto As String
        Dim NumeroRegistros As Integer
        If Text8.Text = "" Then
            Text8.Text = 0
        End If
        If Me.Check3.Value = 0 Then
            If Me.Check4.Value = 0 Then
                VarRent = "N"
            Else
                VarRent = "F"
            End If
        Else
            VarRent = "S"
        End If
        'VERIFICA PRESUPUESTO
        sBuscar = "SELECT SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS TOT FROM ORDEN_RAPIDA INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA WHERE (ORDEN_RAPIDA.FECHA BETWEEN '01/" & Format(Date, "mm/yyyy") & "' AND '" & Format(Date, "dd/mm/yyyy") & "') AND ORDEN_RAPIDA.DEPARTAMENTO = '" & Combo2.Text & "' GROUP BY ORDEN_RAPIDA.DEPARTAMENTO"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            TotCompras = tRs.Fields("TOT")
        Else
            TotCompras = "0"
        End If
        sBuscar = "SELECT PRESUPUESTO_MENSUAL FROM DEPARTAMENTOS WHERE (ESTATUS = 'A') AND (DEPARTAMENTO = '" & Combo2.Text & "')"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            TotPresupuesto = tRs.Fields("PRESUPUESTO_MENSUAL")
        Else
            TotPresupuesto = "0"
        End If
        
        sBuscar = "SELECT ID_ORDEN_RAPIDA FROM ORDEN_RAPIDA WHERE ID_ORDEN_RAPIDA = " & Text8.Text
        Set tRs = cnn.Execute(sBuscar)
        If tRs.EOF And tRs.BOF Then
            If Text16.Text = "" Then
                Text16.Text = "0.00"
            End If
            If CDbl(TotCompras) + CDbl(Text16.Text) > CDbl(TotPresupuesto) Then
                MsgBox "EL DEPARTAMENTO YA SUPERÓ SU LÍMITE DE PRESUPUESTO", vbExclamation, "SACC"
                Exit Sub
            End If
            'COMPROBAR SI LA ORDEN ANTERIOR NO ES PARA EL MISMO PROVEEDOR
            sBuscar = "SELECT SUM(ORDEN_RAPIDA_DETALLE.TOTAL) AS Tot, ORDEN_RAPIDA.ID_ORDEN_RAPIDA FROM ORDEN_RAPIDA INNER JOIN ORDEN_RAPIDA_DETALLE ON ORDEN_RAPIDA.ID_ORDEN_RAPIDA = ORDEN_RAPIDA_DETALLE.ID_ORDEN_RAPIDA WHERE ORDEN_RAPIDA.ID_ORDEN_RAPIDA = " & CDbl(Text8.Text) - 1 & " AND ORDEN_RAPIDA.ID_PROVEEDOR = " & IdProv & " GROUP BY ORDEN_RAPIDA.ID_ORDEN_RAPIDA"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                If MsgBox("La Orden anterior es para el mismo proveedor por un total de $" & tRs.Fields("Tot") & ", desea continuar con la captura?", vbYesNo, "SACC") = vbNo Then
                    Exit Sub
                End If
            End If
        
            If Text9.Text <> "" Or Combo2.Text <> "" Then
                sBuscar = "INSERT INTO ORDEN_RAPIDA (ID_USUARIO, FECHA, ID_PROVEEDOR, MONEDA, ESTADO, RENTAS, COMENTARIO, RETENCION, DEPARTAMENTO) VALUES ('" & VarMen.Text1(0).Text & "', '" & Format(Date, "dd/mm/yyyy") & "', '" & IdProv & "', '" & ListView3.ListItems(1).SubItems(9) & "', 'A', '" & VarRent & "', '" & Text9.Text & "', '" & ret1 & "', '" & Combo2.Text & "');"
                cnn.Execute (sBuscar)
                sBuscar = "SELECT ID_ORDEN_RAPIDA FROM ORDEN_RAPIDA ORDER BY ID_ORDEN_RAPIDA DESC"
                Set tRs = cnn.Execute(sBuscar)
                Folio = tRs.Fields("ID_ORDEN_RAPIDA")
                If Not (tRs.EOF And tRs.BOF) Then
                    NumeroRegistros = ListView3.ListItems.Count
                    For Conta = 1 To NumeroRegistros
                        ListView3.ListItems(Conta).SubItems(7) = ListView3.ListItems(Conta).SubItems(7)
                        ListView3.ListItems(Conta).SubItems(6) = ListView3.ListItems(Conta).SubItems(6)
                        ListView3.ListItems(Conta).SubItems(5) = ListView3.ListItems(Conta).SubItems(5)
                        ListView3.ListItems(Conta).SubItems(4) = ListView3.ListItems(Conta).SubItems(4)
                        ListView3.ListItems(Conta).SubItems(3) = ListView3.ListItems(Conta).SubItems(3)
                        ListView3.ListItems(Conta).SubItems(2) = ListView3.ListItems(Conta).SubItems(2)
                        ListView3.ListItems(Conta).SubItems(1) = ListView3.ListItems(Conta).SubItems(1)
                        ListView3.ListItems(Conta).Text = ListView3.ListItems(Conta).Text
                        If Me.Check3.Value = 0 Then
                            If Me.Check4.Value = 0 Then
                                sBuscar = "INSERT INTO ORDEN_RAPIDA_DETALLE (ID_ORDEN_RAPIDA, ID_PRODUCTO, PRECIO, CANTIDAD, SUBTOTAL, IVA, TOTAL, ISR, IVARETENIDO, IVADIEZ, ISR2, OTROS_IMP, DESCUENTO) VALUES ('" & tRs.Fields("ID_ORDEN_RAPIDA") & "', '" & ListView3.ListItems(Conta).SubItems(1) & "', '" & Replace(ListView3.ListItems(Conta).SubItems(3), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(4), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(5), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(6), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(7), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(15), ",", "") & "','" & Replace(ListView3.ListItems(Conta).SubItems(10), ",", "") & "','" & Replace(ListView3.ListItems(Conta).SubItems(10), ",", "") & "','0','" & Replace(ListView3.ListItems(Conta).SubItems(16), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(13), ",", "") & "');"
                                cnn.Execute (sBuscar)
                            Else
                                sBuscar = "INSERT INTO ORDEN_RAPIDA_DETALLE (ID_ORDEN_RAPIDA, ID_PRODUCTO, PRECIO, CANTIDAD, SUBTOTAL, IVA, TOTAL, ISR, IVARETENIDO, IVADIEZ, TPIVA, ISR2, OTROS_IMP, DESCIENTO) VALUES ('" & tRs.Fields("ID_ORDEN_RAPIDA") & "', '" & ListView3.ListItems(Conta).SubItems(1) & "', '" & Replace(ListView3.ListItems(Conta).SubItems(3), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(4), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(5), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(6), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(7), ",", "") & "', 0,'" & Replace(ListView3.ListItems(Conta).SubItems(15), ",", "") & "','" & Replace(ListView3.ListItems(Conta).SubItems(10), ",", "") & "','" & Replace(ListView3.ListItems(Conta).SubItems(14), ",", "") & "', '0', '" & Replace(ListView3.ListItems(Conta).SubItems(16), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(13), ",", "") & "');"
                                cnn.Execute (sBuscar)
                            End If
                        Else
                            sBuscar = "INSERT INTO ORDEN_RAPIDA_DETALLE (ID_ORDEN_RAPIDA, ID_PRODUCTO, PRECIO, CANTIDAD, SUBTOTAL, IVA, TOTAL, ISR, IVARETENIDO, ISR2, OTROS_IMP, DESCUENTO) VALUES ('" & tRs.Fields("ID_ORDEN_RAPIDA") & "', '" & ListView3.ListItems(Conta).SubItems(1) & "', '" & Replace(ListView3.ListItems(Conta).SubItems(3), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(4), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(5), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(6), ",", "") & "', '" & (CDbl(ListView3.ListItems(Conta).SubItems(7)) - (CDbl(ListView3.ListItems(Conta).SubItems(5)) * 0.2)) & "', " & Format(CDbl(ListView3.ListItems(Conta).SubItems(5)) * 0.1, "0.00") & ", " & Format(CDbl(ListView3.ListItems(Conta).SubItems(5)) * 0.1, "0.00") & ",'" & retisr & "', '" & Replace(ListView3.ListItems(Conta).SubItems(16), ",", "") & "', '" & Replace(ListView3.ListItems(Conta).SubItems(13), ",", "") & "');"
                            cnn.Execute (sBuscar)
                        End If
                        If ListView3.ListItems(Conta).SubItems(8) = "S" Then
                            sBuscar = "INSERT INTO EXISTENCIA_FIJA (ID_PRODUCTO, CANTIDAD, ID_ORDEN_RAPIDA) VALUES ('" & ListView3.ListItems(Conta).SubItems(1) & "', " & ListView3.ListItems(Conta).SubItems(4) & ", " & tRs.Fields("ID_ORDEN_RAPIDA") & ");"
                            cnn.Execute (sBuscar)
                        End If
                    Next Conta
                    ListView3.ListItems.Clear
                    Text1.Enabled = True
                    ListView1.Enabled = True
                End If
                'PARCHE PARA ARREGLAR EL DESCUENTO DE LA RETENCION DEL IVA
                'sBuscar = "UPDATE ORDEN_RAPIDA_DETALLE SET TOTAL = TOTAL - IVARETENIDO WHERE ID_ORDEN_RAPIDA = " & Folio
                'cnn.Execute (sBuscar)
                ListView3.ListItems.Clear
                sBuscar = "SELECT ID_ORDEN_RAPIDA FROM ORDEN_RAPIDA ORDER BY ID_ORDEN_RAPIDA DESC"
                Set tRs = cnn.Execute(sBuscar)
                Text1.Enabled = True
                ListView1.Enabled = True
                Text7.Text = ""
                Combo1.Enabled = True
            Else
                MsgBox "Falta comentario y/o departamento para generar la órden", vbInformation, "SACC"
            End If
        Else
            Folio = tRs.Fields("ID_ORDEN_RAPIDA")
        End If
        Path = App.Path
        FunImprimeORCopia
        FunImprimeOR
        Text8.Text = CDbl(tRs.Fields("ID_ORDEN_RAPIDA")) + 1
        If Check1.Value = 1 And Check5.Value = 1 Then
            sBuscar = "UPDATE ORDEN_RAPIDA SET RETENCION =' " & ret1 & "' WHERE ID_ORDEN_RAPIDA = '" & Folio & "' "
            Set tRs = cnn.Execute(sBuscar)
        End If
    Else
        MsgBox "NO EXISTE UN LISTADO PARA IMPRIMIR!", vbInformation, "SACC"
    End If
    Busca
    SSTab1.Tab = 0
    CommonDialog1.Copies = 1
    IdProv = ""
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image12_Click()
    FrmProdConsumibles.Show vbModal
End Sub
Private Sub Image2_Click()
    FrmProvConsumibles.Show vbModal
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdProv = Item
    Text7.Text = Item.SubItems(1)
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
    End If
End Sub
Private Sub ListView1_LostFocus()
    If IdProv <> "" Then
        Text1.Enabled = False
        ListView1.Enabled = False
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text3.Text = Item
    Text4.Text = Item.SubItems(1)
    Text5.Text = Item.SubItems(2)
    Text6.Text = "1"
End Sub
Private Sub ListView2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text5.SetFocus
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ind = Item.Index
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text1.Text <> "" Then
        BUSCAPROVEEDOR
        ListView1.SetFocus
    End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text16_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Text2.Text <> "" Then
        BUSCAPRODUCTO
        ListView2.SetFocus
    End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text6.SetFocus
    End If
    Dim Valido As String
    Valido = "-1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command1.SetFocus
    End If
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    If NvoMen.Text1(11).Text = "N" Then
        MsgBox "No puede reimprimir por falta de permisos", vbExclamation, "SACC"
    Else
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        'Dim IdProv As String
        If Text8.Text <> "" And KeyAscii = 13 Then
            ListView3.ListItems.Clear
            sBuscar = "SELECT O.ID_ORDEN_RAPIDA, P.NOMBRE, P.ID_PROVEEDOR, O.RENTAS, O.COMENTARIO, O.RETENCION, O.DEPARTAMENTO FROM ORDEN_RAPIDA AS O LEFT JOIN PROVEEDOR_CONSUMO AS P ON O.ID_PROVEEDOR = P.ID_PROVEEDOR WHERE ID_ORDEN_RAPIDA = " & Text8.Text
            Set tRs = cnn.Execute(sBuscar)
            If tRs.EOF And tRs.BOF Then
                MsgBox "NO EXISTE LA ORDEN PEDIDA"
            Else
                Text9.Text = tRs.Fields("COMENTARIO")
                Check1.Value = 1
                Check5.Value = 1
                Me.Check4.Enabled = False
                Text7.Text = tRs.Fields("NOMBRE")
                IdProv = tRs.Fields("ID_PROVEEDOR")
                Combo2.Text = tRs.Fields("DEPARTAMENTO")
                sBuscar = "SELECT O.ID_PRODUCTO, P.Descripcion, O.PRECIO, O.CANTIDAD, O.SUBTOTAL, O.IVA, O.TOTAL, O.OTROS_IMP FROM ORDEN_RAPIDA_DETALLE AS O LEFT JOIN PRODUCTOS_CONSUMIBLES AS P ON O.ID_PRODUCTO = P.ID_PRODUCTO WHERE ID_ORDEN_RAPIDA = " & tRs.Fields("ID_ORDEN_RAPIDA")
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    Do While Not tRs.EOF
                        Dim ItMx As ListItem
                        Set ItMx = Me.ListView3.ListItems.Add(, , IdProv)
                        If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then ItMx.SubItems(1) = tRs.Fields("ID_PRODUCTO")
                        If Not IsNull(tRs.Fields("Descripcion")) Then ItMx.SubItems(2) = tRs.Fields("Descripcion")
                        If Not IsNull(tRs.Fields("PRECIO")) Then ItMx.SubItems(3) = tRs.Fields("PRECIO")
                        If Not IsNull(tRs.Fields("CANTIDAD")) Then ItMx.SubItems(4) = tRs.Fields("CANTIDAD")
                        If Not IsNull(tRs.Fields("SUBTOTAL")) Then ItMx.SubItems(5) = tRs.Fields("SUBTOTAL")
                        If Not IsNull(tRs.Fields("IVA")) Then ItMx.SubItems(6) = tRs.Fields("IVA")
                        If Not IsNull(tRs.Fields("TOTAL")) Then ItMx.SubItems(7) = tRs.Fields("TOTAL")
                        If Not IsNull(tRs.Fields("OTROS_IMP")) Then ItMx.SubItems(16) = tRs.Fields("OTROS_IMP")
                        tRs.MoveNext
                    Loop
                End If
                SSTab1.Tab = 1
            End If
        End If
        Dim Valido As String
        Valido = "1234567890"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub FunImprimeOR()
On Error GoTo ManejaError
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim COMENTARIO As String
    Dim Posi As Integer
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim sBuscar As String
    Dim Moneda As String
    Dim Estado As String
    sBuscar = "SELECT * FROM ORDEN_RAPIDA WHERE ID_ORDEN_RAPIDA = '" & CDbl(Text8.Text) & "'"
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        COMENTARIO = tRs1.Fields("comentario")
        Estado = tRs1.Fields("ESTADO")
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\OrdenCompraRapida.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        Image3.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image3, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 50, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA"
        Set tRs3 = cnn.Execute(sBuscar)
        oDoc.WTextBox 60, 10, 20, 280, tRs3.Fields("NOMBRE"), "F3", 8, hLeft
        oDoc.WTextBox 70, 10, 20, 280, tRs3.Fields("DIRECCION") & "  COl  " & tRs3.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 80, 10, 20, 280, tRs3.Fields("TELEFONO"), "F3", 8, hLeft
        oDoc.WTextBox 50, 340, 20, 280, "Fecha :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 280, "Orden de Compra Rapida : " & Text8.Text, "F3", 8, hCenter
        Moneda = tRs1.Fields("MONEDA")
        ' cuadros encabezado
        oDoc.WTextBox 100, 10, 100, 280, "PROVEEDOR: ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 300, 100, 280, "FACTURAR A :", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR_CONSUMO WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 115, 10, 100, 280, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 135, 10, 100, 280, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("COLONIA")) Then oDoc.WTextBox 145, 10, 100, 280, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CIUDAD")) Then oDoc.WTextBox 155, 10, 100, 280, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CP")) Then oDoc.WTextBox 165, 10, 100, 280, tRs2.Fields("CP"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO1")) Then oDoc.WTextBox 175, 10, 100, 280, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            'If Not IsNull(tRs2.Fields("TELEFONO2")) Then oDoc.WTextBox 175, 10, 100, 280, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("RFC")) Then oDoc.WTextBox 185, 10, 100, 280, tRs2.Fields("RFC"), "F3", 8, hCenter
        End If
        'CAJA2
        oDoc.WTextBox 115, 300, 100, 280, tRs3.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 125, 300, 100, 280, tRs3.Fields("DIRECCION"), "F3", 8, hCenter
        oDoc.WTextBox 135, 300, 100, 280, tRs3.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 145, 300, 100, 280, tRs3.Fields("CD") & ", " & tRs3.Fields("ESTADO"), "F3", 8, hCenter
        oDoc.WTextBox 155, 300, 100, 280, tRs3.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 165, 300, 100, 280, tRs3.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 175, 300, 100, 280, tRs3.Fields("RFC"), "F3", 8, hCenter
        Posi = 210
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 10, 40, 90, "ID PRODUCTO", "F2", 8, hCenter
        oDoc.WTextBox Posi, 112, 60, 80, "DESCRIPCION", "F2", 8, hCenter
        oDoc.WTextBox Posi, 305, 40, 260, "CANTIDAD", "F2", 8, hCenter
        oDoc.WTextBox Posi, 440, 80, 90, "P. UNITARIO ", "F2", 8, hCenter
        oDoc.WTextBox Posi, 490, 80, 90, "SUBTOTAL ", "F2", 8, hCenter
        Posi = Posi + 12
        ' Lineaf     despues d ekas cajas
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        sBuscar = "SELECT *  FROM VsOrdenCompraRapida WHERE ID_ORDEN_RAPIDA = '" & Text8.Text & "'"
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            Do While Not tRs2.EOF
                oDoc.WTextBox Posi, 10, 40, 110, tRs2.Fields("ID_PRODUCTO"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 112, 60, 290, Mid(tRs2.Fields("Descripcion"), 1, 55), "F3", 8, hLeft
                oDoc.WTextBox Posi, 395, 40, 50, Format(tRs2.Fields("CANTIDAD"), "###,###,##0.00"), "F3", 8, hRight
                oDoc.WTextBox Posi, 410, 80, 90, Format(tRs2.Fields("PRECIO"), "###,###,##0.00"), "F3", 8, hRight
                oDoc.WTextBox Posi, 460, 80, 90, Format(CDbl(tRs2.Fields("PRECIO")) * CDbl(tRs2.Fields("CANTIDAD")), "0.00"), "F3", 8, hRight
                Posi = Posi + 12
                tRs2.MoveNext
            Loop
        End If
        Posi = Posi + 6
        sBuscar = "SELECT SUM (SUBTOTAL) AS SUBTOTAL, SUM(IVARETENIDO) AS  IVARETENIDO, SUM(IVADIEZ) AS IVADIEZ, SUM(ISR2) AS ISR2 ,sum(iva) as iva,  SUM (ISR) AS  ISR, SUM (OTROS_IMP) AS OTROS, SUM (TOTAL) AS TOTAL FROM vsordenrapidadetalles WHERE ID_ORDEN_RAPIDA ='" & Text8.Text & "'"
        Set tRs1 = cnn.Execute(sBuscar)
        If Not (tRs1.EOF And tRs1.BOF) Then
            Posi = Posi + 10
            oDoc.WTextBox Posi, 20, 100, 275, "PRECIO EXPRESADO EN " & Moneda, "F3", 8, hLeft, , , 0, vbBlack
            Posi = Posi + 10
            oDoc.WTextBox Posi, 400, 20, 70, "SUBTOTAL:", "F2", 8, hRight
            oDoc.WTextBox Posi, 480, 20, 70, Format(tRs1.Fields("SUBTOTAL"), "###,###,##0.00"), "F2", 8, hRight
            Posi = Posi + 10
            If tRs1.Fields("IVA") > 0 Then
                oDoc.WTextBox Posi, 400, 20, 70, "I.V.A:", "F2", 8, hRight
                oDoc.WTextBox Posi, 480, 20, 70, Format(tRs1.Fields("IVA"), "###,###,##0.00"), "F2", 8, hRight
                Posi = Posi + 10
            Else
                If tRs1.Fields("IVADIEZ") > 0 Then
                    oDoc.WTextBox Posi, 400, 20, 70, "I.V.A:", "F2", 8, hRight
                    oDoc.WTextBox Posi, 480, 20, 70, Format(tRs1.Fields("IVA"), "###,###,##0.00"), "F2", 8, hRight
                    Posi = Posi + 10
                End If
            End If
            If tRs1.Fields("IVADIEZ") > 0 Then
                oDoc.WTextBox Posi, 400, 20, 70, "I.V.A RET 10%:", "F2", 8, hRight
                oDoc.WTextBox Posi, 480, 20, 70, Format((tRs1.Fields("IVADIEZ")), "###,###,##0.00"), "F2", 8, hRight
                Posi = Posi + 10
            End If
            If tRs1.Fields("IVARETENIDO") > 0 Then
                oDoc.WTextBox Posi, 400, 20, 70, "I.V.A RET:", "F2", 8, hRight
                oDoc.WTextBox Posi, 480, 20, 70, Format((tRs1.Fields("IVARETENIDO")), "###,###,##0.00"), "F2", 8, hRight
                Posi = Posi + 10
            End If
            If tRs1.Fields("ISR") > 0 Then
                oDoc.WTextBox Posi, 400, 20, 70, "ISR RET", "F2", 8, hRight
                oDoc.WTextBox Posi, 480, 20, 70, Format((tRs1.Fields("ISR")), "###,###,##0.00"), "F2", 8, hRight
                Posi = Posi + 10
            End If
            If tRs1.Fields("ISR2") > 0 Then
                oDoc.WTextBox Posi, 400, 20, 70, "RET ISR :", "F2", 8, hRight
                oDoc.WTextBox Posi, 480, 20, 70, Format(tRs1.Fields("ISR2"), "###,###,##0.00"), "F2", 8, hRight
                Posi = Posi + 10
            End If
            If tRs1.Fields("OTROS") > 0 Then
                oDoc.WTextBox Posi, 400, 20, 70, "IMPUESTOS :", "F2", 8, hRight
                oDoc.WTextBox Posi, 480, 20, 70, Format(tRs1.Fields("OTROS"), "###,###,##0.00"), "F2", 8, hRight
                Posi = Posi + 10
            End If
            oDoc.WTextBox Posi, 40, 20, 70, "COMENTARIO:", "F2", 8, hRight
            Posi = Posi + 10
            oDoc.WTextBox Posi, 40, 100, 350, COMENTARIO, "F2", 8, hRight
            Posi = Posi - 10
            oDoc.WTextBox Posi, 400, 20, 70, "TOTAL:", "F2", 8, hRight
            oDoc.WTextBox Posi, 480, 20, 70, Format(tRs1.Fields("TOTAL"), "###,###,##0.00"), "F2", 8, hRight
        End If
        Posi = Posi + 60
        If Estado = "C" Then
            oDoc.WTextBox Posi, 120, 100, 350, "CANCELADA", "F2", 36, hCenter, vTop, vbCyan
        End If
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 35, Posi
        oDoc.WLineTo 230, Posi
        oDoc.LineStroke
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 380, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 10
        oDoc.WTextBox Posi, 350, 20, 250, "AUTORIZADO POR(NOMBRE Y FIRMA)", "F3", 9, hCenter
        oDoc.WTextBox Posi, 5, 20, 250, "COMPRADOR(NOMBRE Y FIRMA)", "F3", 9, hCenter
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "El numero de orden no se ha capturado!", vbExclamation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox Err.Number & ": " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub FunImprimeORCopia()
On Error GoTo ManejaError
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim COMENTARIO As String
    Dim Posi As Integer
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim sBuscar As String
    Dim Moneda As String
    Dim Estado As String
    sBuscar = "SELECT * FROM ORDEN_RAPIDA WHERE ID_ORDEN_RAPIDA = '" & CDbl(Text8.Text) & "'"
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Estado = tRs1.Fields("ESTADO")
        COMENTARIO = tRs1.Fields("comentario")
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\OrdenCompraRapidaCopia.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        Image3.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image3, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 50, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA"
        Set tRs3 = cnn.Execute(sBuscar)
        oDoc.WTextBox 60, 10, 20, 280, tRs3.Fields("NOMBRE"), "F3", 8, hLeft
        oDoc.WTextBox 70, 10, 20, 280, tRs3.Fields("DIRECCION") & "  COl  " & tRs3.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 80, 10, 20, 280, tRs3.Fields("TELEFONO"), "F3", 8, hLeft
        oDoc.WTextBox 50, 340, 20, 280, "Fecha :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 280, "Orden de Compra Rapida : " & Text8.Text, "F3", 8, hCenter
        Moneda = tRs1.Fields("MONEDA")
        ' cuadros encabezado
        oDoc.WTextBox 100, 10, 100, 280, "PROVEEDOR: ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 300, 100, 280, "FACTURAR A :", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR_CONSUMO WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 115, 10, 100, 280, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 135, 10, 100, 280, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("COLONIA")) Then oDoc.WTextBox 145, 10, 100, 280, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CIUDAD")) Then oDoc.WTextBox 155, 10, 100, 280, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CP")) Then oDoc.WTextBox 165, 10, 100, 280, tRs2.Fields("CP"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO1")) Then oDoc.WTextBox 175, 10, 100, 280, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            'If Not IsNull(tRs2.Fields("TELEFONO2")) Then oDoc.WTextBox 175, 10, 100, 280, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("RFC")) Then oDoc.WTextBox 185, 10, 100, 280, tRs2.Fields("RFC"), "F3", 8, hCenter
        End If
        'CAJA2
        oDoc.WTextBox 115, 300, 100, 280, tRs3.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 125, 300, 100, 280, tRs3.Fields("DIRECCION"), "F3", 8, hCenter
        oDoc.WTextBox 135, 300, 100, 280, tRs3.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 145, 300, 100, 280, tRs3.Fields("CD") & ", " & tRs3.Fields("ESTADO"), "F3", 8, hCenter
        oDoc.WTextBox 155, 300, 100, 280, tRs3.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 165, 300, 100, 280, tRs3.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 175, 300, 100, 280, tRs3.Fields("RFC"), "F3", 8, hCenter
        Posi = 210
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 10, 40, 90, "ID PRODUCTO", "F2", 8, hCenter
        oDoc.WTextBox Posi, 112, 60, 80, "DESCRIPCION", "F2", 8, hCenter
        oDoc.WTextBox Posi, 305, 40, 260, "CANTIDAD", "F2", 8, hCenter
        oDoc.WTextBox Posi, 440, 80, 90, "P. UNITARIO ", "F2", 8, hCenter
        oDoc.WTextBox Posi, 490, 80, 90, "SUBTOTAL ", "F2", 8, hCenter
        Posi = Posi + 12
        ' Lineaf     despues d ekas cajas
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        sBuscar = "SELECT *  FROM VsOrdenCompraRapida WHERE ID_ORDEN_RAPIDA = '" & Text8.Text & "'"
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            Do While Not tRs2.EOF
                oDoc.WTextBox Posi, 10, 40, 110, tRs2.Fields("ID_PRODUCTO"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 112, 60, 290, Mid(tRs2.Fields("Descripcion"), 1, 55), "F3", 8, hLeft
                oDoc.WTextBox Posi, 395, 40, 50, Format(tRs2.Fields("CANTIDAD"), "###,###,##0.00"), "F3", 8, hRight
                oDoc.WTextBox Posi, 410, 80, 90, Format(tRs2.Fields("PRECIO"), "###,###,##0.00"), "F3", 8, hRight
                oDoc.WTextBox Posi, 460, 80, 90, Format(CDbl(tRs2.Fields("PRECIO")) * CDbl(tRs2.Fields("CANTIDAD")), "###,###,##0.00"), "F3", 8, hRight
                Posi = Posi + 12
                tRs2.MoveNext
            Loop
        End If
        Posi = Posi + 6
        sBuscar = "SELECT SUM (SUBTOTAL) AS SUBTOTAL, SUM(IVARETENIDO) AS  IVARETENIDO, SUM(IVADIEZ) AS IVADIEZ, SUM(ISR2) AS ISR2 ,sum(iva) as iva,  SUM (ISR) AS  ISR,  SUM (OTROS_IMP) AS  OTROS, SUM (TOTAL) AS TOTAL FROM vsordenrapidadetalles WHERE ID_ORDEN_RAPIDA ='" & Text8.Text & "'"
        Set tRs1 = cnn.Execute(sBuscar)
        If Not (tRs1.EOF And tRs1.BOF) Then
            Posi = Posi + 10
            oDoc.WTextBox Posi, 20, 100, 275, "PRECIO EXPRESADO EN " & Moneda, "F3", 8, hLeft, , , 0, vbBlack
            Posi = Posi + 10
            oDoc.WTextBox Posi, 400, 20, 70, "SUBTOTAL:", "F2", 8, hRight
            oDoc.WTextBox Posi, 480, 20, 70, Format(tRs1.Fields("SUBTOTAL"), "###,###,##0.00"), "F2", 8, hRight
            Posi = Posi + 10
            If tRs1.Fields("IVA") > 0 Then
                oDoc.WTextBox Posi, 400, 20, 70, "I.V.A:", "F2", 8, hRight
                oDoc.WTextBox Posi, 480, 20, 70, Format(tRs1.Fields("IVA"), "###,###,##0.00"), "F2", 8, hRight
                Posi = Posi + 10
            Else
                If tRs1.Fields("IVADIEZ") > 0 Then
                    oDoc.WTextBox Posi, 400, 20, 70, "I.V.A:", "F2", 8, hRight
                    oDoc.WTextBox Posi, 480, 20, 70, Format(tRs1.Fields("IVA"), "###,###,##0.00"), "F2", 8, hRight
                    Posi = Posi + 10
                End If
            End If
            If tRs1.Fields("IVADIEZ") > 0 Then
                oDoc.WTextBox Posi, 400, 20, 70, "I.V.A RET 10%:", "F2", 8, hRight
                oDoc.WTextBox Posi, 480, 20, 70, Format((tRs1.Fields("IVADIEZ")), "###,###,##0.00"), "F2", 8, hRight
                Posi = Posi + 10
            End If
            If tRs1.Fields("IVARETENIDO") > 0 Then
                oDoc.WTextBox Posi, 400, 20, 70, "I.V.A RET:", "F2", 8, hRight
                oDoc.WTextBox Posi, 480, 20, 70, Format((tRs1.Fields("IVARETENIDO")), "###,###,##0.00"), "F2", 8, hRight
                Posi = Posi + 10
            End If
            If tRs1.Fields("ISR") > 0 Then
                oDoc.WTextBox Posi, 400, 20, 70, "ISR RET", "F2", 8, hRight
                oDoc.WTextBox Posi, 480, 20, 70, Format((tRs1.Fields("ISR")), "###,###,##0.00"), "F2", 8, hRight
                Posi = Posi + 10
            End If
            If tRs1.Fields("ISR2") > 0 Then
                oDoc.WTextBox Posi, 400, 20, 70, "RET ISR :", "F2", 8, hRight
                oDoc.WTextBox Posi, 480, 20, 70, Format(tRs1.Fields("ISR2"), "###,###,##0.00"), "F2", 8, hRight
                Posi = Posi + 10
            End If
            If tRs1.Fields("OTROS") > 0 Then
                oDoc.WTextBox Posi, 400, 20, 70, "IMPUESTOS :", "F2", 8, hRight
                oDoc.WTextBox Posi, 480, 20, 70, Format(tRs1.Fields("OTROS"), "###,###,##0.00"), "F2", 8, hRight
                Posi = Posi + 10
            End If
            oDoc.WTextBox Posi, 40, 20, 70, "COMENTARIO:", "F2", 8, hRight
            Posi = Posi + 10
            oDoc.WTextBox Posi, 40, 100, 350, COMENTARIO, "F2", 8, hRight
            Posi = Posi - 10
            oDoc.WTextBox Posi, 400, 20, 70, "TOTAL:", "F2", 8, hRight
            oDoc.WTextBox Posi, 480, 20, 70, Format(tRs1.Fields("TOTAL"), "###,###,##0.00"), "F2", 8, hRight
        End If
        Posi = Posi + 60
        If Estado = "C" Then
            oDoc.WTextBox Posi, 120, 100, 350, "COPIA CANCELADA", "F2", 36, hCenter, vTop, vbCyan
        Else
            oDoc.WTextBox Posi, 120, 100, 350, "COPIA", "F2", 36, hCenter, vTop, vbCyan
        End If
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 35, Posi
        oDoc.WLineTo 230, Posi
        oDoc.LineStroke
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 380, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 10
        oDoc.WTextBox Posi, 350, 20, 250, "AUTORIZADO POR(NOMBRE Y FIRMA)", "F3", 9, hCenter
        oDoc.WTextBox Posi, 5, 20, 250, "COMPRADOR(NOMBRE Y FIRMA)", "F3", 9, hCenter
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "El numero de orden no se ha capturado!", vbExclamation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox Err.Number & ": " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
