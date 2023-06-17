VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCotizaRapida 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cotizacion Rapida"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   8880
      TabIndex        =   37
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   9480
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8880
      TabIndex        =   35
      Top             =   3840
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
         TabIndex        =   36
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmCotizaRapida.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCotizaRapida.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8880
      TabIndex        =   33
      Top             =   5040
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
         TabIndex        =   34
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Command2 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FrmCotizaRapida.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmCotizaRapida.frx":2156
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8895
      TabIndex        =   31
      Top             =   6240
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
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCotizaRapida.frx":3D28
         MousePointer    =   99  'Custom
         Picture         =   "FrmCotizaRapida.frx":4032
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmCotizaRapida.frx":6114
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CommonDialog1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Option3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Option4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Option5"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.OptionButton Option5 
         Caption         =   "Por codigo de barras"
         Height          =   195
         Left            =   5760
         TabIndex        =   27
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cantidad"
         Height          =   1455
         Left            =   7080
         TabIndex        =   24
         Top             =   3240
         Width           =   1455
         Begin VB.CommandButton Command4 
            Caption         =   "Agregar"
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
            Left            =   120
            Picture         =   "FrmCotizaRapida.frx":6130
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Eliminar"
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
         Left            =   7200
         Picture         =   "FrmCotizaRapida.frx":8B02
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Promociónes "
         Height          =   1215
         Left            =   5760
         TabIndex        =   20
         Top             =   120
         Width           =   2775
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Aplicar Promoción"
            Height          =   255
            Left            =   600
            TabIndex        =   21
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Articulos de la Venta "
         Height          =   2415
         Left            =   120
         TabIndex        =   12
         Top             =   4800
         Width           =   6855
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "0.00"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "0.00"
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   38
            Text            =   "0.00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "0.00"
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "0.00"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "0.00"
            Top             =   240
            Width           =   975
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   2055
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   3625
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
         Begin VB.Label Label9 
            Caption         =   "Ret."
            Height          =   255
            Left            =   5040
            TabIndex        =   43
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "Imp. 2"
            Height          =   255
            Left            =   5040
            TabIndex        =   41
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Imp. 1"
            Height          =   255
            Left            =   5040
            TabIndex        =   39
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Total"
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
            Left            =   5040
            TabIndex        =   19
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "IVA"
            Height          =   255
            Left            =   5040
            TabIndex        =   18
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Subtotal"
            Height          =   255
            Left            =   5040
            TabIndex        =   17
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   5760
         TabIndex        =   11
         Top             =   2400
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Por Descripcion"
         Height          =   195
         Left            =   5760
         TabIndex        =   10
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   2760
         Width           =   4695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cliente "
         Height          =   2535
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   5535
         Begin VB.OptionButton Option2 
            Caption         =   "Por Nombre"
            Height          =   195
            Left            =   4200
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Por Clave"
            Height          =   195
            Left            =   4200
            TabIndex        =   6
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   3975
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1695
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   2990
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
      Begin VB.CommandButton Command6 
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
         Left            =   7200
         Picture         =   "FrmCotizaRapida.frx":B4D4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Limpiar"
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
         Left            =   7200
         Picture         =   "FrmCotizaRapida.frx":DEA6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6840
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   615
         Left            =   5760
         MaxLength       =   130
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1680
         Width           =   2775
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5880
         Top             =   6120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         PrinterDefault  =   0   'False
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1575
         Left            =   120
         TabIndex        =   28
         Top             =   3120
         Width           =   6855
         _ExtentX        =   12091
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
      Begin VB.Label Label1 
         Caption         =   "Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Comentarios :"
         Height          =   255
         Left            =   5760
         TabIndex        =   29
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   9000
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "FrmCotizaRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim ClvProd As String
Dim DesProd As String
Dim PreProd As String
Dim CLVCLIEN As String
Dim NomClien As String
Dim DesClien As String
Dim DelInd As String
Dim DelDes As String
Dim DelCan As String
Dim DelPre As String
Dim BanCnn As Boolean
Dim IdCotizacion As String
Dim TipoDesc As String
Dim ClasProd As String
Dim IVA As String
Dim IMP1 As String
Dim IMP2 As String
Dim RET As String
Private Sub Check1_Click()
On Error GoTo ManejaError
    If Check1.Value = 1 Then
        Me.Combo1.Enabled = True
    Else
        Me.Combo1.Enabled = False
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Combo1_DropDown()
On Error GoTo ManejaError
    Me.Combo1.Clear
    Dim tRs As ADODB.Recordset
    Dim sBus As String
    sBus = "SELECT TIPO FROM PROMOCION ORDER BY TIPO"
    Set tRs = cnn.Execute(sBus)
    Combo1.AddItem "<NINGUNA>"
    Combo1.AddItem "LICITACIÓN"
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            If Not ISNULL(tRs.Fields("TIPO")) Then Combo1.AddItem tRs.Fields("TIPO")
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HFFE1E1
End Sub
Private Sub Combo1_LostFocus()
      Combo1.BackColor = &H80000005
End Sub
Private Sub Command1_Click()
    If ListView3.ListItems.Count = 0 Then
        Text1.Text = ""
        Text2.Text = ""
        Text8.Text = ""
        Text3.Text = "0.00"
        Text4.Text = "0.00"
        Text5.Text = "0.00"
        ListView1.ListItems.Clear
        ListView2.ListItems.Clear
        ListView3.ListItems.Clear
        Text1.Enabled = True
        Text1.SetFocus
    Else
        MsgBox "No debe tener ningun articulo de venta", , "SACC"
    End If
End Sub
Private Sub Command2_Click()
    Command6.Value = True
    'ImprimeCotiza
    FunImpr
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    If DelInd <> "" Then
        Dim tRs As ADODB.Recordset
        Dim sBuscar As String
        Text3.Text = Format(CDbl(Text3.Text) - CDbl(ListView3.SelectedItem.SubItems(3)), "###,###,##0.00")
        If CDbl(ListView3.SelectedItem.SubItems(5)) <> 0 Then Text4.Text = Format(CDbl(Text4.Text) - (CDbl(ListView3.SelectedItem.SubItems(3)) * CDbl(ListView3.SelectedItem.SubItems(5))), "###,###,##0.00")
        If CDbl(ListView3.SelectedItem.SubItems(6)) <> 0 Then Text9.Text = Format(CDbl(Text9.Text) - (CDbl(ListView3.SelectedItem.SubItems(3)) * CDbl(ListView3.SelectedItem.SubItems(6))), "###,###,##0.00")
        If CDbl(ListView3.SelectedItem.SubItems(7)) <> 0 Then Text10.Text = Format(CDbl(Text10.Text) - (CDbl(ListView3.SelectedItem.SubItems(3)) * CDbl(ListView3.SelectedItem.SubItems(7))), "###,###,##0.00")
        If CDbl(ListView3.SelectedItem.SubItems(8)) <> 0 Then Text11.Text = Format(CDbl(Text11.Text) - (CDbl(ListView3.SelectedItem.SubItems(3)) * CDbl(ListView3.SelectedItem.SubItems(8))), "###,###,##0.00")
        'If IMP1 <> "" Then Text9.Text = Format(CDbl(Text3.Text) * CDbl(IMP1), "###,###,##0.00")
        'If IMP2 <> "" Then Text10.Text = Format(CDbl(Text3.Text) * CDbl(IMP2), "###,###,##0.00")
        'If RET <> "" Then Text11.Text = Format(CDbl(Text3.Text) * CDbl(RET), "###,###,##0.00")
        Text5.Text = Format(CDbl(Text3.Text) + CDbl(Text4.Text) + CDbl(Text9.Text) + CDbl(Text10.Text) - CDbl(Text11.Text), "###,###,##0.00")
        
        'Text3.Text = Format(CDbl(Text3.Text) - CDbl(ListView3.SelectedItem.SubItems(3)), "###,###,##0.00")
        'Text4.Text = Format(CDbl(Text3.Text) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        'Text5.Text = Format(CDbl(Text4.Text) + CDbl(Text3.Text), "###,###,##0.00")
        DelInd = ""
        Command3.Enabled = False
        ListView3.ListItems.Remove ListView3.SelectedItem.Index
    End If
    If CLVCLIEN <> "" And ListView3.ListItems.Count > 0 Then
        Me.Command2.Enabled = True
        Command6.Enabled = True
    Else
        Me.Command2.Enabled = False
        Command6.Enabled = False
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command4_Click()
On Error GoTo ManejaError
    Agregar
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command6_Click()
    Dim tRs As ADODB.Recordset
    Dim sqlComanda As String
    Dim NRegistros  As Integer
    Dim Con As Integer
    Text5.Text = Replace(Text5.Text, ",", "")
    Text3.Text = Replace(Text3.Text, ",", "")
    sqlComanda = "INSERT INTO COTIZA_CLIEN (ID_CLIENTE, ID_AGENTE, FECHA, TOTAL, SUBTOTAL) VALUES (" & CLVCLIEN & ", " & VarMen.Text1(0).Text & ", '" & Format(Date, "dd/mm/yyyy") & "', " & Text5.Text & ", " & Text3.Text & ");"
    cnn.Execute (sqlComanda)
    sqlComanda = "SELECT ID_COTIZA_CLIEN FROM COTIZA_CLIEN WHERE FECHA = '" & Format(Date, "dd/mm/yyyy") & "' ORDER BY ID_COTIZA_CLIEN DESC"
    Set tRs = cnn.Execute(sqlComanda)
    NRegistros = ListView3.ListItems.Count
    IdCotizacion = tRs.Fields("ID_COTIZA_CLIEN")
    For Con = 1 To NRegistros
        ListView3.ListItems(Con).SubItems(2) = Replace(ListView3.ListItems(Con).SubItems(2), ",", "")
        ListView3.ListItems(Con).SubItems(3) = Replace(ListView3.ListItems(Con).SubItems(3), ",", "")
        sqlComanda = "INSERT INTO COTIZA_CLIEN_DETALLE (ID_PRODUCTO, CANTIDAD, PRECIO_VENTA, ID_COTIZA_CLIEN) VALUES ('" & ListView3.ListItems(Con).Text & "', " & ListView3.ListItems(Con).SubItems(2) & ", " & ListView3.ListItems(Con).SubItems(3) & ", " & tRs.Fields("ID_COTIZA_CLIEN") & ");"
        cnn.Execute (sqlComanda)
    Next Con
    'ImprimeCotiza
    FunImpr
    MsgBox "LA COTIZACION A SIDO GUARDADA CON EL NUMERO " & tRs.Fields("ID_COTIZA_CLIEN"), vbInformation, "SACC"
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
    Text1.Text = ""
    Text3.Text = "0.00"
    Text4.Text = "0.00"
    Text5.Text = "0.00"
    Text8.Text = ""
    Text9.Text = "0.00"
    Text10.Text = "0.00"
    Text11.Text = "0.00"
    Combo1.Text = ""
    Check1.Value = 0
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    DesClien = 0
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "# DEL CLIENTE", 0
        .ColumnHeaders.Add , , "NOMBRE", 6100
        .ColumnHeaders.Add , , "DESCUENTO", 1200
        .ColumnHeaders.Add , , "TIPO DE DESCUENTO", 1000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "PRODUCTO", 1500
        .ColumnHeaders.Add , , "DESCRIPCIÓN", 2500
        .ColumnHeaders.Add , , "PRECIO", 700
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "CLASIFICACIÓN", 1000
        .ColumnHeaders.Add , , "IVA", 1000
        .ColumnHeaders.Add , , "IMPUESTO 1", 1000
        .ColumnHeaders.Add , , "IMPUESTO 2", 1000
        .ColumnHeaders.Add , , "RETENCION", 1000
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "PRODUCTO", 2400
        .ColumnHeaders.Add , , "DESCRIPCIÓN", 6800
        .ColumnHeaders.Add , , "CANTIDAD", 2000
        .ColumnHeaders.Add , , "PRECIO", 2000
        .ColumnHeaders.Add , , "PRECIO UNITARIO", 1000
        .ColumnHeaders.Add , , "IVA", 1000
        .ColumnHeaders.Add , , "IMPUESTO 1", 1000
        .ColumnHeaders.Add , , "IMPUESTO 2", 1000
        .ColumnHeaders.Add , , "RETENCION", 1000
    End With
    CLVCLIEN = ""
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image10_Click()
    If ListView3.ListItems.Count > 0 Then
        Dim StrCopi As String
        Dim Con As Integer
        Dim Con2 As Integer
        Dim NumColum As Integer
        Dim Ruta As String
        'FileName
        Me.CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "Guardar como"
        CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
        Me.CommonDialog1.ShowSave
        Ruta = Me.CommonDialog1.FileName
        StrCopi = "Producto" & Chr(9) & "Descripcion" & Chr(9) & "Cantidad" & Chr(9) & "Precio_total" & Chr(9) & "Precio_unitario" & Chr(13)
        If Ruta <> "" Then
            NumColum = ListView3.ColumnHeaders.Count
            For Con = 1 To ListView3.ListItems.Count
                StrCopi = StrCopi & ListView3.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView3.ListItems.Item(Con).SubItems(Con2) & Chr(9)
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
    End If
End Sub
Private Sub Image9_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Text1.Text = Item.SubItems(1)
    CLVCLIEN = Item
    NomClien = Item.SubItems(1)
    DesClien = Item.SubItems(2)
    TipoDesc = Item.SubItems(3)
    If CLVCLIEN <> "" And ListView3.ListItems.Count > 0 Then
        Me.Command2.Enabled = True
        Command6.Enabled = True
    Else
        Me.Command2.Enabled = False
        Command6.Enabled = False
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        If CLVCLIEN <> "" Then
            If Me.ListView1.ListItems.Count <> 0 Then
                If Text1.Text <> "" Then
                    Text1.Enabled = False
                    ListView1.Enabled = False
                    Me.Text2.SetFocus
                End If
            End If
        Else
            Me.Text1.SetFocus
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView1_LostFocus()
On Error GoTo ManejaError
    If CLVCLIEN <> "" Then
        If Me.ListView1.ListItems.Count <> 0 Then
            If Text1.Text <> "" Then
                Text1.Enabled = False
                ListView1.Enabled = False
            End If
        End If
    Else
        Me.Text1.SetFocus
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    If Option4.Value = True Then
        Text2.Text = Item
    Else
        Text2.Text = Item.SubItems(1)
    End If
    ClvProd = Item
    DesProd = Item.SubItems(1)
    PreProd = Item.SubItems(2)
    ClasProd = Item.SubItems(4)
    IVA = Item.SubItems(5)
    IMP1 = Item.SubItems(6)
    IMP2 = Item.SubItems(7)
    RET = Item.SubItems(8)
    If Text8.Text <> "" And ClvProd <> "" Then
        Me.Command4.Enabled = True
    Else
        Me.Command4.Enabled = False
    End If
    If CLVCLIEN <> "" And ListView3.ListItems.Count > 0 Then
        Me.Command2.Enabled = True
        Command6.Enabled = True
    Else
        Me.Command2.Enabled = False
        Command6.Enabled = False
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Text8.SetFocus
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    DelInd = Item
    DelDes = Item.SubItems(1)
    DelCan = Item.SubItems(2)
    DelPre = Item.SubItems(3)
    Me.Command3.Enabled = True
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Option1_Click()
On Error GoTo ManejaError
    Text1.Text = ""
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And Text1.Text <> "" Then
        Me.ListView1.SetFocus
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        Dim sBus As String
        If Option2.Value = True Then
            sBus = "SELECT ID_CLIENTE, NOMBRE, DESCUENTO, ID_DESCUENTO FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%' AND VALORACION <> 'E'"
        Else
            sBus = "SELECT ID_CLIENTE, NOMBRE, DESCUENTO, ID_DESCUENTO FROM CLIENTE WHERE ID_CLIENTE = " & Text1.Text & " AND VALORACION <> 'E'"
        End If
        Set tRs = cnn.Execute(sBus)
        With tRs
            ListView1.ListItems.Clear
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                If Not ISNULL(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & ""
                If Not ISNULL(.Fields("DESCUENTO")) Then
                    tLi.SubItems(2) = .Fields("DESCUENTO") & ""
                Else
                    tLi.SubItems(2) = "0.00"
                End If
                If Not ISNULL(.Fields("ID_DESCUENTO")) Then tLi.SubItems(3) = .Fields("ID_DESCUENTO") & ""
                .MoveNext
            Loop
        End With
    End If
    Dim Valido As String
    If Option1.Value = True Then
        Valido = "1234567890"
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    Else
        Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text2_GotFocus()
On Error GoTo ManejaError
    Text2.BackColor = &HFFE1E1
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And Text2.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim tRs2 As ADODB.Recordset
        Dim tLi As ListItem
        Dim sBus As String
        If Option4.Value = True Then
            sBus = "SELECT ALMACEN3.ID_PRODUCTO, ALMACEN3.DESCRIPCION, ALMACEN3.GANANCIA, ALMACEN3.PRECIO_COSTO, ALMACEN3.CLASIFICACION, dbo.ALMACEN3.IVA, ALMACEN3.IMPUESTO1, ALMACEN3.IMPUESTO2, ALMACEN3.RETENCION, ISNULL(EXISTENCIAS.CANTIDAD, 0) As EXISTENCIA FROM ALMACEN3 LEFT OUTER JOIN EXISTENCIAS ON ALMACEN3.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE ALMACEN3.ID_PRODUCTO LIKE '%" & Trim(Text2.Text) & "%'"
        End If
        If Option3.Value = True Then
            sBus = "SELECT ALMACEN3.ID_PRODUCTO, ALMACEN3.DESCRIPCION, ALMACEN3.GANANCIA, ALMACEN3.PRECIO_COSTO, ALMACEN3.CLASIFICACION, dbo.ALMACEN3.IVA, ALMACEN3.IMPUESTO1, ALMACEN3.IMPUESTO2, ALMACEN3.RETENCION, ISNULL(EXISTENCIAS.CANTIDAD, 0) As EXISTENCIA FROM ALMACEN3 LEFT OUTER JOIN EXISTENCIAS ON ALMACEN3.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE ALMACEN3.DESCRIPCION LIKE '%" & Trim(Text2.Text) & "%'"
        End If
        If Option5.Value = True Then
            sBus = "SELECT ID_PRODUCTO FROM ENTRADA_PRODUCTO WHERE CODIGO_BARAS = '" & Text2.Text & "'"
            Set tRs = cnn.Execute(sBus)
            If Not (tRs.EOF And tRs.BOF) Then
                sBus = "SELECT ALMACEN3.ID_PRODUCTO, ALMACEN3.DESCRIPCION, ALMACEN3.GANANCIA, ALMACEN3.PRECIO_COSTO, ALMACEN3.CLASIFICACION, dbo.ALMACEN3.IVA, ALMACEN3.IMPUESTO1, ALMACEN3.IMPUESTO2, ALMACEN3.RETENCION, ISNULL(EXISTENCIAS.CANTIDAD, 0) As EXISTENCIA FROM ALMACEN3 LEFT OUTER JOIN EXISTENCIAS ON ALMACEN3.ID_PRODUCTO = EXISTENCIAS.ID_PRODUCTO WHERE ALMACEN3.ID_PRODUCTO LIKE '%" & tRs.Fields("ID_PRODUCTO") & "%'"
            Else
                MsgBox "EL CODIGO DE BARRAS NO ESTA REGISTRADO, INTENTE OTRO MODO DE BUSQUEDA", vbInformation, "SACC"
            End If
        End If
        If sBus <> "" Then
            Set tRs = cnn.Execute(sBus)
            With tRs
                ListView2.ListItems.Clear
                Do While Not .EOF
                    If Not ISNULL(.Fields("GANANCIA")) And Not ISNULL(.Fields("PRECIO_COSTO")) Then
                        Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                        If Not ISNULL(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion")
                        If Not ISNULL(.Fields("GANANCIA")) And Not ISNULL(.Fields("PRECIO_COSTO")) Then
                            tLi.SubItems(2) = Format((1 + CDbl(.Fields("GANANCIA"))) * CDbl(.Fields("PRECIO_COSTO")), "###,###,##0.00")
                        End If
                        If Not ISNULL(.Fields("EXISTENCIA")) Then tLi.SubItems(3) = .Fields("EXISTENCIA")
                    End If
                    If Not ISNULL(.Fields("CLASIFICACION")) Then tLi.SubItems(4) = .Fields("CLASIFICACION")
                    If Not ISNULL(.Fields("IVA")) Then tLi.SubItems(5) = .Fields("IVA")
                    If Not ISNULL(.Fields("IMPUESTO1")) Then tLi.SubItems(6) = .Fields("IMPUESTO1")
                    If Not ISNULL(.Fields("IMPUESTO2")) Then tLi.SubItems(7) = .Fields("IMPUESTO2")
                    If Not ISNULL(.Fields("RETENCION")) Then tLi.SubItems(8) = .Fields("RETENCION")
                    .MoveNext
                Loop
            End With
        End If
        Me.ListView2.SetFocus
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
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
Private Sub Text6_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ -()/&%@!?*+,"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text8_Change()
On Error GoTo ManejaError
    If Text8.Text <> "" And ClvProd <> "" Then
        Me.Command4.Enabled = True
    Else
        Me.Command4.Enabled = False
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text8_GotFocus()
On Error GoTo ManejaError
    Text8.BackColor = &HFFE1E1
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.Text)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text8_LostFocus()
      Text8.BackColor = &H80000005
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Agregar
        Text2.SetFocus
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
Private Sub Agregar()
On Error GoTo ManejaError
If PreProd <> "" Then
    If Text8.Text = "" Then
        Text8.Text = 1
    End If
    If CLVCLIEN <> "" Then
        Me.Command2.Enabled = True
        Command6.Enabled = True
    End If
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim cant As String
    Dim tLi As ListItem
    Dim PreFin As Double
    Dim PreFinDes As Double
    cant = Text8.Text
    cant = Replace(cant, ",", "")
    Set tLi = ListView3.ListItems.Add(, , ClvProd)
    tLi.SubItems(1) = DesProd
    tLi.SubItems(2) = Text8.Text
    If Combo1.Text = "<NINGUNA>" Or Combo1.Text = "" Then
        If TipoDesc <> "" Then
            sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & TipoDesc & "' AND CLASIFICACION = '" & ClasProd & "'"
            Set tRs1 = cnn.Execute(sBuscar)
            If tRs1.EOF And tRs1.BOF Then
                sBuscar = "SELECT PRECIO_OFERTA FROM PROMOCION WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
                Set tRs1 = cnn.Execute(sBuscar)
                If tRs1.EOF And tRs1.BOF Then
                    Set tRs2 = cnn.Execute(sBuscar)
                    PreFin = Format(CDbl(PreProd) * CDbl(Text8.Text), "###,###,##0.00")
                    If DesClien <> "" Then
                        PreFinDes = PreFin * (100 - Val(DesClien)) / 100
                    End If
                    If Not (tRs2.EOF And tRs2.BOF) Then
                        PreFin = Format(CDbl(tRs2.Fields("PRECIO_OFERTA")) * CDbl(Text8.Text), "###,###,##0.00")
                    End If
                    If DesClien <> "" Then
                        If PreFin > PreFinDes Then
                            PreFin = Format(PreFinDes, "###,###,##0.00")
                        Else
                            PreFin = Format(PreFin, "###,###,##0.00")
                        End If
                    End If
                    Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "###,###,##0.00")
                    tLi.SubItems(3) = Format(CDbl(PreFin), "###,###,##0.00")
                End If
            Else
                PreFin = Format((CDbl(PreProd) - (CDbl(PreProd) * CDbl(tRs1.Fields("PORCENTAJE") / 100))) * CDbl(Text8.Text), "###,###,##0.00")
                Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "###,###,##0.00")
                tLi.SubItems(3) = Format((CDbl(PreProd) - (CDbl(PreProd) * CDbl(tRs1.Fields("PORCENTAJE") / 100))) * CDbl(Text8.Text), "###,###,##0.00")
            End If
        Else
            sBuscar = "SELECT PRECIO_OFERTA FROM PROMOCION WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
            Set tRs2 = cnn.Execute(sBuscar)
            
            PreFin = Format(CDbl(PreProd) * CDbl(Text8.Text), "###,###,##0.00")
            If DesClien <> "" Then
                PreFinDes = PreFin * (100 - Val(DesClien)) / 100
            End If
            If Not (tRs2.EOF And tRs2.BOF) Then
                PreFin = Format(CDbl(tRs2.Fields("PRECIO_OFERTA")) * CDbl(Text8.Text), "###,###,##0.00")
            End If
            If DesClien <> "" Then
                If PreFin > PreFinDes Then
                    PreFin = Format(PreFinDes, "###,###,##0.00")
                Else
                    PreFin = Format(PreFin, "###,###,##0.00")
                End If
            End If
            Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreFin), "###,###,##0.00")
            tLi.SubItems(3) = Format(CDbl(PreFin), "###,###,##0.00")
        End If
    Else
        If Combo1.Text = "LICITACIÓN" Then
            If CLVCLIEN <> "" Then
                sBuscar = "SELECT PRECIO_VENTA FROM LICITACIONES WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND ID_CLIENTE = " & CLVCLIEN
                Set tRs2 = cnn.Execute(sBuscar)
                If Not (tRs2.EOF And tRs2.BOF) Then
                    tLi.SubItems(3) = Format(CDbl(tRs2.Fields("PRECIO_VENTA")) * CDbl(Text8.Text), "###,###,##0.00")
                    Text3.Text = Format(CDbl(Text3.Text) + CDbl(tRs2.Fields("PRECIO_VENTA")) * CDbl(Text8.Text), "###,###,##0.00")
                Else
                    tLi.SubItems(3) = Format(CDbl(PreProd) * CDbl(Text8.Text), "###,###,##0.00")
                    Text3.Text = Format(CDbl(Text3.Text) + CDbl(PreProd) * CDbl(Text8.Text), "###,###,##0.00")
                End If
            Else
                MsgBox "DEBE SELECCIONAR UN CLIENTE PARA DAR ESTE TIPO DE PROMOCIÓN", vbInformation, "SACC"
                Exit Sub
            End If
        Else
            sBuscar = "SELECT PRECIO_OFERTA FROM PROMOCION WHERE FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
            Set tRs2 = cnn.Execute(sBuscar)
            PreFin = Format(CDbl(PreProd) * CDbl(Text8.Text), "###,###,##0.00")
            If DesClien <> "" Then
                PreFinDes = PreFin * (100 - Val(DesClien)) / 100
            End If
            If Not (tRs2.EOF And tRs2.BOF) Then
                PreFin = Format(CDbl((1 + (tRs2.Fields("PRECIO_OFERTA"))) / 100) * CDbl(Text8.Text), "###,###,##0.00")
            End If
            If DesClien <> "" Then
                If PreFin > PreFinDes Then
                    PreFin = PreFinDes
                End If
            End If
            tLi.SubItems(3) = Format(PreFin, "###,###,##0.00")
            Text3.Text = Format(CDbl(Text3.Text) + PreFin, "###,###,##0.00")
        End If
    End If
    If Not ISNULL(IVA) Then
        tLi.SubItems(5) = Format(IVA, "###,###,##0.00")
    Else
        tLi.SubItems(5) = "0.00"
    End If
    If Not ISNULL(IMP1) Then
        tLi.SubItems(6) = Format(IMP1, "###,###,##0.00")
    Else
        tLi.SubItems(6) = "0.00"
    End If
    If Not ISNULL(IMP2) Then
        tLi.SubItems(7) = Format(IMP2, "###,###,##0.00")
    Else
        tLi.SubItems(7) = "0.00"
    End If
    If Not ISNULL(RET) Then
        tLi.SubItems(8) = Format(RET, "###,###,##0.00")
    Else
        tLi.SubItems(8) = "0.00"
    End If
    If IVA <> "" Then Text4.Text = Format(CDbl(Text3.Text) * CDbl(IVA), "###,###,##0.00")
    If IMP1 <> "" Then Text9.Text = Format(CDbl(Text3.Text) * CDbl(IMP1), "###,###,##0.00")
    If IMP2 <> "" Then Text10.Text = Format(CDbl(Text3.Text) * CDbl(IMP2), "###,###,##0.00")
    If RET <> "" Then Text11.Text = Format(CDbl(Text3.Text) * CDbl(RET), "###,###,##0.00")
    Text5.Text = Format(CDbl(Text3.Text) + CDbl(Text4.Text) + CDbl(Text9.Text) + CDbl(Text10.Text) - CDbl(Text11.Text), "###,###,##0.00")
    Text8.Text = ""
    ClvProd = ""
    DesProd = ""
    PreProd = ""
End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ImprimeCotiza()
    On Error GoTo ManejaError
    CommonDialog1.Flags = 64
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    Text3.Text = Replace(Text3.Text, ",", "")
    Text4.Text = Replace(Text4.Text, ",", "")
    Text5.Text = Replace(Text5.Text, ",", "")
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim cant As String
    Dim P_ven As String
    Dim IdVentAut As String
    Dim NumeroRegistros As Integer
    Dim NRegistros As Integer
    Dim Con As Integer
    Dim POSY As Integer
    NumeroRegistros = ListView3.ListItems.Count
    '********************************IMPRIMIR TICKET********************************************
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
    Printer.Print VarMen.Text5(0).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print ""
    Printer.Print ""
    Printer.Print "     CLIENTE : " & NomClien
    Printer.Print "     FECHA : " & Format(Date, "dd/mm/yyyy")
    Printer.Print "     SUCURSAL : " & VarMen.Text4(0).Text
    Printer.Print "     ATENDIDO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
    Printer.Print "     No. FOLIO   : " & IdCotizacion
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("COTIZACION")) / 2
    Printer.Print "COTIZACION"
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    NRegistros = ListView3.ListItems.Count
    POSY = 3400
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 1600
    Printer.Print "Descripcion"
    Printer.CurrentY = POSY
    Printer.CurrentX = 9000
    Printer.Print "Cant."
    Printer.CurrentY = POSY
    Printer.CurrentX = 10500
    Printer.Print "P/U"
    For Con = 1 To NRegistros
        POSY = POSY + 200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print ListView3.ListItems(Con).Text
        Printer.CurrentY = POSY
        Printer.CurrentX = 1600
        Printer.Print ListView3.ListItems(Con).SubItems(1)
        Printer.CurrentY = POSY
        Printer.CurrentX = 9130
        Printer.Print ListView3.ListItems(Con).SubItems(2)
        Printer.CurrentY = POSY
        Printer.CurrentX = 10200
        Printer.Print Val(Replace(ListView3.ListItems(Con).SubItems(3), ",", "")) / Val(Replace(ListView3.ListItems(Con).SubItems(2), ",", ""))
        If POSY >= 14200 Then
            Printer.NewPage
            Printer.Print ""
            Printer.Print ""
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
            Printer.Print VarMen.Text5(0).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
            Printer.Print "R.F.C. " & VarMen.Text5(8).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
            Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
            Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print ""
            Printer.Print ""
            Printer.Print "     CLIENTE : " & NomClien
            Printer.Print "     FECHA : " & Format(Date, "dd/mm/yyyy")
            Printer.Print "     SUCURSAL : " & VarMen.Text4(0).Text
            Printer.Print "     ATENDIDO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
            Printer.Print "     No. FOLIO   : " & IdCotizacion
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("COTIZACION")) / 2
            Printer.Print "COTIZACION"
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            NRegistros = ListView3.ListItems.Count
            POSY = 3400
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print "Producto"
            Printer.CurrentY = POSY
            Printer.CurrentX = 1600
            Printer.Print "Descripcion"
            Printer.CurrentY = POSY
            Printer.CurrentX = 9000
            Printer.Print "Cant."
            Printer.CurrentY = POSY
            Printer.CurrentX = 10500
            Printer.Print "P/U"
        End If
    Next Con
    Printer.Print ""
    Printer.CurrentX = 8700
    Printer.Print "         S U B T O T A L :    " & Text3.Text
    Printer.CurrentX = 8700
    Printer.Print "         I V A                   :    " & Text4.Text
    Printer.CurrentX = 8700
    Printer.Print "         T O T A L           :    " & Text5.Text
    Printer.Print ""
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "                                                                                  COTIZACIÓN SUJETA A CAMBIOS SIN PREVIO AVISO."
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "Comentarios : " & Text6.Text
    Printer.EndDoc
    IdCotizacion = ""
    CommonDialog1.Copies = 1
    Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub FunImpr()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim Moneda As String
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tRs6 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    ConPag = 1
    sBuscar = "SELECT COTIZA_CLIEN.ID_COTIZA_CLIEN, CLIENTE.NOMBRE, USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS AS AGENTE, COTIZA_CLIEN.FECHA, COTIZA_CLIEN.SUBTOTAL, COTIZA_CLIEN.IVA, COTIZA_CLIEN.IMPUESTO1, COTIZA_CLIEN.IMPUESTO2, COTIZA_CLIEN.RETENCION, COTIZA_CLIEN.TOTAL, COTIZA_CLIEN.COMENTARIOS FROM COTIZA_CLIEN INNER JOIN CLIENTE ON COTIZA_CLIEN.ID_CLIENTE = CLIENTE.ID_CLIENTE INNER JOIN USUARIOS ON dbo.COTIZA_CLIEN.ID_AGENTE = USUARIOS.ID_USUARIO WHERE COTIZA_CLIEN.ID_COTIZA_CLIEN = " & IdCotizacion
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        'If Not IsNull(tRs1.Fields("MONEDA")) Then Moneda = tRs1.Fields("MONEDA")
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\Cotizacion.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        ' asi se agregan los logos... solo te falto poner un control IMAGE1 para cargar la imagen en el
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 80, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 90, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 30, 380, 20, 250, "Fecha:" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "Cotizacion : ", "F3", 8, hCenter
        oDoc.WTextBox 60, 390, 20, 250, IdCotizacion, "F2", 11, hCenter
        ' cuadros encabezado
        oDoc.WTextBox 100, 20, 105, 450, "Cordial saludo.", "F2", 10, hLeft ', , , 1, vbBlack
        oDoc.WTextBox 120, 20, 105, 450, "Empresa : " & tRs1.Fields("NOMBRE"), "F2", 10, hLeft ', , , 1, vbBlack
        oDoc.WTextBox 140, 20, 105, 450, "Comentarios : " & tRs1.Fields("COMENTARIOS"), "F2", 10, hLeft ', , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        oDoc.WTextBox 50, 350, 20, 250, "Atendio : " & tRs1.Fields("AGENTE"), "F3", 8, hCenter
        'If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 115, 20, 100, 175, "Empresa : " & tRs2.Fields("NOMBRE"), "F3", 8, hCenter
        'If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 135, 20, 100, 175, "Vendedor : " & tRs2.Fields("DIRECCION"), "F3", 8, hCenter
        Posi = 210
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 20, 20, 90, "Producto", "F2", 8, hCenter, , , 1, vbBlack
        oDoc.WTextBox Posi, 110, 20, 310, "Descripcion", "F2", 8, hCenter, , , 1, vbBlack
        oDoc.WTextBox Posi, 420, 20, 50, "Cantidad", "F2", 8, hCenter, , , 1, vbBlack
        oDoc.WTextBox Posi, 470, 20, 50, "Precio", "F2", 8, hCenter, , , 1, vbBlack
        oDoc.WTextBox Posi, 520, 20, 50, "Monto", "F2", 8, hCenter, , , 1, vbBlack
        Posi = Posi + 20
        ' Linea
        'oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        'oDoc.MoveTo 10, Posi
        'oDoc.WLineTo 580, Posi
        'oDoc.LineStroke
        'Posi = Posi + 6
        ' DETALLE
        sBuscar = "SELECT COTIZA_CLIEN_DETALLE.ID_PRODUCTO, ALMACEN3.DESCRIPCION, COTIZA_CLIEN_DETALLE.CANTIDAD, COTIZA_CLIEN_DETALLE.PRECIO_VENTA FROM COTIZA_CLIEN_DETALLE INNER JOIN ALMACEN3 ON COTIZA_CLIEN_DETALLE.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO WHERE COTIZA_CLIEN_DETALLE.ID_COTIZA_CLIEN = " & tRs1.Fields("ID_COTIZA_CLIEN")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 20, 15, 90, Mid(tRs3.Fields("ID_PRODUCTO"), 1, 18), "F3", 7, hLeft, , , 1, vbBlack
                oDoc.WTextBox Posi, 110, 15, 310, tRs3.Fields("DESCRIPCION"), "F3", 7, hLeft, , , 1, vbBlack
                oDoc.WTextBox Posi, 420, 15, 50, Format(tRs3.Fields("CANTIDAD"), "###,###,##0.00"), "F3", 7, hRight, , , 1, vbBlack
                oDoc.WTextBox Posi, 470, 15, 50, Format(tRs3.Fields("PRECIO_VENTA"), "###,###,##0.00"), "F3", 7, hRight, , , 1, vbBlack
                oDoc.WTextBox Posi, 520, 15, 50, Format(CDbl(tRs3.Fields("PRECIO_VENTA")) * CDbl(tRs3.Fields("CANTIDAD")), "###,###,##0.00"), "F3", 7, hRight, , , 1, vbBlack
                Posi = Posi + 15
                tRs3.MoveNext
                If Posi >= 620 Then
                    oDoc.NewPage A4_Vertical
                    Posi = 50
                    oDoc.WImage 50, 40, 43, 161, "Logo"
                    oDoc.WTextBox 40, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
                    oDoc.WTextBox 60, 224, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hLeft
                    oDoc.WTextBox 80, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
                    oDoc.WTextBox 90, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
                    oDoc.WTextBox 30, 380, 20, 250, "Fecha:" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
                    oDoc.WTextBox 60, 340, 20, 250, "Cotizacion : ", "F3", 8, hCenter
                    oDoc.WTextBox 60, 390, 20, 250, IdCotizacion, "F2", 11, hCenter
                    ' cuadros encabezado
                    oDoc.WTextBox 100, 20, 105, 175, "Cordial saludo.", "F2", 10, hCenter ', , , 1, vbBlack
                    oDoc.WTextBox 100, 205, 105, 175, "Empresa : " & tRs1.Fields("NOMBRE"), "F2", 10, hCenter, , , 1, vbBlack
                    oDoc.WTextBox 100, 390, 105, 175, "Comentarios : " & tRs1.Fields("COMENTARIOS"), "F2", 10, hCenter, , , 1, vbBlack
                    ' LLENADO DE LAS CAJAS
                    oDoc.WTextBox 50, 350, 20, 250, "Atendio : " & tRs1.Fields("AGENTE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hCenter
                    If Not ISNULL(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 115, 20, 100, 175, "Empresa : " & tRs2.Fields("NOMBRE"), "F3", 8, hCenter
                    If Not ISNULL(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 135, 20, 100, 175, "Vendedor : " & tRs2.Fields("DIRECCION"), "F3", 8, hCenter
                    Posi = 210
                    ' ENCABEZADO DEL DETALLE
                    oDoc.WTextBox Posi, 20, 20, 90, "Producto", "F2", 8, hCenter, , , 1, vbBlack
                    oDoc.WTextBox Posi, 110, 20, 310, "Descripcion", "F2", 8, hCenter, , , 1, vbBlack
                    oDoc.WTextBox Posi, 420, 20, 50, "Cantidad", "F2", 8, hCenter, , , 1, vbBlack
                    oDoc.WTextBox Posi, 470, 20, 50, "Precio", "F2", 8, hCenter, , , 1, vbBlack
                    oDoc.WTextBox Posi, 520, 20, 50, "Monto", "F2", 8, hCenter, , , 1, vbBlack
                    Posi = Posi + 20
                    ' Linea
                    'oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    'oDoc.MoveTo 10, Posi
                    'oDoc.WLineTo 580, Posi
                    'oDoc.LineStroke
                    'Posi = Posi + 6
                End If
            Loop
        End If
        ' Linea
        'Posi = Posi + 6
        'oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        'oDoc.MoveTo 10, Posi
        'oDoc.WLineTo 580, Posi
        'oDoc.LineStroke
        'Posi = Posi + 6
        ' TEXTO ABAJO
        Posi = Posi + 10
        oDoc.WTextBox Posi, 470, 15, 50, "SUBTOTAL", "F3", 7, hRight, , , 1, vbBlack
        oDoc.WTextBox Posi, 520, 15, 50, Format(CDbl(tRs1.Fields("SUBTOTAL")), "###,###,##0.00"), "F3", 7, hRight, , , 1, vbBlack
        Posi = Posi + 15
        If Not ISNULL(tRs1.Fields("IVA")) Then
            oDoc.WTextBox Posi, 470, 15, 50, "IVA", "F3", 7, hRight, , , 1, vbBlack
            oDoc.WTextBox Posi, 520, 15, 50, Format(CDbl(tRs1.Fields("IVA")), "###,###,##0.00"), "F3", 7, hRight, , , 1, vbBlack
            Posi = Posi + 15
        End If
        If Not ISNULL(tRs1.Fields("IMPUESTO1")) Then
            oDoc.WTextBox Posi, 470, 15, 50, "IMPUESTO1", "F3", 7, hRight, , , 1, vbBlack
            oDoc.WTextBox Posi, 520, 15, 50, Format(CDbl(tRs1.Fields("IMPUESTO1")), "###,###,##0.00"), "F3", 7, hRight, , , 1, vbBlack
            Posi = Posi + 15
        End If
        If Not ISNULL(tRs1.Fields("IMPUESTO2")) Then
            oDoc.WTextBox Posi, 470, 15, 50, "IMPUESTO2", "F3", 7, hRight, , , 1, vbBlack
            oDoc.WTextBox Posi, 520, 15, 50, Format(CDbl(tRs1.Fields("IMPUESTO2")), "###,###,##0.00"), "F3", 7, hRight, , , 1, vbBlack
            Posi = Posi + 15
        End If
        If Not ISNULL(tRs1.Fields("RETENCION")) Then
            oDoc.WTextBox Posi, 470, 15, 50, "RETENCION", "F3", 7, hRight, , , 1, vbBlack
            oDoc.WTextBox Posi, 520, 15, 50, Format(CDbl(tRs1.Fields("RETENCION")), "###,###,##0.00"), "F3", 7, hRight, , , 1, vbBlack
            Posi = Posi + 15
        End If
        If Not ISNULL(tRs1.Fields("TOTAL")) Then
            oDoc.WTextBox Posi, 470, 15, 50, "TOTAL", "F3", 7, hRight, , , 1, vbBlack
            oDoc.WTextBox Posi, 520, 15, 50, Format(CDbl(tRs1.Fields("TOTAL")), "###,###,##0.00"), "F3", 7, hRight, , , 1, vbBlack
            Posi = Posi + 15
        End If
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub

