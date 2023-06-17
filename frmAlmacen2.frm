VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmacen2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALTA PRODUCTOS ALMACEN 1 Y 2"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   6720
      ScaleHeight     =   6075
      ScaleWidth      =   1875
      TabIndex        =   30
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton cmdCancelar 
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
         Height          =   375
         Left            =   120
         Picture         =   "frmAlmacen2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdGuardar 
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
         Left            =   120
         Picture         =   "frmAlmacen2.frx":29D2
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdMod 
         Caption         =   "Modificar"
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
         Picture         =   "frmAlmacen2.frx":53A4
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Nuevo"
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
         Picture         =   "frmAlmacen2.frx":7D76
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   240
         TabIndex        =   31
         Top             =   3960
         Width           =   975
         Begin VB.Label Label8 
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
            TabIndex        =   32
            Top             =   960
            Width           =   975
         End
         Begin VB.Image cmdCancelar21 
            Height          =   705
            Left            =   120
            MouseIcon       =   "frmAlmacen2.frx":A748
            MousePointer    =   99  'Custom
            Picture         =   "frmAlmacen2.frx":AA52
            Top             =   240
            Width           =   705
         End
      End
   End
   Begin VB.OptionButton opn2 
      Caption         =   "Almacen 2"
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
      Left            =   5160
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
   End
   Begin VB.OptionButton opn1 
      Caption         =   "Almacen 1"
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
      Left            =   5160
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtMarcaAl1 
      Height          =   195
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   5040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtTipoAl1 
      Height          =   195
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   5040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtProductoID 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   885
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtDescuento 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txtCMax 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox txtCMin 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   4
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox txtTipoAl2 
      Height          =   195
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtMarcaAl2 
      Height          =   195
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtMaterial 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   7
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox txtColor 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   8
      Top             =   4920
      Width           =   2895
   End
   Begin VB.ComboBox cboMarca 
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   ">"
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
      Picture         =   "frmAlmacen2.frx":C504
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   495
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Text            =   "SIMPLE"
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox txtIDMarca2 
      Height          =   195
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<"
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
      Picture         =   "frmAlmacen2.frx":EED6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   495
   End
   Begin MSAdodcLib.Adodc adcMarca 
      Height          =   330
      Left            =   2280
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adcAlmacen2 
      Height          =   330
      Left            =   5040
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   5160
      TabIndex        =   16
      Top             =   480
      Width           =   1455
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Siguiente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc adcAlmacen1 
      Height          =   330
      Left            =   5280
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Clave del Producto :"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Descripción :"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Descuento :"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Cantidad Maxima :"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Cantidad Minima :"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo :"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Marca :"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Material :"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Color :"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4920
      Width           =   1575
   End
End
Attribute VB_Name = "frmAlmacen2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bGuardar As Byte
Private Sub cboMarca_DropDown()
On Error GoTo ManejaError
    If Me.adcMarca.Recordset.EOF = False Then
        Do While Me.adcMarca.Recordset.EOF = False
            Me.cboMarca.AddItem Me.adcMarca.Recordset.Fields("MARCA")
            Me.adcMarca.Recordset.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cboMarca_GotFocus()
    cboMarca.BackColor = &HFFE1E1
End Sub
Private Sub cboMarca_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    Else
        KeyAscii = 0
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cboMarca_LostFocus()
On Error GoTo ManejaError
    cboMarca.BackColor = &H80000005
    Select Case bGuardar
    Case 1:
        Me.txtMarcaAl1.Text = Me.cboMarca.Text
    Case 2:
        Me.txtMarcaAl2.Text = Me.cboMarca.Text
    End Select
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cboTipo_DropDown()
On Error GoTo ManejaError
    Me.cboTipo.Clear
    Me.cboTipo.AddItem "SIMPLE", 0
    Me.cboTipo.AddItem "COMPPUESTO", 1
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cboTipo_GotFocus()
    cboTipo.BackColor = &HFFE1E1
End Sub
Private Sub cboTipo_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    Else
        KeyAscii = 0
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cboTipo_LostFocus()
On Error GoTo ManejaError
    cboTipo.BackColor = &H80000005
    Select Case bGuardar
    Case 1:
        Me.txtTipoAl1.Text = Me.cboTipo.Text
    Case 2:
        Me.txtTipoAl2.Text = Me.cboTipo.Text
    End Select
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cmdAdd_Click()
On Error GoTo ManejaError
    If MsgBox("¿SEGURO QUE DESEA AGREGAR UN NUEVO REGISTRO?", vbYesNo + vbDefaultButton1 + vbQuestion, "MENSAJE DEL SISTEMA") = vbYes Then
        Select Case bGuardar
        Case 1:
            Me.adcAlmacen1.Recordset.AddNew
        Case 2:
            Me.adcAlmacen2.Recordset.AddNew
        End Select
        Abrir_Campos_Almacen2
        Me.txtProductoID.SetFocus
        Me.cboMarca.Clear
        Me.cboTipo.Clear
        Me.cmdAdd.Enabled = False
        Me.cmdMod.Enabled = False
        Me.cmdAnterior.Enabled = False
        Me.cmdSiguiente.Enabled = False
        Me.cmdGuardar.Enabled = True
        Me.cboMarca.Enabled = True
        Me.cboTipo.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cmdAnterior_Click()
On Error GoTo ManejaError
    Select Case bGuardar
    Case 1:
        If Me.adcAlmacen1.Recordset.BOF Then
            Me.cmdAnterior.Enabled = True
        Else
            Me.adcAlmacen1.Recordset.MovePrevious
            Me.cboMarca.Text = Me.txtMarcaAl1.Text
            Me.cboTipo.Text = Me.txtTipoAl1.Text
        End If
    Case 2:
        If Me.adcAlmacen2.Recordset.BOF Then
            Me.cmdAnterior.Enabled = True
        Else
            Me.adcAlmacen2.Recordset.MovePrevious
            Me.cboMarca.Text = Me.txtMarcaAl2.Text
            Me.cboTipo.Text = Me.txtTipoAl2.Text
        End If
    End Select
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cmdCancelar_Click()
On Error GoTo ManejaError
    Unload Me
    FrmAltaProdAlm1y2.Show
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub CmdGuardar_Click()
On Error GoTo ManejaError
    If MsgBox("¿SEGURO QUE DESEA GUARDAR?", vbYesNo + vbDefaultButton1 + vbQuestion, "MENSAJE DEL SISTEMA") = vbYes Then
        If Puede_Guardar_Almacen2 = True Then
            Select Case bGuardar
            Case 1:
                Me.adcAlmacen1.Recordset.Update
                Me.adcAlmacen1.Recordset.MoveNext
                Me.adcAlmacen1.Recordset.MovePrevious
            Case 2:
                Me.adcAlmacen2.Recordset.Update
                Me.adcAlmacen2.Recordset.MoveNext
                Me.adcAlmacen2.Recordset.MovePrevious
            End Select
        Else
            Exit Sub
        End If
            MsgBox "GUARDADO", vbInformation, "MENSAJE DEL SISTEMA"
            Me.cmdAdd.Enabled = True
            Me.cmdMod.Enabled = True
            Me.cmdSiguiente.Enabled = True
            Me.cmdAnterior.Enabled = True
            Me.cmdGuardar.Enabled = False
            Me.cboMarca.Enabled = False
            Me.cboTipo.Enabled = False
            Cerrar_Campos_Almacen2
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cmdMod_Click()
On Error GoTo ManejaError
    Select Case bGuardar
    Case 1:
        With Me.adcAlmacen1.Recordset
            If Not (.EOF Or .BOF) Then
                If MsgBox("¿SEGURO QUE DESEA MODIFICAR EL REGISTRO?", vbYesNo + vbDefaultButton1 + vbQuestion, "MENSAJE DEL SISTEMA") = vbYes Then
                Me.cmdAdd.Enabled = False
                Me.cmdMod.Enabled = False
                Me.cmdAnterior.Enabled = False
                Me.cmdSiguiente.Enabled = False
                Me.cmdGuardar.Enabled = True
                Me.cboMarca.Enabled = True
                Me.cboTipo.Enabled = True
                Abrir_Campos_Almacen2
                Me.txtProductoID.SetFocus
                End If
            Else
                MsgBox "¡NO HAY REGISTROS, DE CLICK EN NUEVO!", vbCritical, "MENSAJE DEL SISTEMA"
            End If
        End With
    Case 2:
        With Me.adcAlmacen2.Recordset
            If Not (.EOF Or .BOF) Then
                If MsgBox("¿SEGURO QUE DESEA MODIFICAR EL REGISTRO?", vbYesNo + vbDefaultButton1 + vbQuestion, "MENSAJE DEL SISTEMA") = vbYes Then
                Me.cmdAdd.Enabled = False
                Me.cmdMod.Enabled = False
                Me.cmdAnterior.Enabled = False
                Me.cmdSiguiente.Enabled = False
                Me.cmdGuardar.Enabled = True
                Me.cboMarca.Enabled = True
                Me.cboTipo.Enabled = True
                Abrir_Campos_Almacen2
                Me.txtProductoID.SetFocus
                End If
            Else
                MsgBox "¡NO HAY REGISTROS, DE CLICK EN NUEVO!", vbCritical, "MENSAJE DEL SISTEMA"
            End If
        End With
    End Select
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub

Private Sub cmdCancelar21_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub

Private Sub cmdSiguiente_Click()
On Error GoTo ManejaError
    Select Case bGuardar
    Case 1:
        If Me.adcAlmacen1.Recordset.EOF Then
            Me.cmdSiguiente.Enabled = True
        Else
            Me.adcAlmacen1.Recordset.MoveNext
            Me.cboMarca.Text = Me.txtMarcaAl1.Text
            Me.cboTipo.Text = Me.txtTipoAl1.Text
        End If
    Case 2:
        If Me.adcAlmacen2.Recordset.EOF Then
            Me.cmdSiguiente.Enabled = True
        Else
            Me.adcAlmacen2.Recordset.MoveNext
            Me.cboMarca.Text = Me.txtMarcaAl2.Text
            Me.cboTipo.Text = Me.txtTipoAl2.Text
        End If
    End Select
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo ManejaError
    Const sPathBase As String = "LINUX"
    With Me.adcAlmacen1
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "ALMACEN1"
    End With
    With Me.adcAlmacen2
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "ALMACEN2"
    End With
    With Me.adcMarca
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .RecordSource = "MARCA"
    End With
    Set cnn1 = New ADODB.Connection
    Set rst1 = New ADODB.Recordset
    With cnn1
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    rst1.Open "SELECT * FROM ALMACEN1", cnn1, adOpenDynamic, adLockOptimistic
    Set cnn2 = New ADODB.Connection
    Set rst2 = New ADODB.Recordset
    With cnn2
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    rst2.Open "SELECT * FROM ALMACEN2", cnn2, adOpenDynamic, adLockOptimistic
    Set Me.txtIDMarca2.DataSource = Me.adcMarca
    Me.cboMarca.DataField = "MARCA"
    Me.txtIDMarca2.DataField = "MARCA"
    Me.cmdGuardar.Enabled = False
    Me.cboMarca.Enabled = False
    Me.cboTipo.Enabled = False
    Me.opn1.Value = True
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub opn1_Click()
On Error GoTo ManejaError
    bGuardar = 1
    Set Me.txtCMax.DataSource = Me.adcAlmacen1
    Set Me.txtCMin.DataSource = Me.adcAlmacen1
    Set Me.txtColor.DataSource = Me.adcAlmacen1
    Set Me.txtDescripcion.DataSource = Me.adcAlmacen1
    Set Me.txtDescuento.DataSource = Me.adcAlmacen1
    Set Me.txtMarcaAl1.DataSource = Me.adcAlmacen1
    Set Me.txtMaterial.DataSource = Me.adcAlmacen1
    Set Me.txtProductoID.DataSource = Me.adcAlmacen1
    Set Me.txtTipoAl1.DataSource = Me.adcAlmacen1
    Me.txtCMax.DataField = "C_MAXIMA"
    Me.txtCMin.DataField = "C_MINIMA"
    Me.txtColor.DataField = "COLOR"
    Me.txtDescripcion.DataField = "DESCRIPCION"
    Me.txtDescuento.DataField = "ID_DESCUENTO"
    Me.txtMarcaAl1.DataField = "MARCA"
    Me.txtMaterial.DataField = "MATERIAL"
    Me.txtProductoID.DataField = "ID_PRODUCTO"
    Me.txtTipoAl1.DataField = "TIPO"
    Me.cboMarca.Text = Me.txtMarcaAl1.Text
    Me.cboTipo.Text = Me.txtTipoAl1.Text
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub opn2_Click()
On Error GoTo ManejaError
    bGuardar = 2
    Set Me.txtCMax.DataSource = Me.adcAlmacen2
    Set Me.txtCMin.DataSource = Me.adcAlmacen2
    Set Me.txtColor.DataSource = Me.adcAlmacen2
    Set Me.txtDescripcion.DataSource = Me.adcAlmacen2
    Set Me.txtDescuento.DataSource = Me.adcAlmacen2
    Set Me.txtMarcaAl2.DataSource = Me.adcAlmacen2
    Set Me.txtMaterial.DataSource = Me.adcAlmacen2
    Set Me.txtProductoID.DataSource = Me.adcAlmacen2
    Set Me.txtTipoAl2.DataSource = Me.adcAlmacen2
    Me.txtCMax.DataField = "C_MAXIMA"
    Me.txtCMin.DataField = "C_MINIMA"
    Me.txtColor.DataField = "COLOR"
    Me.txtDescripcion.DataField = "DESCRIPCION"
    Me.txtDescuento.DataField = "ID_DESCUENTO"
    Me.txtMarcaAl2.DataField = "MARCA"
    Me.txtMaterial.DataField = "MATERIAL"
    Me.txtProductoID.DataField = "ID_PRODUCTO"
    Me.txtTipoAl2.DataField = "TIPO"
    Me.cboMarca.Text = Me.txtMarcaAl2.Text
    Me.cboTipo.Text = Me.txtTipoAl2.Text
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtCMax_GotFocus()
On Error GoTo ManejaError
    txtCMax.BackColor = &HFFE1E1
    Me.txtCMax.SelStart = 0
    Me.txtCMax.SelLength = Len(Me.txtCMax.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtCMax_LostFocus()
    txtCMax.BackColor = &H80000005
End Sub
Private Sub txtCMax_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If Me.txtCMax.Locked = True Then MsgBox "Si decea 'modificar' este campo de click en MODIFICAR. Si decea crear un 'nuevo registro' de click en NUEVO. Cuado termine de click en GUARDAR.", vbOKOnly, "MENSAJE DEL SISTEMA"
    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    Else
        Dim Valido As String
        Valido = "1234567890."
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtCMin_GotFocus()
On Error GoTo ManejaError
    txtCMin.BackColor = &HFFE1E1
    Me.txtCMin.SelStart = 0
    Me.txtCMin.SelLength = Len(Me.txtCMin.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtCMin_LostFocus()
    txtCMin.BackColor = &H80000005
End Sub
Private Sub txtCMin_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If Me.txtCMin.Locked = True Then MsgBox "Si decea 'modificar' este campo de click en MODIFICAR. Si decea crear un 'nuevo registro' de click en NUEVO. Cuado termine de click en GUARDAR.", vbOKOnly, "MENSAJE DEL SISTEMA"
    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    Else
        Dim Valido As String
        Valido = "1234567890."
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii > 26 Then
                If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
                End If
            End If
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtColor_GotFocus()
On Error GoTo ManejaError
        txtColor.BackColor = &HFFE1E1
        Me.txtColor.SelStart = 0
        Me.txtColor.SelLength = Len(Me.txtColor.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtColor_LostFocus()
    txtColor.BackColor = &H80000005
End Sub
Private Sub txtColor_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If Me.txtColor.Locked = True Then MsgBox "Si decea 'modificar' este campo de click en MODIFICAR. Si decea crear un 'nuevo registro' de click en NUEVO. Cuado termine de click en GUARDAR.", vbOKOnly, "MENSAJE DEL SISTEMA"
    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    Else
        KeyAscii = Mayusculas(KeyAscii)
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtDescripcion_GotFocus()
On Error GoTo ManejaError
    txtDescripcion.BackColor = &HFFE1E1
    Me.txtDescripcion.SelStart = 0
    Me.txtDescripcion.SelLength = Len(Me.txtDescripcion.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If Me.txtDescripcion.Locked = True Then MsgBox "Si decea 'modificar' este campo de click en MODIFICAR. Si decea crear un 'nuevo registro' de click en NUEVO. Cuado termine de click en GUARDAR.", vbOKOnly, "MENSAJE DEL SISTEMA"
    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    Else
        KeyAscii = Mayusculas(KeyAscii)
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
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtDescuento_GotFocus()
On Error GoTo ManejaError
    txtDescuento.BackColor = &HFFE1E1
    Me.txtDescuento.SelStart = 0
    Me.txtDescuento.SelLength = Len(Me.txtDescuento.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtDescuento_LostFocus()
    txtDescuento.BackColor = &H80000005
End Sub
Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If Me.txtDescuento.Locked = True Then MsgBox "Si decea 'modificar' este campo de click en MODIFICAR. Si decea crear un 'nuevo registro' de click en NUEVO. Cuado termine de click en GUARDAR.", vbOKOnly, "MENSAJE DEL SISTEMA"
    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    Else
        Dim Valido As String
        Valido = "1234567890."
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtMaterial_GotFocus()
On Error GoTo ManejaError
        txtMaterial.BackColor = &HFFE1E1
        Me.txtMaterial.SelStart = 0
        Me.txtMaterial.SelLength = Len(Me.txtMaterial.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtMaterial_LostFocus()
    txtMaterial.BackColor = &H80000005
End Sub
Private Sub txtMaterial_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If Me.txtMaterial.Locked = True Then MsgBox "Si decea 'modificar' este campo de click en MODIFICAR. Si decea crear un 'nuevo registro' de click en NUEVO. Cuado termine de click en GUARDAR.", vbOKOnly, "MENSAJE DEL SISTEMA"
    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    Else
        KeyAscii = Mayusculas(KeyAscii)
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtProductoID_GotFocus()
On Error GoTo ManejaError
    txtProductoID.BackColor = &HFFE1E1
    Me.txtProductoID.SelStart = 0
    Me.txtProductoID.SelLength = Len(Me.txtProductoID.Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub txtProductoID_LostFocus()
    txtProductoID.BackColor = &H80000005
End Sub
Private Sub txtProductoID_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If Me.txtProductoID.Locked = True Then MsgBox "Si decea 'modificar' este campo de click en MODIFICAR. Si decea crear un 'nuevo registro' de click en NUEVO. Cuado termine de click en GUARDAR.", vbOKOnly, "MENSAJE DEL SISTEMA"
    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    Else
        KeyAscii = Mayusculas(KeyAscii)
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
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Public Function Puede_Guardar_Almacen2() As Boolean
On Error GoTo ManejaError
    If Trim(Me.txtProductoID.Text) = "" Then
        MsgBox "POR FAVOR, ESCRIBA EL 'ID' DEL PRODUCTO", vbInformation, "MENSAJE DEL SISTEMA"
        Puede_Guardar_Almacen2 = False
        Me.txtProductoID.SetFocus
        Exit Function
    End If

    If Trim(Me.txtDescripcion.Text) = "" Then
        MsgBox "POR FAVOR, ESCRIBA LA 'DESCRIPCION' DEL PRODUCTO", vbInformation, "MENSAJE DEL SISTEMA"
        Puede_Guardar_Almacen2 = False
        Me.txtDescripcion.SetFocus
        Exit Function
    End If
        
    If Trim(Me.cboTipo.Text) = "" Then
        MsgBox "POR FAVOR, ESCRIBA EL 'TIPO' DEL PRODUCTO", vbInformation, "MENSAJE DEL SISTEMA"
        Puede_Guardar_Almacen2 = False
        Me.cboTipo.SetFocus
        Exit Function
    End If

    If Trim(Me.cboMarca.Text) = "" Then
        MsgBox "POR FAVOR, SELECCIONE LA 'MARCA' DEL PRODUCTO", vbInformation, "MENSAJE DEL SISTEMA"
        Puede_Guardar_Almacen2 = False
        Me.cboMarca.SetFocus
        Exit Function
    End If

    Puede_Guardar_Almacen2 = True
Exit Function
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Function
Sub Abrir_Campos_Almacen2()
On Error GoTo ManejaError
    Me.txtCMax.Locked = False
    Me.txtCMin.Locked = False
    Me.txtColor.Locked = False
    Me.txtDescripcion.Locked = False
    Me.txtDescuento.Locked = False
    Me.txtMaterial.Locked = False
    Me.txtProductoID.Locked = False
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Sub Cerrar_Campos_Almacen2()
On Error GoTo ManejaError
    Me.txtCMax.Locked = True
    Me.txtCMin.Locked = True
    Me.txtColor.Locked = True
    Me.txtDescripcion.Locked = True
    Me.txtDescuento.Locked = True
    Me.txtMaterial.Locked = True
    Me.txtProductoID.Locked = True
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
