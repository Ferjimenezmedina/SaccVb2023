VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmAT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ASISTENCIAS TECNICAS"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   9240
      ScaleHeight     =   6315
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   0
      Width           =   1575
      Begin VB.CommandButton cmdACT 
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
         Picture         =   "frmAT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdListo 
         Caption         =   "Listo"
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
         Picture         =   "frmAT.frx":29D2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   240
         TabIndex        =   4
         Top             =   4920
         Width           =   975
         Begin VB.Image cmdCancelar21 
            Height          =   705
            Left            =   120
            MouseIcon       =   "frmAT.frx":53A4
            MousePointer    =   99  'Custom
            Picture         =   "frmAT.frx":56AE
            Top             =   240
            Width           =   705
         End
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
            TabIndex        =   5
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin VB.TextBox txtComentario 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      MaxLength       =   100
      TabIndex        =   1
      Top             =   5880
      Width           =   8055
   End
   Begin MSComctlLib.ListView lvwAT 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9975
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "GARANTIA"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "A DOMICILIO"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "SUCURSAL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "USUARIO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CLIENTE"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "DESCRIPCION"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "TIPO_ARTICULO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "MODELO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "MARCA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "COMENTARIO COTIZACION"
         Object.Width           =   5080
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Comentarios"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   1095
   End
End
Attribute VB_Name = "frmAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ITMX As ListItem

Private Sub cmdACT_Click()
On Error GoTo ManejaError
    If Trim(Me.txtComentario.Text) = "" Then
        MsgBox "ESCRIBA POR FAVOR EL COMENTARIO", vbInformation, "MENSAJE DEL SISTEMA"
        Me.txtComentario.SetFocus
    Else
        If Me.lvwAT.SelectedItem.Selected = True Then
            deAPTONER.COMENTARIO_AT Trim(Me.txtComentario.Text), Me.lvwAT.SelectedItem
        End If
    End If
    Me.cmdACT.Enabled = False
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub cmdListo_Click()
On Error GoTo ManejaError
    If Me.lvwAT.SelectedItem.Selected = True Then
        deAPTONER.LISTO_AT Me.lvwAT.SelectedItem
    End If
    Llenar_Lista_AT
    Me.cmdListo.Enabled = False
    Me.cmdACT.Enabled = False
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
Private Sub Form_Load()
On Error GoTo ManejaError
    Llenar_Lista_AT
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Sub Llenar_Lista_AT()
On Error GoTo ManejaError
    Me.lvwAT.ListItems.Clear
    deAPTONER.LISTA_AT
    With deAPTONER.rsLISTA_AT
        While Not .EOF
            Set ITMX = Me.lvwAT.ListItems.Add(, , !ID_AS_TEC)
            If Not IsNull(!GARANTIA) Then ITMX.SubItems(1) = Trim(!GARANTIA)
            If Not IsNull(!A_DOMICILIO) Then ITMX.SubItems(2) = Trim(!A_DOMICILIO)
            If Not IsNull(!NOMBRE_USUARIO) Then ITMX.SubItems(3) = Trim(!Sucursal)
            If Not IsNull(!NOMBRE_USUARIO) Then ITMX.SubItems(4) = Trim(!NOMBRE_USUARIO)
            If Not IsNull(!NOMBRE_CLIENTE) Then ITMX.SubItems(5) = Trim(!NOMBRE_CLIENTE)
            If Not IsNull(!FECHA_DEBE_ATENDER) Then ITMX.SubItems(6) = Trim(!FECHA_DEBE_ATENDER)
            If Not IsNull(!DESCRIPCION_PIEZAS) Then ITMX.SubItems(7) = Trim(!DESCRIPCION_PIEZAS)
            If Not IsNull(!TIPO_ARTICULO) Then ITMX.SubItems(8) = Trim(!TIPO_ARTICULO)
            If Not IsNull(!MODELO) Then ITMX.SubItems(9) = Trim(!MODELO)
            If Not IsNull(!Marca) Then ITMX.SubItems(10) = Trim(!Marca)
            If Not IsNull(!COMENTARIOS_COTIZACION) Then ITMX.SubItems(11) = Trim(!COMENTARIO_COTIZACION)
            .MoveNext
            Wend
        .Close
    End With
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub lvwAT_Click()
On Error GoTo ManejaError
    Me.cmdListo.Enabled = False
    Me.cmdACT.Enabled = False
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub lvwAT_DblClick()
On Error GoTo ManejaError
    If Me.lvwAT.ListItems.Count <> 0 Then
        Me.cmdListo.Enabled = True
        Me.cmdACT.Enabled = True
        Me.txtComentario.SetFocus
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
