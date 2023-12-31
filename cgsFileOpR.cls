VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cgsFileOpR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
'cgsFileOpR
'Clase reducida para manejo de ficheros y di�logos comunes, etc.
'
'�Guillermo 'guille' Som, 1997-99 <guille@costasol.net>
'
'Primera tentativa:                             (04:57 25/Ago/97)
'Versi�n reducida, s�lo Abrir                   (10:20 18/Sep/97)
'
'Versi�n reducida:                              (00:53  2/Feb/99)
'Con el DeleteSetting                           (21:25 21/Feb/99)
'----------------------------------------------------------------------------------
'Esta clase tiene las siguientes funciones/m�todos:
'   ShowOpen                Di�logo com�n Abrir
'   GetSetting/LeerIni      Leer ficheros INIs
'   SaveSetting/GuardarIni  Guardar en ficheros INIs
'   DeleteSetting           Borrar secciones o claves de ficheros INIs
'   QuitarCaracterEx        Quitar/cambiar caracteres de una cadena
'   AddBackSlash            A�adir \ al final si no la tiene
'   QuitarBackSlash         Quitar el �ltimo \ de una cadena
'   SplitPath               Partir el path indicado en diferentes partes
'----------------------------------------------------------------------------------
Option Explicit
Private sFilter As String
'Esta propiedad har� referencia al hWnd de un Form
Public hWnd As Long
'Propiedades gen�ricas de los di�logos comunes
Public DialogTitle As String
Public CancelError As Boolean
Public Flags As eOFN 'Long
'Propiedades para Abrir y Guardar como
Public DefaultExt As String
Public FileName As String
Public FileTitle As String
'Public Filter As String
Public FilterIndex As Long
Public InitDir As String
'Public MaxFileSize As Long (ser� 260)
'----------------------------------------------------------------------------
'Estructura de datos para Abrir y Guardar como...
'----------------------------------------------------------------------------
Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenFilename As OPENFILENAME) As Long
'Constantes para las funciones de archivos
Public Enum eOFN
    'Tama�o m�ximo de un nombre de archivo (incluyendo el path)
    MAX_PATH = 260
    'Constantes para el di�logo de archivos
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
    OFN_EXPLORER = &H80000                         '  new look commdlg
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
End Enum
'Declaraciones para leer ficheros INI
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Sub DeleteSetting(ByVal sIniFile As String, ByVal sSection As String, Optional ByVal sKey As String = "")
    ' Borrar una clave o entrada de un fichero INI                      (16/Feb/99)
    ' Si no se indica sKey, se borrar� la secci�n indicada en sSection
    ' En otro caso, se supone que es la entrada (clave) lo que se quiere borrar
    If Len(sKey) = 0 Then
        ' Borrar una secci�n
        Call WritePrivateProfileString(sSection, 0&, 0&, sIniFile)
    Else
        ' Borrar una entrada
        Call WritePrivateProfileString(sSection, sKey, 0&, sIniFile)
    End If
End Sub
Public Sub SplitPath(ByVal sTodo As String, sPath As String, Optional vNombre, Optional vExt)
    '----------------------------------------------------------------
    'Divide el nombre recibido en la ruta, nombre y extensi�n
    '(c)Guillermo Som, 1997                         ( 1/Mar/97)
    'Esta rutina aceptar� los siguientes par�metros:
    'sTodo      Valor de entrada con la ruta completa
    'Devolver� la informaci�n en:
    'sPath      Ruta completa, incluida la unidad
    'vNombre    Nombre del archivo incluida la extensi�n
    'vExt       Extensi�n del archivo (sin el punto)
    'Los par�metros opcionales s�lo se usar�n si se han especificado
    '----------------------------------------------------------------
    Dim bNombre As Boolean      'Flag para saber si hay que devolver el nombre
    Dim i As Integer
    If Not IsMissing(vNombre) Then
        bNombre = True
        vNombre = sTodo
    End If
    'La extensi�n se debe buscar desde atr�s, ya que el nombre puede    (13/Ene/99)
    'contener puntos en el nombre.
    If Not IsMissing(vExt) Then
        vExt = ""
        For i = Len(sTodo) To 1 Step -1
            If Mid$(sTodo, i, 1) = "." Then
                vExt = Mid$(sTodo, i + 1)
                Exit For
            End If
        Next
        If Len(vExt) = 0 Then
            vExt = "*"
        End If
    End If
    sPath = ""
    'Asignar el path
    For i = Len(sTodo) To 1 Step -1
        If Mid$(sTodo, i, 1) = "\" Then
            sPath = Left$(sTodo, i - 1)
            'Si hay que devolver el nombre
            If bNombre Then
                vNombre = Mid$(sTodo, i + 1)
            End If
            Exit For
        End If
    Next
End Sub
Public Function QuitarCaracterEx(ByVal sValor As String, ByVal sCaracter As String, Optional ByVal sPoner) As String
    '----------------------------------------------------------
    ' Cambiar/Quitar caracteres                     (17/Sep/97)
    ' Si se especifica sPoner, se cambiar� por ese car�cter
    '
    'Esta versi�n permite cambiar los caracteres    (17/Sep/97)
    'y sustituirlos por el/los indicados
    'a diferencia de QuitarCaracter, no se buscan uno a uno,
    'sino todos juntos
    '----------------------------------------------------------
    Dim i As Long
    Dim sCh As String
    Dim bPoner As Boolean
    Dim iLen As Long
    bPoner = False
    If Not IsMissing(sPoner) Then
        sCh = sPoner
        bPoner = True
    End If
    iLen = Len(sCaracter)
    If iLen = 0 Then
        QuitarCaracterEx = sValor
        Exit Function
    End If
    'Si el caracter a quitar/cambiar es Chr$(0), usar otro m�todo
    If Asc(sCaracter) = 0 Then
        'Quitar todos los chr$(0) del final
        Do While Right$(sValor, 1) = Chr$(0)
            sValor = Left$(sValor, Len(sValor) - 1)
            If Len(sValor) = 0 Then Exit Do
        Loop
        iLen = 1
        Do
            i = InStr(iLen, sValor, sCaracter)
            If i Then
                If bPoner Then
                    sValor = Left$(sValor, i - 1) & sCh & Mid$(sValor, i + 1)
                Else
                    sValor = Left$(sValor, i - 1) & Mid$(sValor, i + 1)
                End If
                iLen = i
            Else
                'ya no hay m�s, salir del bucle
                Exit Do
            End If
        Loop
    Else
        i = 1
        Do While i <= Len(sValor)
            'Debug.Print Mid$(sValor, i, 1); Asc(Mid$(sValor, i, 1));
            If Mid$(sValor, i, iLen) = sCaracter Then
                If bPoner Then
                    sValor = Left$(sValor, i - 1) & sCh & Mid$(sValor, i + iLen)
                    i = i - 1
                Else
                    sValor = Left$(sValor, i - 1) & Mid$(sValor, i + iLen)
                End If
            End If
            i = i + 1
        Loop
    End If
    QuitarCaracterEx = sValor
End Function
Public Function ShowOpen(Optional ByVal vFileName, Optional ByVal vTitle, Optional ByVal vFilter, Optional vFlags, Optional ByVal vhWnd) As Boolean
    '----------------------------------------------------------
    'M�todo para mostrar el cuadro de di�logo de Abrir
    '(c) Guillermo Som Cerezo                  24/Oct/93
    'Convertido en objeto (clase)                   (25/Ago/97)
    'Los par�metros opcionales especificar�n:
    '   vFileName   El nombre del archivo
    '   vTitle      T�tulo del cuadro de di�logo
    '   vFilter     Extensiones
    '   vFlags      Los flags
    '   vhWnd       El hWnd del Form
    '----------------------------------------------------------
    Dim resultado As Long
    Dim ofn As OPENFILENAME
    Err.Clear
    Err.Number = 0
    If Not IsMissing(vFileName) Then FileName = CStr(vFileName)
    If Not IsMissing(vhWnd) Then hWnd = CLng(vhWnd)
    If Not IsMissing(vFilter) Then Me.Filter = CStr(vFilter)
    If Not IsMissing(vTitle) Then DialogTitle = CStr(vTitle)
    If Not IsMissing(vFlags) Then Flags = CLng(vFlags)
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = hWnd
        .hInstance = 0
        If Len(sFilter) = 0 Then sFilter = "Todos los Archivos (*.*)" & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
        .lpstrFilter = sFilter
        .nFilterIndex = FilterIndex
        .lpstrFile = Left$(FileName & String$(MAX_PATH, 0), MAX_PATH)
        .nMaxFile = MAX_PATH
        .nFileOffset = 0
        .nFileExtension = 0
        .lpstrDefExt = DefaultExt
        .lpstrFileTitle = Left$(FileTitle & String$(260, 0), MAX_PATH)
        .nMaxFileTitle = MAX_PATH
        .lpstrInitialDir = Left$(InitDir & String$(260, 0), MAX_PATH)
        'Nombres largos y estilo explorer           (21/Oct/97)
        'y otros valore "obvios"
        'Flags = Flags Or OFN_LONGNAMES Or OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
        .Flags = Flags
        If Len(DialogTitle) = 0 Then    'Si no se especifica el t�tulo
            DialogTitle = "Abrir"
        End If
        .lpstrTitle = DialogTitle
        .nFileOffset = 0
        .lpstrDefExt = 0
        .lCustData = 0
        .lpfnHook = 0
        .lpTemplateName = 0
    End With
    resultado = GetOpenFileName(ofn)
    If resultado <> 0 Then
        If Flags And OFN_ALLOWMULTISELECT Then
            'Si est� multiselect, se separan los nombres con Chr$(0)
            FileName = Chr$(34) & Trim$(QuitarCaracterEx(ofn.lpstrFile, Chr$(0), Chr$(34) & " " & Chr$(34))) & Chr$(34)
            FileTitle = Chr$(34) & Trim$(QuitarCaracterEx(ofn.lpstrFileTitle, Chr$(0), Chr$(34) & " " & Chr$(34))) & Chr$(34)
            InitDir = Chr$(34) & Trim$(QuitarCaracterEx(ofn.lpstrInitialDir, Chr$(0), Chr$(34) & " " & Chr$(34))) & Chr$(34)
        Else
            FileName = Left$(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr$(0)) - 1)
            FileTitle = Left$(ofn.lpstrFileTitle, InStr(ofn.lpstrFileTitle, Chr$(0)) - 1)
            InitDir = Left$(ofn.lpstrInitialDir, InStr(ofn.lpstrInitialDir, Chr$(0)) - 1)
        End If
    Else
        If CancelError Then
            'Err.Raise 32755, "cComDlgR.ShowOpen", "Error en Abrir (clase cComDlg)"
            With Err
                .Source = "cComDlgR.ShowOpen"
                .Number = 32755
                .Description = "Error en Abrir (clase cComDlgR)"
            End With
        End If
    End If
    'Devuelve True si se puede abrir
    ShowOpen = (resultado <> 0)
End Function
Public Property Let Action(vNewValue As Integer)
    '0   Ninguna acci�n.
    '1   Muestra el cuadro de di�logo Abrir.
    '2   Muestra el cuadro de di�logo Guardar como.
    '3   Muestra el cuadro de di�logo Color.
    '4   Muestra el cuadro de di�logo Fuente.
    '5   Muestra el cuadro de di�logo Impresora.
    '6   Ejecuta WINHELP.EXE.
    Select Case vNewValue
    Case 1: ShowOpen
    'Case 2: ShowSave
    'Case 3: 'ShowColor
    'Case 4: 'ShowFont
    'Case 5: 'ShowPrinter
    'Case 6: 'ShowHelp
    'Case Else
        'nada que mostrar
    End Select
End Property
Public Property Let Filter(ByVal sNewFilter As String)
    'Procesar el par�metro para convertirlo a formato C,
    'Se usar� | como separador.
    Dim i As Integer, j As Integer
    Dim sTmp As String
    sTmp = ""
    If InStr(sNewFilter, "|") Then
        sNewFilter = Trim$(sNewFilter)
        If Right$(sNewFilter, 1) <> "|" Then
            sNewFilter = sNewFilter & "|"
        End If
        Do
            i = InStr(sNewFilter, "|")
            If i Then
                sTmp = sTmp & Left$(sNewFilter, i - 1) & Chr$(0)
                sNewFilter = Mid$(sNewFilter, i + 1)
            Else
                Exit Do
            End If
        Loop While i
        If Right$(sTmp, 1) = Chr$(0) Then
            sNewFilter = sTmp & Chr$(0)
        Else
            sNewFilter = sTmp & Chr$(0) & Chr$(0)
        End If
    ElseIf InStr(sNewFilter, Chr$(0)) = 0 Then
        sNewFilter = ""
    End If
    sFilter = sNewFilter
End Property
Public Function LeerIni(ByVal lpFileName As String, ByVal lpAppName As String, ByVal lpKeyName As String, Optional ByVal vDefault As Variant) As String
    If IsMissing(vDefault) Then
        LeerIni = Me.GetSetting(lpFileName, lpAppName, lpKeyName)
    Else
        LeerIni = Me.GetSetting(lpFileName, lpAppName, lpKeyName, vDefault)
    End If
End Function
Public Function GetSetting(ByVal lpFileName As String, ByVal lpAppName As String, ByVal lpKeyName As String, Optional ByVal vDefault As Variant) As String
    'Nota 14/Abr/98,    antes el valor devuelto era Variant
    '                   Lo he cambiado para compatibilizarlo
    '                   con LeerIni
    'Los par�metros son:
    'lpFileName:    La Aplicaci�n (fichero INI)
    'lpAppName:     La secci�n que suele estar entrre corchetes
    'lpKeyName:     Clave
    'vDefault:      Valor opcional que devolver�
    '               si no se encuentra la clave.
    Dim lpString As String
    Dim LTmp As Long
    Dim sRetVal As String
    'Si no se especifica el valor por defecto,
    'asignar incialmente una cadena vac�a
    If IsMissing(vDefault) Then
        lpString = ""
    Else
        lpString = CStr(vDefault)
    End If
    sRetVal = String$(255, 0)
    LTmp = GetPrivateProfileString(lpAppName, lpKeyName, lpString, sRetVal, Len(sRetVal), lpFileName)
    If LTmp = 0 Then
        GetSetting = lpString
    Else
        GetSetting = Left(sRetVal, LTmp)
    End If
End Function
Public Sub GuardarIni(ByVal lpFileName As String, ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String)
    Me.SaveSetting lpFileName, lpAppName, lpKeyName, lpString
End Sub
Public Sub SaveSetting(ByVal lpFileName As String, ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String)
    'Guarda los datos de configuraci�n
    'Los par�metros son los mismos que en LeerIni
    'Siendo lpString el valor a guardar
    Dim LTmp As Long
    LTmp = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub
Public Function AddBackSlash(ByVal sPath As String) As String
    'Si no tiene la barra de directorio a�adirsela
    'Nota: Para quitarla, ver QuitarBackSlah            (13/abr/98)
        If Len(sPath) Then                                              '                       (30/Ene/99)
        If Right$(sPath, 1) <> "\" Then
            sPath = sPath & "\"
        End If
    End If
    AddBackSlash = sPath
End Function
Public Function QuitarBackSlash(ByVal sPath As String) As String
    'Quitarle el \ del final
    'Para a�adirsela, ver AddBackSlash
    If Right$(sPath, 1) = "\" Then
        sPath = Left$(sPath, Len(sPath) - 1)
    End If
    QuitarBackSlash = sPath
End Function

