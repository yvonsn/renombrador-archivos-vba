VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRenombrador 
   Caption         =   "YSN - Renombrador de archivos"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12165
   OleObjectBlob   =   "frmRenombrador.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmRenombrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ArchivosOriginales() As String
Private ArchivosNuevos() As String
Private ArchivosTemporales() As String
Private VistaPreviaGenerada As Boolean

Private Sub UserForm_Initialize()
    If Not CarpetaExiste(CARPETA_TRABAJO) Then
        MsgBox "No existe la carpeta de trabajo:" & vbCrLf & _
               CARPETA_TRABAJO, vbExclamation, "Carpeta no encontrada"
        Unload Me
        Exit Sub
    End If

    lblCarpeta.Caption = CARPETA_TRABAJO

    With LstPreview
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "220 pt;220 pt"
    End With

    optNumerar.Value = True
    ActualizarControles
    InvalidarVistaPrevia
End Sub

Private Sub optNumerar_Click()
    ActualizarControles
    InvalidarVistaPrevia
End Sub

Private Sub optTexto_Click()
    ActualizarControles
    InvalidarVistaPrevia
End Sub

Private Sub optCapitalizar_Click()
    ActualizarControles
    InvalidarVistaPrevia
End Sub

Private Sub txtNumero_Change()
    InvalidarVistaPrevia
End Sub

Private Sub txtTexto_Change()
    InvalidarVistaPrevia
End Sub

Private Sub ActualizarControles()
    lblNumero.Visible = optNumerar.Value
    txtNumero.Visible = optNumerar.Value
    lblTexto.Visible = optTexto.Value
    txtTexto.Visible = optTexto.Value
End Sub

Private Sub InvalidarVistaPrevia()
    VistaPreviaGenerada = False
    LstPreview.Clear
    Erase ArchivosOriginales
    Erase ArchivosNuevos
    Erase ArchivosTemporales
End Sub

Private Function CarpetaExiste(ByVal Ruta As String) As Boolean
    On Error GoTo NoExiste
    CarpetaExiste = ((GetAttr(Ruta) And vbDirectory) = vbDirectory)
    Exit Function
NoExiste:
    CarpetaExiste = False
End Function

Private Function ArchivoExiste(ByVal RutaCompleta As String) As Boolean
    Dim SistemaArchivos As Object

    On Error GoTo NoExiste
    Set SistemaArchivos = CreateObject("Scripting.FileSystemObject")
    ArchivoExiste = SistemaArchivos.FileExists(RutaCompleta)
    Exit Function
NoExiste:
    ArchivoExiste = False
End Function

Private Function ObtenerArchivos() As Collection
    Dim Resultado As New Collection
    Dim Archivo As String
    Dim RutaCompleta As String

    Archivo = Dir$(CARPETA_TRABAJO & "*", vbNormal Or vbHidden Or vbSystem Or vbReadOnly)

    Do While Len(Archivo) > 0
        RutaCompleta = CARPETA_TRABAJO & Archivo

        If (GetAttr(RutaCompleta) And vbDirectory) = 0 Then
            If Left$(Archivo, 12) <> "__REN_TMP__" Then
                Resultado.Add Archivo
            End If
        End If

        Archivo = Dir$
    Loop

    Set ObtenerArchivos = Resultado
End Function

Private Sub OrdenarArray(ByRef Datos() As String)
    Dim i As Long
    Dim j As Long
    Dim Temporal As String

    For i = LBound(Datos) To UBound(Datos) - 1
        For j = i + 1 To UBound(Datos)
            If StrComp(Datos(i), Datos(j), vbTextCompare) > 0 Then
                Temporal = Datos(i)
                Datos(i) = Datos(j)
                Datos(j) = Temporal
            End If
        Next j
    Next i
End Sub

Private Function SepararNombre(ByVal Archivo As String, _
                               ByRef NombreBase As String, _
                               ByRef Extension As String) As Boolean
    Dim PosicionPunto As Long

    PosicionPunto = InStrRev(Archivo, ".")

    If PosicionPunto <= 1 Or PosicionPunto = Len(Archivo) Then
        NombreBase = Archivo
        Extension = vbNullString
    Else
        NombreBase = Left$(Archivo, PosicionPunto - 1)
        Extension = Mid$(Archivo, PosicionPunto)
    End If

    SepararNombre = (Len(NombreBase) > 0)
End Function

Private Function NumeroInicialValido(ByRef Numero As Long) As Boolean
    Dim Texto As String
    Dim Valor As Double

    Texto = Trim$(txtNumero.Text)

    If Len(Texto) = 0 Or Not IsNumeric(Texto) Then
        MsgBox "Introduzca un número inicial válido.", vbExclamation, "Dato incorrecto"
        txtNumero.SetFocus
        Exit Function
    End If

    Valor = CDbl(Texto)

    If Valor < 0 Or Valor > 2147483647# Or Valor <> Fix(Valor) Then
        MsgBox "El número inicial debe ser un entero entre 0 y 2147483647.", _
               vbExclamation, "Dato incorrecto"
        txtNumero.SetFocus
        Exit Function
    End If

    Numero = CLng(Valor)
    NumeroInicialValido = True
End Function

Private Function TextoValido(ByVal Texto As String) As Boolean
    Dim CaracteresProhibidos As Variant
    Dim Caracter As Variant

    Texto = Trim$(Texto)

    If Len(Texto) = 0 Then
        MsgBox "Introduzca el texto que desea añadir.", vbExclamation, "Dato incorrecto"
        txtTexto.SetFocus
        Exit Function
    End If

    CaracteresProhibidos = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")

    For Each Caracter In CaracteresProhibidos
        If InStr(1, Texto, CStr(Caracter), vbBinaryCompare) > 0 Then
            MsgBox "El texto contiene un carácter no permitido: " & CStr(Caracter), _
                   vbExclamation, "Dato incorrecto"
            txtTexto.SetFocus
            Exit Function
        End If
    Next Caracter

    If Right$(Texto, 1) = "." Or Right$(Texto, 1) = " " Then
        MsgBox "El texto no puede terminar en un punto o un espacio.", _
               vbExclamation, "Dato incorrecto"
        txtTexto.SetFocus
        Exit Function
    End If

    TextoValido = True
End Function

Private Function NombreWindowsValido(ByVal Nombre As String) As Boolean
    NombreWindowsValido = (Len(Nombre) > 0 And Len(Nombre) <= 255)
End Function

Private Function EsNombreNumerico(ByVal Archivo As String, ByRef Numero As Long) As Boolean
    Dim Base As String
    Dim Extension As String
    Dim Valor As Double

    Call SepararNombre(Archivo, Base, Extension)

    If Len(Base) = 0 Or Not IsNumeric(Base) Then Exit Function

    Valor = CDbl(Base)
    If Valor < 0 Or Valor > 2147483647# Or Valor <> Fix(Valor) Then Exit Function

    Numero = CLng(Valor)
    EsNombreNumerico = True
End Function

Private Function InformeHuecos(ByRef Archivos() As String) As String
    Dim Numeros() As Long
    Dim Cantidad As Long
    Dim i As Long
    Dim j As Long
    Dim Numero As Long
    Dim Temporal As Long
    Dim Faltante As Long
    Dim Informe As String
    Dim TotalFaltantes As Long

    For i = LBound(Archivos) To UBound(Archivos)
        If EsNombreNumerico(Archivos(i), Numero) Then
            Cantidad = Cantidad + 1
            ReDim Preserve Numeros(1 To Cantidad)
            Numeros(Cantidad) = Numero
        End If
    Next i

    If Cantidad < 2 Then Exit Function

    For i = 1 To Cantidad - 1
        For j = i + 1 To Cantidad
            If Numeros(i) > Numeros(j) Then
                Temporal = Numeros(i)
                Numeros(i) = Numeros(j)
                Numeros(j) = Temporal
            End If
        Next j
    Next i

    For i = 1 To Cantidad - 1
        If Numeros(i + 1) > Numeros(i) + 1 Then
            For Faltante = Numeros(i) + 1 To Numeros(i + 1) - 1
                TotalFaltantes = TotalFaltantes + 1
                If TotalFaltantes <= 30 Then
                    If Len(Informe) > 0 Then Informe = Informe & ", "
                    Informe = Informe & CStr(Faltante)
                End If
            Next Faltante
        End If
    Next i

    If TotalFaltantes > 30 Then
        Informe = Informe & " ... (" & TotalFaltantes & " números ausentes)"
    End If

    InformeHuecos = Informe
End Function

Private Function ValidarDestinos(ByRef Originales() As String, _
                                 ByRef Nuevos() As String) As Boolean
    Dim Destinos As Object
    Dim Origenes As Object
    Dim i As Long
    Dim Clave As String

    Set Destinos = CreateObject("Scripting.Dictionary")
    Set Origenes = CreateObject("Scripting.Dictionary")
    Destinos.CompareMode = vbTextCompare
    Origenes.CompareMode = vbTextCompare

    For i = LBound(Originales) To UBound(Originales)
        Origenes(Originales(i)) = True
    Next i

    For i = LBound(Nuevos) To UBound(Nuevos)
        If Not NombreWindowsValido(Nuevos(i)) Then
            MsgBox "El nombre resultante no es válido:" & vbCrLf & Nuevos(i), _
                   vbCritical, "Nombre no válido"
            Exit Function
        End If

        Clave = Nuevos(i)

        If Destinos.Exists(Clave) Then
            MsgBox "Dos archivos producirían el mismo nombre:" & vbCrLf & Clave, _
                   vbCritical, "Nombre duplicado"
            Exit Function
        End If

        Destinos.Add Clave, True

        If Not Origenes.Exists(Clave) Then
            If ArchivoExiste(CARPETA_TRABAJO & Clave) Then
                MsgBox "Ya existe un archivo con el nombre de destino:" & vbCrLf & Clave, _
                       vbCritical, "Conflicto de nombres"
                Exit Function
            End If
        End If
    Next i

    ValidarDestinos = True
End Function

Private Sub cmdVistaPrevia_Click()
    Dim Coleccion As Collection
    Dim Datos() As String
    Dim i As Long
    Dim Numero As Long
    Dim NombreBase As String
    Dim Extension As String
    Dim TextoAgregar As String
    Dim Huecos As String

    InvalidarVistaPrevia

    If Not CarpetaExiste(CARPETA_TRABAJO) Then
        MsgBox "No existe la carpeta de trabajo:" & vbCrLf & CARPETA_TRABAJO, _
               vbExclamation, "Carpeta no encontrada"
        Exit Sub
    End If

    If optNumerar.Value Then
        If Not NumeroInicialValido(Numero) Then Exit Sub
    ElseIf optTexto.Value Then
        TextoAgregar = Trim$(txtTexto.Text)
        If Not TextoValido(TextoAgregar) Then Exit Sub
    End If

    Set Coleccion = ObtenerArchivos()

    If Coleccion.Count = 0 Then
        MsgBox "No existen archivos en la carpeta de trabajo.", _
               vbInformation, "Carpeta vacía"
        Exit Sub
    End If

    If optNumerar.Value Then
        If CDbl(Numero) + CDbl(Coleccion.Count) - 1# > 2147483647# Then
            MsgBox "El número inicial es demasiado alto para la cantidad de archivos.", _
                   vbExclamation, "Numeración fuera de rango"
            Exit Sub
        End If
    End If

    ReDim Datos(1 To Coleccion.Count)
    ReDim ArchivosOriginales(1 To Coleccion.Count)
    ReDim ArchivosNuevos(1 To Coleccion.Count)

    For i = 1 To Coleccion.Count
        Datos(i) = CStr(Coleccion(i))
    Next i

    OrdenarArray Datos

    If optNumerar.Value Then
        Huecos = InformeHuecos(Datos)
        If Len(Huecos) > 0 Then
            MsgBox "Se han detectado números ausentes en la secuencia:" & vbCrLf & _
                   Huecos & vbCrLf & vbCrLf & _
                   "La vista previa mostrará la nueva numeración correlativa.", _
                   vbInformation, "Verificación de secuencia"
        Else
            MsgBox "No se han detectado saltos en los nombres numéricos analizados." & _
                   vbCrLf & "La vista previa mostrará la nueva numeración.", _
                   vbInformation, "Verificación de secuencia"
        End If
    End If

    For i = LBound(Datos) To UBound(Datos)
        ArchivosOriginales(i) = Datos(i)

        If Not SepararNombre(Datos(i), NombreBase, Extension) Then
            MsgBox "No se pudo interpretar el archivo:" & vbCrLf & Datos(i), _
                   vbCritical, "Nombre no válido"
            InvalidarVistaPrevia
            Exit Sub
        End If

        If optNumerar.Value Then
            ArchivosNuevos(i) = CStr(Numero) & Extension
            Numero = Numero + 1
        ElseIf optTexto.Value Then
            ArchivosNuevos(i) = NombreBase & "-" & TextoAgregar & Extension
        ElseIf optCapitalizar.Value Then
            ArchivosNuevos(i) = Replace$( _
                StrConv(Replace$(NombreBase, "-", " "), vbProperCase), _
                " ", "-") & Extension
        End If

        LstPreview.AddItem ArchivosOriginales(i)
        LstPreview.List(LstPreview.ListCount - 1, 1) = ArchivosNuevos(i)
    Next i

    If Not ValidarDestinos(ArchivosOriginales, ArchivosNuevos) Then
        InvalidarVistaPrevia
        Exit Sub
    End If

    VistaPreviaGenerada = True
End Sub

Private Function CrearNombresTemporales(ByVal Cantidad As Long) As Boolean
    Dim i As Long
    Dim Intento As Long
    Dim Candidato As String
    Dim Identificador As String

    ReDim ArchivosTemporales(1 To Cantidad)
    Identificador = Format$(Now, "yyyymmddhhnnss")

    For i = 1 To Cantidad
        Intento = 0
        Do
            Intento = Intento + 1
            Candidato = "__REN_TMP__" & Identificador & "_" & _
                        Format$(i, "00000") & "_" & Format$(Intento, "000") & ".tmp"
        Loop While ArchivoExiste(CARPETA_TRABAJO & Candidato)

        ArchivosTemporales(i) = Candidato
    Next i

    CrearNombresTemporales = True
End Function

Private Sub cmdEjecutar_Click()
    Dim i As Long
    Dim Fase1Completada As Long
    Dim Fase2Completada As Long
    Dim Respuesta As VbMsgBoxResult
    Dim ErrorNumero As Long
    Dim ErrorDescripcion As String

    If Not VistaPreviaGenerada Then
        MsgBox "Primero genere y revise la vista previa.", _
               vbExclamation, "Vista previa necesaria"
        Exit Sub
    End If

    If Not ValidarDestinos(ArchivosOriginales, ArchivosNuevos) Then Exit Sub

    For i = LBound(ArchivosOriginales) To UBound(ArchivosOriginales)
        If Not ArchivoExiste(CARPETA_TRABAJO & ArchivosOriginales(i)) Then
            MsgBox "El contenido de la carpeta ha cambiado." & vbCrLf & _
                   "Genere de nuevo la vista previa.", vbExclamation, "Archivos modificados"
            InvalidarVistaPrevia
            Exit Sub
        End If
    Next i

    Respuesta = MsgBox( _
        "Se cambiarán " & CStr(UBound(ArchivosOriginales)) & " archivos." & vbCrLf & _
        "¿Desea continuar?", _
        vbQuestion Or vbYesNo Or vbDefaultButton2, _
        "Confirmar cambios")

    If Respuesta <> vbYes Then Exit Sub

    Call CrearNombresTemporales(UBound(ArchivosOriginales))

    On Error GoTo ControlError

    ' Fase 1: todos los originales pasan a nombres temporales únicos.
    For i = LBound(ArchivosOriginales) To UBound(ArchivosOriginales)
        Name CARPETA_TRABAJO & ArchivosOriginales(i) As _
             CARPETA_TRABAJO & ArchivosTemporales(i)
        Fase1Completada = i
    Next i

    ' Fase 2: los temporales reciben sus nombres definitivos.
    For i = LBound(ArchivosOriginales) To UBound(ArchivosOriginales)
        Name CARPETA_TRABAJO & ArchivosTemporales(i) As _
             CARPETA_TRABAJO & ArchivosNuevos(i)
        Fase2Completada = i
    Next i

    On Error GoTo 0

    MsgBox "Proceso terminado correctamente.", vbInformation, "Renombrador"
    InvalidarVistaPrevia
    Exit Sub

ControlError:
    ErrorNumero = Err.Number
    ErrorDescripcion = Err.Description
    On Error Resume Next

    ' Los destinos ya creados vuelven primero a sus temporales.
    For i = Fase2Completada To 1 Step -1
        If ArchivoExiste(CARPETA_TRABAJO & ArchivosNuevos(i)) Then
            Name CARPETA_TRABAJO & ArchivosNuevos(i) As _
                 CARPETA_TRABAJO & ArchivosTemporales(i)
        End If
    Next i

    ' Todos los temporales vuelven después a sus nombres originales.
    For i = Fase1Completada To 1 Step -1
        If ArchivoExiste(CARPETA_TRABAJO & ArchivosTemporales(i)) Then
            Name CARPETA_TRABAJO & ArchivosTemporales(i) As _
                 CARPETA_TRABAJO & ArchivosOriginales(i)
        End If
    Next i

    On Error GoTo 0
    InvalidarVistaPrevia

    MsgBox "No se pudo completar el proceso." & vbCrLf & _
           "Se intentaron restaurar los nombres originales." & vbCrLf & vbCrLf & _
           "Error " & CStr(ErrorNumero) & ": " & ErrorDescripcion, _
           vbCritical, "Error de renombrado"
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub


