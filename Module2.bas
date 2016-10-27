Attribute VB_Name = "Module2"

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hnameMappings As Long
  lpszProgressTitle As Long
End Type

Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40

Public Const GWL_STYLE = (-16)
Public Const ES_NUMBER = &H2000&
Public Const ES_UPPERCASE = &H8&
Public Const ES_LOWERCASE = &H10&

Public Const WM_USER = &H400

Public Const EM_CANUNDO = &HC6
Public Const EM_EMPTYUNDOBUFFER = &HCD
Public Const EM_GETMODIFY = &HB8
Public Const EM_LIMITTEXT = &HC5
Public Const EM_SETMODIFY = &HB9
Public Const EM_SETREADONLY = (WM_USER + 31)




' Windows API calls
Declare Function SetWindowPos Lib "user32" (ByVal h&, ByVal hb&, ByVal X&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal f&) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Declare Function SHFileOperation Lib "shell32.dll" Alias " SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Function cAlteraVparaP(ByVal cTexto)
    Dim iPosicao
    
    iPosicao = InStr(1, cTexto, Chr(44))
    If iPosicao <> 0 Then
        cAlteraVparaP = Mid$(cTexto, 1, (iPosicao - 1)) & "." & _
                    Mid$(cTexto, (iPosicao + 1), (Len(cTexto) - iPosicao))
    Else
        cAlteraVparaP = cTexto
    End If
End Function



'===============================================================
'Esta Função vai buscar o valor de um item de determinada secção
'ao Ini File
'PARÂMETROS:
'Section -- Secção do Ini File
'Item    -- Item do Ini File
'Default -- Valor a assumir por omissão
'===============================================================
Public Function GetINIString(ByVal Item As String, ByVal Default As String, ByVal Section As String) As String
    Dim strRetorno As String
    Dim lngTamanho As Long
    strRetorno$ = String(255, 32)
    lngTamanho& = GetPrivateProfileString(Section, Item, Default, strRetorno$, Len(strRetorno$), cApl_Path & "\" & App.EXEName & ".INI")
    GetINIString$ = Left$(strRetorno$, lngTamanho&)
End Function

'Esta função escreve o numero introduzido para extenso
Public Function NumeroParaExtenso(ByVal Valor As Double) As String
    Static unidades(0 To 9) As String
    Static dezena(0 To 9) As String
    Static dezenas(0 To 9) As String
    Static centenas(0 To 9) As String
    Static milhares(0 To 7) As String
    Dim i As Integer, nPosition As Integer
    Dim nDigit As Integer, bAllZeros As Integer
    Dim strResult As String, strTemp As String
    Dim tmpBuff As String

    unidades(0) = "zero"
    unidades(1) = "um"
    unidades(2) = "dois"
    unidades(3) = "três"
    unidades(4) = "quatro"
    unidades(5) = "cinco"
    unidades(6) = "seis"
    unidades(7) = "sete"
    unidades(8) = "oito"
    unidades(9) = "nove"

    dezena(0) = "dez"
    dezena(1) = "onze"
    dezena(2) = "doze"
    dezena(3) = "treze"
    dezena(4) = "quatorze"
    dezena(5) = "quinze"
    dezena(6) = "dezasseis"
    dezena(7) = "dezassete"
    dezena(8) = "dezoito"
    dezena(9) = "dezanove"

    dezenas(0) = ""
    dezenas(1) = "dez"
    dezenas(2) = "vinte"
    dezenas(3) = "trinta"
    dezenas(4) = "quarenta"
    dezenas(5) = "cinquenta"
    dezenas(6) = "sessenta"
    dezenas(7) = "setenta"
    dezenas(8) = "oitenta"
    dezenas(9) = "noventa"

    centenas(0) = "cem"
    centenas(1) = "cento"
    centenas(2) = "duzentos"
    centenas(3) = "trezentos"
    centenas(4) = "quatrocentos"
    centenas(5) = "quinhentos"
    centenas(6) = "seiscentos"
    centenas(7) = "setecentos"
    centenas(8) = "oitocentos"
    centenas(9) = "novecentos"
    
    milhares(0) = ""
    milhares(1) = "mil"
    milhares(2) = "milhões"
    milhares(3) = "biliões"
    milhares(4) = "triliões"
    milhares(5) = "milhão"
    milhares(6) = "bilião"
    milhares(7) = "trilião"
    
    'Controla erros
    On Error GoTo NumeroParaExtensoError
        
    strTemp = CStr(Int(Valor))
    strResult = ""
    
    For i = Len(strTemp) To 1 Step -1
        'Retira o valor do digito
        nDigit = Val(Mid$(strTemp, i, 1))
        'Retira valor da coluna
        nPosition = (Len(strTemp) - i) + 1
        'a acção depende das colunas das 1's, 10's or 100's
        Select Case (nPosition Mod 3)
            Case 1  'unidades
                bAllZeros = False
                If i = 1 Then   'Se for o ultimo numero da esquerda
                    If nPosition = 4 And nDigit = 1 Then
                        tmpBuff = ""
                    Else
                        tmpBuff = unidades(nDigit) & " "
                    End If
                ElseIf Mid$(strTemp, i - 1, 1) = "1" Then
                    i = i - 1   'Skip dezenas posição
                    nPosition = nPosition + 1
                    If i = 1 Then
                        tmpBuff = dezena(nDigit) & " "
                    Else
                        tmpBuff = "e " & dezena(nDigit) & " "
                    End If
                    
                ElseIf nDigit > 0 Then
                    tmpBuff = "e " & unidades(nDigit) & " "
                Else
                    'Se as dezenas e as centenas forem tambem zero
                    'não nostra 'milhares'
                    bAllZeros = True
                    If i > 1 Then
                        If Mid$(strTemp, i - 1, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    If i > 2 Then
                        If Mid$(strTemp, i - 2, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    tmpBuff = ""
                End If
                If bAllZeros = False And nPosition > 1 Then
                    If nDigit = 1 Then
                        If (((Len(strTemp) - i) + 1) Mod 3) = 1 And nPosition <> 4 Then
                            If i = 1 Then
                                tmpBuff = tmpBuff & milhares(Int(nPosition / 3) + 3) & " "
                            Else
                                If Val(Mid$(Right$(strTemp, nPosition + 2), 1, 3)) > 1 Then
                                    tmpBuff = tmpBuff & milhares(Int(nPosition / 3)) & " "
                                Else
                                    tmpBuff = tmpBuff & milhares(Int(nPosition / 3) + 3) & " "
                                End If
                            End If
                        Else
                            tmpBuff = tmpBuff & milhares(Int(nPosition / 3)) & " "
                        End If
                    Else
                        tmpBuff = tmpBuff & milhares(Int(nPosition / 3)) & " "
                    End If
                End If
                strResult = tmpBuff & strResult
            Case 2  'dezenas
                If nDigit > 0 Then
                    If i = 1 Then   'Se for o ultimo numero da esquerda
                        strResult = dezenas(nDigit) & " " & strResult
                    Else
                        strResult = "e " & dezenas(nDigit) & " " & strResult
                    End If
                End If
            Case 0  'centenas
                If nDigit > 0 Then
                    If nDigit = 1 Then
                        'verifica se os 2 digitos anteriores são zeros
                        If Mid$(strTemp, i + 1, 1) <> "0" Or Mid$(strTemp, i + 2, 1) <> "0" Then
                            strResult = centenas(nDigit) & " " & strResult
                        Else
                            strResult = centenas(0) & " " & strResult
                        End If
                    Else
                        strResult = centenas(nDigit) & " " & strResult
                    End If
                    If i <> 1 And Mid$(strTemp, i + 1, 1) = "0" And Mid$(strTemp, i + 2, 1) = "0" Then
                        strResult = "e " & strResult
                    End If
                End If
        End Select
    Next i
    'Converte a primeira letra para maiuscula
    If Len(strResult) > 0 Then
        strResult = UCase$(Left$(strResult, 1)) & Mid$(strResult, 2)
    End If

EndNumeroParaExtenso:
    'Return result
    NumeroParaExtenso = strResult
    Exit Function

NumeroParaExtensoError:
    strResult = "#Erro#"
    Resume EndNumeroParaExtenso
End Function
'Esta função escreve o numero introduzido para extenso
Public Function NumeroParaExtenso_Euro(ByVal Valor As Double) As String
    Static unidades(0 To 9) As String
    Static dezena(0 To 9) As String
    Static dezenas(0 To 9) As String
    Static centenas(0 To 9) As String
    Static milhares(0 To 7) As String
    Dim i As Integer, nPosition As Integer
    Dim nDigit As Integer, bAllZeros As Integer
    Dim strResult As String, strResult2 As String, strTemp1 As String, strTemp2 As String
    Dim tmpBuff As String

    unidades(0) = "zero"
    unidades(1) = "um"
    unidades(2) = "dois"
    unidades(3) = "três"
    unidades(4) = "quatro"
    unidades(5) = "cinco"
    unidades(6) = "seis"
    unidades(7) = "sete"
    unidades(8) = "oito"
    unidades(9) = "nove"

    dezena(0) = "dez"
    dezena(1) = "onze"
    dezena(2) = "doze"
    dezena(3) = "treze"
    dezena(4) = "quatorze"
    dezena(5) = "quinze"
    dezena(6) = "dezasseis"
    dezena(7) = "dezassete"
    dezena(8) = "dezoito"
    dezena(9) = "dezanove"

    dezenas(0) = ""
    dezenas(1) = "dez"
    dezenas(2) = "vinte"
    dezenas(3) = "trinta"
    dezenas(4) = "quarenta"
    dezenas(5) = "cinquenta"
    dezenas(6) = "sessenta"
    dezenas(7) = "setenta"
    dezenas(8) = "oitenta"
    dezenas(9) = "noventa"

    centenas(0) = "cem"
    centenas(1) = "cento"
    centenas(2) = "duzentos"
    centenas(3) = "trezentos"
    centenas(4) = "quatrocentos"
    centenas(5) = "quinhentos"
    centenas(6) = "seiscentos"
    centenas(7) = "setecentos"
    centenas(8) = "oitocentos"
    centenas(9) = "novecentos"
    
    milhares(0) = ""
    milhares(1) = "mil"
    milhares(2) = "milhões"
    milhares(3) = "biliões"
    milhares(4) = "triliões"
    milhares(5) = "milhão"
    milhares(6) = "bilião"
    milhares(7) = "trilião"
    
    'Controla erros
    On Error GoTo NumeroParaExtensoError
        
    strTemp1 = CStr(Int(Valor))
    strResult = ""
    
    For i = Len(strTemp1) To 1 Step -1
        'Retira o valor do digito
        nDigit = Val(Mid$(strTemp1, i, 1))
        'Retira valor da coluna
        nPosition = (Len(strTemp1) - i) + 1
        'a acção depende das colunas das 1's, 10's or 100's
        Select Case (nPosition Mod 3)
            Case 1  'unidades
                bAllZeros = False
                If i = 1 Then   'Se for o ultimo numero da esquerda
                    If nPosition = 4 And nDigit = 1 Then
                        tmpBuff = ""
                    Else
                        tmpBuff = unidades(nDigit) & " "
                    End If
                ElseIf Mid$(strTemp1, i - 1, 1) = "1" Then
                    i = i - 1   'Skip dezenas posição
                    nPosition = nPosition + 1
                    If i = 1 Then
                        tmpBuff = dezena(nDigit) & " "
                    Else
                        tmpBuff = "e " & dezena(nDigit) & " "
                    End If
                    
                ElseIf nDigit > 0 Then
                    tmpBuff = "e " & unidades(nDigit) & " "
                Else
                    'Se as dezenas e as centenas forem tambem zero
                    'não nostra 'milhares'
                    bAllZeros = True
                    If i > 1 Then
                        If Mid$(strTemp1, i - 1, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    If i > 2 Then
                        If Mid$(strTemp1, i - 2, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    tmpBuff = ""
                End If
                If bAllZeros = False And nPosition > 1 Then
                    If nDigit = 1 Then
                        If (((Len(strTemp1) - i) + 1) Mod 3) = 1 And nPosition <> 4 Then
                            If i = 1 Then
                                tmpBuff = tmpBuff & milhares(Int(nPosition / 3) + 3) & " "
                            Else
                                If Val(Mid$(Right$(strTemp1, nPosition + 2), 1, 3)) > 1 Then
                                    tmpBuff = tmpBuff & milhares(Int(nPosition / 3)) & " "
                                Else
                                    tmpBuff = tmpBuff & milhares(Int(nPosition / 3) + 3) & " "
                                End If
                            End If
                        Else
                            tmpBuff = tmpBuff & milhares(Int(nPosition / 3)) & " "
                        End If
                    Else
                        tmpBuff = tmpBuff & milhares(Int(nPosition / 3)) & " "
                    End If
                End If
                strResult = tmpBuff & strResult
            Case 2  'dezenas
                If nDigit > 0 Then
                    If i = 1 Then   'Se for o ultimo numero da esquerda
                        strResult = dezenas(nDigit) & " " & strResult
                    Else
                        strResult = "e " & dezenas(nDigit) & " " & strResult
                    End If
                End If
            Case 0  'centenas
                If nDigit > 0 Then
                    If nDigit = 1 Then
                        'verifica se os 2 digitos anteriores são zeros
                        If Mid$(strTemp1, i + 1, 1) <> "0" Or Mid$(strTemp1, i + 2, 1) <> "0" Then
                            strResult = centenas(nDigit) & " " & strResult
                        Else
                            strResult = centenas(0) & " " & strResult
                        End If
                    Else
                        strResult = centenas(nDigit) & " " & strResult
                    End If
                    If i <> 1 And Mid$(strTemp1, i + 1, 1) = "0" And Mid$(strTemp1, i + 2, 1) = "0" Then
                        strResult = "e " & strResult
                    End If
                End If
        End Select
    Next i
    'Converte a primeira letra para maiuscula
    If Len(strResult) > 0 Then
        strResult = UCase$(Left$(strResult, 1)) & Mid$(strResult, 2)
    End If

' Parte Decimal
    If InStr(1, CStr(Valor), ",") = 0 Then
        strTemp2 = ""
    Else
        strTemp2 = Mid$(CStr(Format(Valor, "0.00")), InStr(1, CStr(Valor), ",") + 1, 2)
    End If
    strTemp2 = CStr(Val(strTemp2))
    
    For i = Len(strTemp2) To 1 Step -1
        'Retira o valor do digito
        nDigit = Val(Mid$(strTemp2, i, 1))
        'Retira valor da coluna
        nPosition = (Len(strTemp2) - i) + 1
        'a acção depende das colunas das 1's, 10's or 100's
        Select Case (nPosition Mod 3)
            Case 1  'unidades
                bAllZeros = False
                If i = 1 Then   'Se for o ultimo numero da esquerda
                    If nPosition = 4 And nDigit = 1 Then
                        tmpBuff = ""
                    Else
                        tmpBuff = unidades(nDigit) & " "
                    End If
                ElseIf Mid$(strTemp2, i - 1, 1) = "1" Then
                    i = i - 1   'Skip dezenas posição
                    nPosition = nPosition + 1
                    If i = 1 Then
                        tmpBuff = dezena(nDigit) & " "
                    Else
                        tmpBuff = "e " & dezena(nDigit) & " "
                    End If
                    
                ElseIf nDigit > 0 Then
                    tmpBuff = "e " & unidades(nDigit) & " "
                Else
                    'Se as dezenas e as centenas forem tambem zero
                    'não nostra 'milhares'
                    bAllZeros = True
                    If i > 1 Then
                        If Mid$(strTemp2, i - 1, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    If i > 2 Then
                        If Mid$(strTemp2, i - 2, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    tmpBuff = ""
                End If
                If bAllZeros = False And nPosition > 1 Then
                    If nDigit = 1 Then
                        If (((Len(strTemp2) - i) + 1) Mod 3) = 1 And nPosition <> 4 Then
                            If i = 1 Then
                                tmpBuff = tmpBuff & milhares(Int(nPosition / 3) + 3) & " "
                            Else
                                If Val(Mid$(Right$(strTemp2, nPosition + 2), 1, 3)) > 1 Then
                                    tmpBuff = tmpBuff & milhares(Int(nPosition / 3)) & " "
                                Else
                                    tmpBuff = tmpBuff & milhares(Int(nPosition / 3) + 3) & " "
                                End If
                            End If
                        Else
                            tmpBuff = tmpBuff & milhares(Int(nPosition / 3)) & " "
                        End If
                    Else
                        tmpBuff = tmpBuff & milhares(Int(nPosition / 3)) & " "
                    End If
                End If
                strResult2 = tmpBuff & strResult2
            Case 2  'dezenas
                If nDigit > 0 Then
                    If i = 1 Then   'Se for o ultimo numero da esquerda
                        strResult2 = dezenas(nDigit) & " " & strResult2
                    Else
                        strResult2 = "e " & dezenas(nDigit) & " " & strResult2
                    End If
                End If
            Case 0  'centenas
                If nDigit > 0 Then
                    If nDigit = 1 Then
                        'verifica se os 2 digitos anteriores são zeros
                        If Mid$(strTemp2, i + 1, 1) <> "0" Or Mid$(strTemp2, i + 2, 1) <> "0" Then
                            strResult2 = centenas(nDigit) & " " & strResult2
                        Else
                            strResult2 = centenas(0) & " " & strResult2
                        End If
                    Else
                        strResult2 = centenas(nDigit) & " " & strResult2
                    End If
                    If i <> 1 And Mid$(strTemp2, i + 1, 1) = "0" And Mid$(strTemp2, i + 2, 1) = "0" Then
                        strResult2 = "e " & strResult2
                    End If
                End If
        End Select
    Next i

    'Converte a primeira letra para maiuscula
    If Len(strResult2) > 0 Then
        strResult2 = UCase$(Left$(strResult2, 1)) & Mid$(strResult2, 2)
    End If

EndNumeroParaExtenso:
    'Return result
    NumeroParaExtenso_Euro = IIf(strTemp1 > 0, strResult & IIf(strTemp1 > 1, "Euros", "Euro"), vbNullString) & _
                        IIf(strTemp1 > 0, IIf(strTemp2 > 0, " e ", vbNullString), vbNullString) & _
                        IIf(strTemp2 > 0, strResult2 & IIf(strTemp2 > 1, "Cêntimos", "Cêntimo"), vbNullString)
    Exit Function

NumeroParaExtensoError:
    strResult = "#Erro#"
    Resume EndNumeroParaExtenso
End Function


'===============================================================
'Esta Função Converte o a primeira letra de uma string para Maiuscula
'PARÂMETRO:
'Mystring -- string
'===============================================================''
Public Function PrimeiraLetraMaiuscula(ByVal MyString As String) As String
    Dim PosSpc%
    Mid$(MyString, 1, 1) = UCase$(Mid$(MyString, 1, 1))
    PosSpc% = InStr(MyString, " ")
    While PosSpc% <> 0
        Mid$(MyString, PosSpc% + 1, 1) = UCase$(Mid$(MyString, PosSpc% + 1, 1))
        PosSpc% = InStr(PosSpc% + 1, MyString, " ")
    Wend
    PrimeiraLetraMaiuscula = MyString
End Function

'Esta função filtra os campos da tabela para os nulos sejam controlados
Public Function vFiltraCamposNulos(Campo As Field)
  Dim vRetorno
  If IsNull(Campo) Then
    Select Case Campo.Type
    Case 1                  'Campos logicos
      vRetorno = False
    Case 2, 3, 4, 5, 6, 7   'Campos Numericos
      vRetorno = 0
    Case 10, 12             'Campos de Texto e Memo
      vRetorno = vbNullString
    End Select
  Else
    vRetorno = Campo.Value
  End If
    
  vFiltraCamposNulos = vRetorno
End Function

'Este procedimento mostra uma caixa com o Erro
'e grava num ficheiro a ocurrencia
Public Sub ErrosGerais(ByVal cModName As String, lErrorNumber As Long, Optional cDescription As Variant)
    Dim cMsg, cFicheiroErro, cDatetime
    
    On Error GoTo ErrosGerais_Erro
    
    'Defime o caminho para o ficheiro de registo do Erro
    cFicheiroErro = cApl_Path & "\ERROS.TXT"
    
    Select Case lErrorNumber
        Case 3022
            cMsg = "Os Dados deste registo não foram gravados." & vbCrLf & _
                "Porque já existe um registo com este indice."
        Case 3186
            cMsg = "Não foi possivel gravar os dados." & vbCrLf & _
                "Porque a Tabela encontra-se Bloqueada por outro utilizador." & vbCrLf & _
                "Espere alguns segundos e tente gravar novamente"
        Case 3260
            cMsg = "Não foi possivel gravar os dados." & vbCrLf & _
                "Porque a Base de Dados encontra-se Bloqueada por outro utilizador." & vbCrLf & _
                "Espere alguns segundos e tente gravar novamente"
        Case Else
            'Constroi a mensagem de erro
            cMsg = "Ocorreu um erro no modulo " & cModName
            cMsg = cMsg & " Nº do Erro (" & lErrorNumber & ") : " & Error(lErrorNumber) & "."
            If Not IsMissing(cDescription) And cDescription <> Error(lErrorNumber) Then
                cMsg = cMsg & vbCrLf & " {" & cDescription & "}"
            End If
    End Select
    
    MsgBox cMsg, vbCritical, "Erro na Aplicação"
    
    cDatetime = Format$(Date, "yyyy/mm/dd") & " " & Format$(Time, "hh:mm")
    'Regista o erro no ficheiro ERROR.LOG
    Open cFicheiroErro For Append As #1
    If IsMissing(cDescription) Then
        Write #1, cDatetime, cModName, lErrorNumber, Error(lErrorNumber)
    Else
        Write #1, cDatetime, cModName, lErrorNumber, Error(lErrorNumber), cDescription
    End If
    Close #1
    
    'Repõe o cursor do mouse no normal
    Screen.MousePointer = vbDefault
    
    Exit Sub
ErrosGerais_Erro:
    Resume Next
End Sub

'Este procedimento apaga um ficheiro para o Recycle Bin
Public Sub ApagaFicheiro(ByVal strNome As String)
  Dim SHop As SHFILEOPSTRUCT
  With SHop
    .wFunc = FO_DELETE
    .pFrom = strNome
    .fFlags = FOF_ALLOWUNDO
  End With
  SHFileOperation SHop
End Sub

'Esta função retorna o nome do computador onde esta instalado o sistema
Public Function NomeDoComputador() As String
    Dim s$, cnt&, dl&, CurComputador$
  
    cnt& = 199
    s$ = String$(200, 0)
    dl = GetComputerName(s, cnt)
    If dl <> 0 Then CurComputador$ = Left$(s, cnt) Else CurComputador$ = ""
    NomeDoComputador = CurComputador$
End Function

'Esta função retorna o nome do utilizador liga ao sistema
Public Function NomeDoUtilizador() As String
  Dim s$, cnt&, dl&, CurUser$
  
  cnt& = 199
  s$ = String$(200, 0)
  dl = GetUserName(s, cnt)
  If dl <> 0 Then CurUser$ = Left$(s, cnt) Else CurUser = ""
  NomeDoUtilizador = CurUser$
End Function

'Esta função retorna o numero de serie do disco ou disquete
Public Function NumeroDeSerie(strDrive As String) As Long
  Dim SeriealNum&, Res&, Temp1$, Temp2$
  
  Temp1 = String$(255, Chr$(0))
  Temp2 = String$(255, Chr$(0))
  Res = GetVolumeInformation(strDrive, Temp1, Len(Temp1), SeriealNum&, 0, 0, Temp2, Len(Temp2))
  NumeroDeSerie = SeriealNum&
End Function

'Este procedimento configura as textbox so para receber numeros
Public Sub SetNumber(NumberText As TextBox, Flag As Boolean)
  Dim CURSTYLE&, NEWSTYLE&
  CURSTYLE = GetWindowLong(NumberText.hwnd, GWL_STYLE)
  If Flag Then
    CURSTYLE = CURSTYLE Or ES_NUMBER
  Else
    CURSTYLE = CURSTYLE And (Not ES_NUMBER)
  End If
  NEWSTYLE = SetWindowLong(NumberText.hwnd, GWL_STYLE, CURSTYLE)
  NumberText.Refresh
End Sub

Public Sub Showsplash()
    Dim Success%, strTemporizador$
    
    On Error GoTo SplashLoadErr
    
    Screen.MousePointer = vbHourglass
    
    Load frmSplash
    frmSplash.Show
    
    'Inicializa o relogio
    strTemporizador$ = Time$
    
    DoEvents
    'Mete o ecran de entrada sempre em primeiro
    Success% = SetWindowPos(frmSplash.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    
    Set fFrmMDIPrincipal = New frmMDIPrincipal
    Load fFrmMDIPrincipal
    fFrmMDIPrincipal.Show
    
    'força que a entrada no programa demore pelo menos 3 segundos
    Do While DateDiff("s", strTemporizador$, Time$) < 3
    Loop
    
    Success% = SetWindowPos(frmSplash.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    Unload frmSplash
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
SplashLoadErr:
    Success% = SetWindowPos(frmSplash.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    Unload frmSplash
    Screen.MousePointer = vbDefault
    MsgBox str$(Err), vbExclamation, "Erro na aplicação"
    Exit Sub
End Sub

Public Function DoubleString(ByVal str As String, ByVal strDoubleString As String) As String
    Dim intStringLength As Integer
    Dim intDoubleStringLength As Integer
    Dim intPosition As Integer
    Dim strTemp As String
    
    intStringLength = Len(str)
    intDoubleStringLength = Len(strDoubleString)
    strTemp = str
    If intStringLength >= intDoubleStringLength And intDoubleStringLength > 0 Then
        intPosition = 1
        Do While (intPosition > 0) And (intPosition <= intStringLength)
            intPosition = InStr(intPosition, strTemp, strDoubleString)
            If intPosition > 0 Then
                strTemp = Left$(strTemp, intPosition - 1 + intDoubleStringLength) & strDoubleString & Mid$(strTemp, intPosition + intDoubleStringLength, intStringLength)
                intStringLength = Len(strTemp)
                intPosition = intPosition + (intDoubleStringLength * 2)
            End If
        Loop
    End If
    DoubleString = strTemp
End Function

Public Function SetTopWindow(hwnd As Long, bState As Boolean) As Boolean
      
  If bState = True Then 'Put the window on top
    SetTopWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  ElseIf bState = False Then ' Turn off the TopMost flag
    SetTopWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
  Else
    SetTopWindow = False
  End If
  
End Function

' Centers the Form object passed to it.
Public Sub CenterMe(frmForm As Form)
  frmForm.Left = (Screen.Width - frmForm.Width) / 2
  frmForm.Top = ((Screen.Height - frmForm.Height) / 2) - 800
End Sub


Public Sub LoadResStrings(frm As Form)
    Dim ctl As Control, sCtlType As String
    
    On Error Resume Next
    
    If Val(frm.Tag) > 0 Then frm.Icon = LoadResPicture(Val(frm.Tag), 1)

    'set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls
    For Each ctl In frm.Controls
        sCtlType = TypeName(ctl)
        If sCtlType = "CommandButton" Then
            If Val(ctl.Tag) > 0 Then ctl.Picture = LoadResPicture(Val(ctl.Tag), 1)
        End If
    Next
End Sub
'Esta função verifica se houve ou não alteração numa textbox
Public Function tHouveAlteracao(ByVal ctlTextBox As TextBox)
    tHouveAlteracao = SendMessage(ctlTextBox.hwnd, EM_GETMODIFY, 0, 0) <> 0
End Function

'Este procedimento seleciona todo o texto do control
Public Sub SelecionaTexto(txt As Control)
    If TypeName(txt) = "TextBox" Then
        With txt
            If (GetKeyState(vbKeyTab) < 0) Or (GetKeyState(vbKeyMenu) < 0) Then
                'If TextWidth(.Text) > .Width Then
                '    SendKeys "{End}", True
                '    SendKeys "+{Home}", True
                'Else
                    .SelStart = 0
                    .SelLength = Len(.Text)
                'End If
            Else
                .SelLength = 0
            End If
        End With
    End If
End Sub

Public Function iErrosDeBD()
    Static tblockout
    Select Case Err.Number
        Case 70, 3260
            Static lockcounter
            Static clastproc
            If clastproc <> Err.Source Then
                clastproc = Err.Source
                lockcounter = 0
            End If
            lockcounter = lockcounter + 1
            If lockcounter > 2 Then
                If tblockout = True Then
                    MsgBox "unabled to save due to a " & _
                    "persistent record lock" & _
                    vbCrLf & _
                    "your data my need to be reentered"
                    lockcounter = 0
                    tblockout = True
                    iErrosDeBD = 3
                Else
                    Dim lpausetime, lStart, frm As Form
                    lpausetime = 5
                    Screen.MousePointer = vbHourglass
                    Set frm = Screen.ActiveForm
                    frm.Enabled = False
                    'frmwait.Show
                    lStart = Timer + lpausetime
                    Do While Timer < lStart
                        DoEvents
                    Loop
                    Screen.MousePointer = vbDefault
                    'Unload frmwait
                    frm.Enabled = True
                    'Set frmwait = Nothing
                    tblockout = True
                    iErrosDeBD = 1
                End If
            Else
                iErrosDeBD = 1
            End If
        Case 53
        Case 3045
            MsgBox "database already in use"
            iErrosDeBD = 3
        Case 3343
            MsgBox "the database appears to be corrupted"
            iErrosDeBD = 3
        End Select
                        
End Function

'Este procedimento configura as textbox so para receber Maiusculas
Public Sub SetUpperCase(UpperCaseText As TextBox, Flag As Boolean)
    Dim CURSTYLE&, NEWSTYLE&
    CURSTYLE = GetWindowLong(UpperCaseText.hwnd, GWL_STYLE)
    If Flag Then
      CURSTYLE = CURSTYLE Or ES_UPPERCASE
    Else
      CURSTYLE = CURSTYLE And (Not ES_UPPERCASE)
    End If
    NEWSTYLE = SetWindowLong(UpperCaseText.hwnd, GWL_STYLE, CURSTYLE)
    UpperCaseText.Refresh
End Sub

'Este procedimento configura as textbox so para receber Minusculas
Public Sub SetLowerCase(LowerCaseText As TextBox, Flag As Boolean)
    Dim CURSTYLE&, NEWSTYLE&
    CURSTYLE = GetWindowLong(LowerCaseText.hwnd, GWL_STYLE)
    If Flag Then
      CURSTYLE = CURSTYLE Or ES_LOWERCASE
    Else
      CURSTYLE = CURSTYLE And (Not ES_LOWERCASE)
    End If
    NEWSTYLE = SetWindowLong(LowerCaseText.hwnd, GWL_STYLE, CURSTYLE)
    LowerCaseText.Refresh
End Sub

'Este procedimento define o numero de caracteres que se podem introduzir
Public Sub LimitaTexto(ByRef Caixa As Control, ByVal Limite As Long)
    Dim l, cTipo
    cTipo = TypeName(Caixa)
    If cTipo = "SSDBCombo" Then
        l = SendMessage(Caixa.HwndEdit, EM_LIMITTEXT, Limite, 0)
    ElseIf cTipo = "TextBox" Then
        l = SendMessage(Caixa.hwnd, EM_LIMITTEXT, Limite, 0)
    End If
End Sub

Public Function dArredonda(ByVal dValor As Double, Optional vNCDecimais As Variant)
    Dim iNCDecimais
    
    If IsMissing(vNCDecimais) Then
        iNCDecimais = 0
    Else
        iNCDecimais = CInt(vNCDecimais)
    End If
    
    dArredonda = Int((dValor * 10 ^ iNCDecimais) + 0.5) / 10 ^ iNCDecimais
        
End Function

'Esta função verifica se o dado ficheiro existe
Public Function tFicheiroExiste(ByVal cPathMaisNome As String)
    tFicheiroExiste = CBool(Len(Dir(cPathMaisNome, vbNormal)))
End Function

Public Function iOpcaoSelecionada(ByRef Opcoes As Object)
    Dim iCiclo, iRetorno
    
    On Error GoTo ErroNaOpcaoSelecionada
    'Valor padrão
    iRetorno = -1
    For iCiclo = Opcoes.LBound To Opcoes.UBound
        If Opcoes(iCiclo) Then iRetorno = iCiclo: Exit For
    Next iCiclo
    
ErroNaOpcaoSelecionada:
    iOpcaoSelecionada = iRetorno
End Function

Public Sub CompactarBD(ByVal cCaminho, ByVal cBD)
    Dim cNovaBD
    cCaminho = Trim$(cCaminho) & "\"
    cBD = Trim$(cBD)
    cNovaBD = cBD & "_Compactada"
    If MsgBox("Confirma a Compactação da Base de Dados?", vbQuestion + vbDefaultButton1 + vbYesNo, "Compactar Base Dados") = vbYes Then
        If tFicheiroExiste((cCaminho & cBD)) Then
            DBEngine.CompactDatabase (cCaminho & cBD), (cCaminho & cNovaBD), dbLangGeneral
            Name (cCaminho & cBD) As (cCaminho & Left(cBD, Len(cBD) - 4) & ".MDS")
            Name (cCaminho & cNovaBD) As (cCaminho & cBD)
            Kill (cCaminho & cNovaBD)
        End If
    End If
End Sub

Public Function cSQLString(ByVal cString)
    cSQLString = Replace(cString, Chr$(39), Chr$(34))
End Function


