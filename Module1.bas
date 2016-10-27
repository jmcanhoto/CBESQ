Attribute VB_Name = "Module1"
Option Explicit

'Definições para as variaveis ao longo da aplicação
DefBool F, T
DefByte B
DefInt I
DefLng L
DefSng S
DefDbl D
DefStr C
DefVar V

'Definição para a configuração da impressora
Type POINTAPI
        X As Long
        Y As Long
End Type

'Definição do tipo de variavel para o Utilizador
Type Utilizador
    Nome As String
    Nivel As Byte
    Perfil As String
    Password As String
End Type

'Definição do tipo de variavel para a Empresa
Type Empresa
    Codigo As String
    Nome As String
    Linha1 As String
    Linha2 As String
    Linha3 As String
    Linha4 As String
    Linha5 As String
    Linha6 As String
    Linha7 As String
    Linha8 As String
End Type

'Definição do tipo de variavela para os Cabeçalhos dos Mapas
Type Titulo
    Titulo_1 As String
    Titulo_2 As String
    Titulo_3 As String
    Titulo_4 As String
    Titulo_5 As String
End Type

'Declaração das variaveis gobais da aplicação
Global fFrmMDIPrincipal As MDIForm

Global fFichaSocio As Form
Global fPagaQuotaSocio As Form

Global fFichaInscricao As Form
Global fFichaInscricaoCAIF As Form

Global fFichaUtente As Form
Global fFichaUtenteCAIF As Form

Global fFichaFuncionario As Form

Global fFichaProlHora As Form

Global fFichaAbsent As Form

Global cApl_Path                'Caminho da aplicção
Global cBD_Path                 'Caminho da BD e do Sytem
Global dCorAmarelo              'Cor amarela para os objectos
Global dCorSelecionado          'Cor selecionada quando os objectos recebem o focus
Global dCorNormal               'Cor normal do objecto
Global gwsInicial As Workspace  'Workspace inicial
Global gUtilizador As Utilizador
Global gEmpresa As Empresa
Global Mapa As Titulo
Global cNomeBD                  'Nome da BD
Global cNomeBDTemp              'Nome da BD Temporaria
Global cBDComNomeUtilizador     'Nome da BD Temporaria para o Utilizador
Global cFormatoDefault          'Formato Default para os valores monetarios
Global bNCDecimaisDefault       'Numero de casas decimais por Default
Global cSeparadorDecimal        'Separador Decimal do Windows
Global cSeparadorMilhares       'Separador do milhares do Windows

Global cBotaoOrigem             ' para saber de onde é chamada a janela

' Funções da API do Windows
Declare Function Escape Lib "gdi32" (ByVal hdc As Long, ByVal nEscape As Long, ByVal nCount As Long, ByVal lpInData As String, lpOutData As Any) As Long
Public Sub CarregaDadosUtilizador(ByRef mBD As Database, ByVal cCodigo As String)
    Static mrecTabUtilizadores As Recordset
    
    On Error GoTo TrataErro
    Set mrecTabUtilizadores = mBD.OpenRecordset("SELECT UTILIZADOR,NOME,NIVEL,PERFIL FROM TABUTILIZADORES", dbOpenSnapshot, dbReadOnly)
    With mrecTabUtilizadores
        If Not (.EOF And .BOF) Then
            .FindFirst "UTILIZADOR='" & cCodigo & "'"
            If Not .NoMatch Then
                gUtilizador.Nivel = vFiltraCamposNulos(.Fields("NIVEL"))
                gUtilizador.Perfil = vFiltraCamposNulos(.Fields("PERFIL"))
            Else
                gUtilizador.Nivel = 0
                gUtilizador.Perfil = ""
            End If
        End If
        .Close
    End With
    
    Set mrecTabUtilizadores = Nothing
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Principal-Carrega Dados Utilizador", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub


Public Sub CarregaTitulosDoINI()
    'Vai Procurar o Cabeçalho para o Mapa
    Mapa.Titulo_1 = GetINIString("Titulo_1", " ", "CAB-MAPAS")
    Mapa.Titulo_2 = GetINIString("Titulo_2", " ", "CAB-MAPAS")
    Mapa.Titulo_3 = GetINIString("Titulo_3", " ", "CAB-MAPAS")
    Mapa.Titulo_4 = GetINIString("Titulo_4", " ", "CAB-MAPAS")
    Mapa.Titulo_5 = GetINIString("Titulo_5", " ", "CAB-MAPAS")
End Sub

'Este procedimento faz a configuração das DataCombos
'Public Sub ConfiguraDataCombo(ByRef DataCombo As GTMaskDate, Optional AnoInicial As Variant, Optional AnoFinal As Variant)
'    Static cDataPadrao

'On Error Resume Next
'
'    If IsMissing(AnoInicial) Then AnoInicial = CInt(Year(Date)) - 10
'    If IsMissing(AnoFinal) Then AnoFinal = CInt(Year(Date))
'
'    If AnoFinal < CInt(Year(Date)) Then
'        cDataPadrao = Format$(CStr(AnoFinal) & "/12/31", "yyyy/mm/dd")
'    Else
'        cDataPadrao = Format$(CStr(DateSerial(AnoFinal, Month(Date), Day(Date))), "yyyy/mm/dd")
'    End If
'
'    With DataCombo
'        .StyleSets.Add "FimDeSemana"
'        .StyleSets("FimDeSemana").BackColor = vbInfoBackground
'        .StyleSets("FimDeSemana").ForeColor = vbBlack
'        .StyleSets("FimDeSemana").Font.Name = "MS Sans Serif"
'        .StyleSets("FimDeSemana").Font.Size = 10
'        .StyleSets("FimDeSemana").Font.Bold = False
'
'        .AllowEdit = True
'        .AllowNullDate = True
'        .AutoRestore = True
'        .AutoSelect = True
'        .AutoValidate = True
'        .BackColor = dCorNormal
'        .BackColorSelected = vbActiveTitleBar
'        .BeepOnError = True
'        .ClipMode = 0
'        .DateSeparator = "/"
'        .DayofWeek(1).StyleSet = "FimDeSemana"  'Domingo
'        .DayofWeek(7).StyleSet = "FimDeSemana"  'Sabado
''        .DefaultDate = cDataPadrao
'        .DropDownFontName = "MS Sans Serif"
'        .DropDownFontSize = 10
'        .EditMode = 1
'        .Font.Name = "MS Sans Serif"
'        .Font.Size = 10
''        .MinDate = DateSerial(AnoInicial, 1, 1)
''        .MaxDate = DateSerial(AnoFinal, 12, 31)
'        .MinDate = "01-01-1900"
'        .MaxDate = "31-12-2100"
'        .Mask = 3
'        .ShowCentury = True
'        .StartOfWeek = 2
'    End With
'End Sub

'Este procedimento vai alterar o menu conforme se vão abrindo as janelas
Public Sub AlteraWindowList(sNewDocID As String)

    With fFrmMDIPrincipal.atbOpcoes
        .Tools("mnuJan").Menu.Tools.Add sNewDocID, , .Tools("mnuJan").Menu.Tools.Count + 1
        .Tools(sNewDocID).Name = sNewDocID
        .Tools(sNewDocID).Type = ssTypeStateButton
        .Tools(sNewDocID).Group = "WindowList"
        .Tools(sNewDocID).GroupAllowAllUp = False
        .Tools(sNewDocID).State = ssChecked
        .Tools(sNewDocID).PictureDown = LoadResPicture(1001, 0)
    End With

End Sub


Sub Main()
    Dim ffrmSenha As Form
    'Verifica se a aplicação já esta a ser executada
    If App.PrevInstance = True Then Call MsgBox("Este Programa já esta activo!", vbExclamation): End
    
    'Passa para a variavel global o path da aplicação
    cApl_Path = App.Path
    'Passa para a variavel global o path da Base de dados
    cBD_Path = GetINIString("Local", cApl_Path, "BD")
    'Passa para a variavel o nome da bd
    cNomeBD = "\CBESQ2000.MDB"
    cNomeBDTemp = "\TEMP2000.MDB"
    'Define as cores do sistema
    dCorAmarelo = RGB(255, 255, 200)
    dCorSelecionado = dCorAmarelo
    dCorNormal = &H80000005
    'Configura as variaveis numericas da aplicação
    cSeparadorDecimal = Mid$(FormatNumber(1000, 2, vbFalse, vbFalse, vbTrue), 6, 1)
    cSeparadorMilhares = Mid$(FormatNumber(1000, 2, vbFalse, vbFalse, vbTrue), 2, 1)
    cFormatoDefault = "###,###,##0.00"
    bNCDecimaisDefault = 2
    DBEngine.SystemDB = cBD_Path & "\System.mdw"
    
    Set ffrmSenha = New frmSenha
    Load ffrmSenha
    
End Sub
'Este procedimento vai carregar para uma
'variavel global os dados da empresa
Public Sub CarregaDadosEmpresa(ByRef mBD As Database)
    Dim mrecTabEmpresas As Recordset
    
    On Error GoTo TrataErro
    Set mrecTabEmpresas = mBD.OpenRecordset("SELECT CODIGO,DESIGNACAO,TITULO_1,TITULO_2,TITULO_3,TITULO_4,TITULO_5,TITULO_6,TITULO_7,TITULO_8 FROM TABEMPRESAS", dbOpenSnapshot, dbReadOnly)
    With mrecTabEmpresas
        If Not (.EOF And .BOF) Then
            .MoveFirst
            gEmpresa.Codigo = (.Fields("CODIGO"))
            gEmpresa.Nome = vFiltraCamposNulos(.Fields("DESIGNACAO"))
            gEmpresa.Linha1 = vFiltraCamposNulos(.Fields("TITULO_1"))
            gEmpresa.Linha2 = vFiltraCamposNulos(.Fields("TITULO_2"))
            gEmpresa.Linha3 = vFiltraCamposNulos(.Fields("TITULO_3"))
            gEmpresa.Linha4 = vFiltraCamposNulos(.Fields("TITULO_4"))
            gEmpresa.Linha5 = vFiltraCamposNulos(.Fields("TITULO_5"))
            gEmpresa.Linha6 = vFiltraCamposNulos(.Fields("TITULO_6"))
            gEmpresa.Linha7 = vFiltraCamposNulos(.Fields("TITULO_7"))
            gEmpresa.Linha8 = vFiltraCamposNulos(.Fields("TITULO_8"))
        Else
            gEmpresa.Codigo = "0000"
            gEmpresa.Nome = "Joca Software"
            gEmpresa.Linha1 = "Joca"
            gEmpresa.Linha2 = "Joca Software"
            gEmpresa.Linha3 = vbNullString
            gEmpresa.Linha4 = vbNullString
            gEmpresa.Linha5 = vbNullString
            gEmpresa.Linha6 = vbNullString
            gEmpresa.Linha7 = vbNullString
            gEmpresa.Linha8 = vbNullString
        End If
        .Close
    End With
    
    Set mrecTabEmpresas = Nothing
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Principal-Carrega Dados Empresa", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:

End Sub


'Este procedimento muda a cor do fundo control
Public Sub CorDeFundo(ByRef Caixa As Control, ByVal Flag As Boolean)
    If Flag Then
        Caixa.BackColor = dCorSelecionado
    Else
        Caixa.BackColor = dCorNormal
    End If
End Sub

