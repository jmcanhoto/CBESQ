Attribute VB_Name = "Module3"
Option Explicit

Public Sub CarregacboNomeUtentes(ByRef Combo As SSDBCombo, ByVal cInst, ByVal cSala)
    Dim mBDBaseDados As Database
    Dim recTabelaUTENTES As Recordset
    Dim cLinha
    Dim cSql
    
    Combo.Redraw = False
    Combo.RemoveAll
    
    ' começa a construir o Sql
    cSql = "SELECT NOME,NUM_UTENTE FROM UTENTES WHERE ISNULL(DATA_SAIDA)"
    ' se seleccionou Instituição tem de filtrar
    If cInst <> "<Todas as Instituições>" Then
        cSql = cSql & " AND COD_INST='" & cCodificaInstituicao(cInst) & "'"
    End If
    ' se seleccionou Sala tem de filtrar
    If cSala <> "<Todas as Salas>" Then
        cSql = cSql & " AND COD_SALA='" & cCodificaSala(cCodificaInstituicao(cInst), cSala) & "'"
    End If
    ' Estabelace a ordem dos registos
    cSql = cSql & " ORDER BY NOME ASC"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTabelaUTENTES = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    
    If Not recTabelaUTENTES.EOF And Not recTabelaUTENTES.BOF Then
        While Not recTabelaUTENTES.EOF
            Combo.AddItem recTabelaUTENTES!Nome & vbTab & _
                            recTabelaUTENTES!NUM_UTENTE
            recTabelaUTENTES.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTabelaUTENTES.Close
    Set recTabelaUTENTES = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    Combo.Redraw = True
End Sub
Public Sub CarregacboNomeUtentesCAIF(ByRef Combo As SSDBCombo, ByVal cInst, ByVal cSala)
    Dim mBDBaseDados As Database
    Dim recTabelaUTENTES As Recordset
    Dim cLinha
    Dim cSql
    
    Combo.Redraw = False
    Combo.RemoveAll
    
    ' começa a construir o Sql
    cSql = "SELECT NOME,NUM_UTENTE FROM UTENTES_IDOSOS WHERE ISNULL(DATA_SAIDA)"
    ' se seleccionou Instituição tem de filtrar
    If cInst <> "<Todas as Instituições>" Then
        cSql = cSql & " AND COD_INST='" & cCodificaInstituicao(cInst) & "'"
    End If
    ' se seleccionou Sala tem de filtrar
    If cSala <> "<Todas as Salas>" Then
        cSql = cSql & " AND COD_SALA='" & cCodificaSala(cCodificaInstituicao(cInst), cSala) & "'"
    End If
    ' Estabelace a ordem dos registos
    cSql = cSql & " ORDER BY NOME ASC"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTabelaUTENTES = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    
    If Not recTabelaUTENTES.EOF And Not recTabelaUTENTES.BOF Then
        While Not recTabelaUTENTES.EOF
            Combo.AddItem recTabelaUTENTES!Nome & vbTab & _
                            recTabelaUTENTES!NUM_UTENTE
            recTabelaUTENTES.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTabelaUTENTES.Close
    Set recTabelaUTENTES = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    Combo.Redraw = True
End Sub

Public Sub CarregacboNomeUtentesTodos(ByRef Combo As SSDBCombo, ByVal cInst, ByVal cSala)
    Dim mBDBaseDados As Database
    Dim recTabelaUTENTES As Recordset
    Dim cLinha
    Dim cSql
    
    Combo.Redraw = False
    Combo.RemoveAll
    
    ' começa a construir o Sql
    cSql = "SELECT NOME,NUM_UTENTE FROM UTENTES "
    ' se seleccionou Instituição tem de filtrar
    If cInst <> "<Todas as Instituições>" Then
        cSql = cSql & " WHERE COD_INST='" & cCodificaInstituicao(cInst) & "'"
    End If
    ' se seleccionou Sala tem de filtrar
    If cSala <> "<Todas as Salas>" Then
        cSql = cSql & " AND COD_SALA='" & cCodificaSala(cCodificaInstituicao(cInst), cSala) & "'"
    End If
    ' Estabelace a ordem dos registos
    cSql = cSql & " ORDER BY NOME ASC"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTabelaUTENTES = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    
    If Not recTabelaUTENTES.EOF And Not recTabelaUTENTES.BOF Then
        While Not recTabelaUTENTES.EOF
            Combo.AddItem recTabelaUTENTES!Nome & vbTab & _
                            recTabelaUTENTES!NUM_UTENTE
            recTabelaUTENTES.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTabelaUTENTES.Close
    Set recTabelaUTENTES = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    Combo.Redraw = True
End Sub

Public Sub CarregacboNomeFunc(ByRef Combo As SSDBCombo, ByVal cInst)
    Dim mBDBaseDados As Database
    Dim recTabelaFUNC As Recordset
    Dim cLinha
    Dim cSql
    
    Combo.Redraw = False
    Combo.RemoveAll
    
    ' começa a construir o Sql
    cSql = "SELECT NOME,NUM_FUNCIONARIO FROM FUNCIONARIOS WHERE ISNULL(DATA_DEMISSAO)"
    ' se seleccionou Instituição tem de filtrar
    If cInst <> "<Todas as Instituições>" Then
        cSql = cSql & " AND COD_INST='" & cCodificaInstituicao(cInst) & "'"
    End If
    ' Estabelace a ordem dos registos
    cSql = cSql & " ORDER BY NOME ASC"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTabelaFUNC = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    
    If Not recTabelaFUNC.EOF And Not recTabelaFUNC.BOF Then
        While Not recTabelaFUNC.EOF
            Combo.AddItem recTabelaFUNC!Nome & vbTab & _
                            recTabelaFUNC!NUM_FUNCIONARIO
            recTabelaFUNC.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTabelaFUNC.Close
    Set recTabelaFUNC = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    Combo.Redraw = True
End Sub

Public Sub CarregacboNumUtentes(ByRef Combo As SSDBCombo, ByVal cInst, ByVal cSala)
    Dim mBDBaseDados As Database
    Dim recTabelaUTENTES As Recordset
    Dim cLinha
    Dim cSql
    
    Combo.Redraw = False
    Combo.RemoveAll
    
    ' começa a construir o Sql
    cSql = "SELECT NUM_UTENTE,NOME FROM UTENTES WHERE ISNULL(DATA_SAIDA)"
    ' se seleccionou Instituição tem de filtrar
    If cInst <> "<Todas as Instituições>" Then
        cSql = cSql & " AND COD_INST='" & cCodificaInstituicao(cInst) & "'"
    End If
    ' se seleccionou Sala tem de filtrar
    If cSala <> "<Todas as Salas>" Then
        cSql = cSql & " AND COD_SALA='" & cCodificaSala(cCodificaInstituicao(cInst), cSala) & "'"
    End If
    ' Estabelace a ordem dos registos
    cSql = cSql & " ORDER BY NUM_UTENTE ASC"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTabelaUTENTES = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    
    If Not recTabelaUTENTES.EOF And Not recTabelaUTENTES.BOF Then
        While Not recTabelaUTENTES.EOF
            Combo.AddItem recTabelaUTENTES!NUM_UTENTE & vbTab & _
                            recTabelaUTENTES!Nome
            recTabelaUTENTES.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTabelaUTENTES.Close
    Set recTabelaUTENTES = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    Combo.Redraw = True
End Sub
Public Sub CarregacboNumUtentesCAIF(ByRef Combo As SSDBCombo, ByVal cInst, ByVal cSala)
    Dim mBDBaseDados As Database
    Dim recTabelaUTENTES As Recordset
    Dim cLinha
    Dim cSql
    
    Combo.Redraw = False
    Combo.RemoveAll
    
    ' começa a construir o Sql
    cSql = "SELECT NUM_UTENTE,NOME FROM UTENTES_IDOSOS WHERE ISNULL(DATA_SAIDA)"
    ' se seleccionou Instituição tem de filtrar
    If cInst <> "<Todas as Instituições>" Then
        cSql = cSql & " AND COD_INST='" & cCodificaInstituicao(cInst) & "'"
    End If
    ' se seleccionou Sala tem de filtrar
    If cSala <> "<Todas as Salas>" Then
        cSql = cSql & " AND COD_SALA='" & cCodificaSala(cCodificaInstituicao(cInst), cSala) & "'"
    End If
    ' Estabelace a ordem dos registos
    cSql = cSql & " ORDER BY NUM_UTENTE ASC"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTabelaUTENTES = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    
    If Not recTabelaUTENTES.EOF And Not recTabelaUTENTES.BOF Then
        While Not recTabelaUTENTES.EOF
            Combo.AddItem recTabelaUTENTES!NUM_UTENTE & vbTab & _
                            recTabelaUTENTES!Nome
            recTabelaUTENTES.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTabelaUTENTES.Close
    Set recTabelaUTENTES = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    Combo.Redraw = True
End Sub
Public Sub CarregacboNumUtentesCAIFTodos(ByRef Combo As SSDBCombo, ByVal cInst, ByVal cSala)
    Dim mBDBaseDados As Database
    Dim recTabelaUTENTES As Recordset
    Dim cLinha
    Dim cSql
    
    Combo.Redraw = False
    Combo.RemoveAll
    
    ' começa a construir o Sql
    cSql = "SELECT NUM_UTENTE,NOME FROM UTENTES_IDOSOS "
    ' se seleccionou Instituição tem de filtrar
    If cInst <> "<Todas as Instituições>" Then
        cSql = cSql & " WHERE COD_INST='" & cCodificaInstituicao(cInst) & "'"
    End If
    ' se seleccionou Sala tem de filtrar
    If cSala <> "<Todas as Salas>" Then
        cSql = cSql & " AND COD_SALA='" & cCodificaSala(cCodificaInstituicao(cInst), cSala) & "'"
    End If
    ' Estabelace a ordem dos registos
    cSql = cSql & " ORDER BY NUM_UTENTE ASC"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTabelaUTENTES = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    
    If Not recTabelaUTENTES.EOF And Not recTabelaUTENTES.BOF Then
        While Not recTabelaUTENTES.EOF
            Combo.AddItem recTabelaUTENTES!NUM_UTENTE & vbTab & _
                            recTabelaUTENTES!Nome
            recTabelaUTENTES.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTabelaUTENTES.Close
    Set recTabelaUTENTES = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    Combo.Redraw = True
End Sub

Public Sub CarregacboNumUtentesTodos(ByRef Combo As SSDBCombo, ByVal cInst, ByVal cSala)
    Dim mBDBaseDados As Database
    Dim recTabelaUTENTES As Recordset
    Dim cLinha
    Dim cSql
    
    Combo.Redraw = False
    Combo.RemoveAll
    
    ' começa a construir o Sql
    cSql = "SELECT NUM_UTENTE,NOME FROM UTENTES"
    ' se seleccionou Instituição tem de filtrar
    If cInst <> "<Todas as Instituições>" Then
        cSql = cSql & " WHERE COD_INST='" & cCodificaInstituicao(cInst) & "'"
    End If
    ' se seleccionou Sala tem de filtrar
    If cSala <> "<Todas as Salas>" Then
        cSql = cSql & " AND COD_SALA='" & cCodificaSala(cCodificaInstituicao(cInst), cSala) & "'"
    End If
    ' Estabelace a ordem dos registos
    cSql = cSql & " ORDER BY NUM_UTENTE ASC"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTabelaUTENTES = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    
    If Not recTabelaUTENTES.EOF And Not recTabelaUTENTES.BOF Then
        While Not recTabelaUTENTES.EOF
            Combo.AddItem recTabelaUTENTES!NUM_UTENTE & vbTab & _
                            recTabelaUTENTES!Nome
            recTabelaUTENTES.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTabelaUTENTES.Close
    Set recTabelaUTENTES = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    Combo.Redraw = True
End Sub

Public Sub CarregacboNumFunc(ByRef Combo As SSDBCombo, ByVal cInst)
    Dim mBDBaseDados As Database
    Dim recTabelaFUNC As Recordset
    Dim cLinha
    Dim cSql
    
    Combo.Redraw = False
    Combo.RemoveAll
    
    ' começa a construir o Sql
    cSql = "SELECT NUM_FUNCIONARIO,NOME FROM FUNCIONARIOS WHERE ISNULL(DATA_DEMISSAO)"
    ' se seleccionou Instituição tem de filtrar
    If cInst <> "<Todas as Instituições>" Then
        cSql = cSql & " AND COD_INST='" & cCodificaInstituicao(cInst) & "'"
    End If
    ' Estabelace a ordem dos registos
    cSql = cSql & " ORDER BY NUM_FUNCIONARIO ASC"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTabelaFUNC = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    
    If Not recTabelaFUNC.EOF And Not recTabelaFUNC.BOF Then
        While Not recTabelaFUNC.EOF
            Combo.AddItem recTabelaFUNC!NUM_FUNCIONARIO & vbTab & _
                            recTabelaFUNC!Nome
            recTabelaFUNC.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTabelaFUNC.Close
    Set recTabelaFUNC = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    Combo.Redraw = True
End Sub

Public Function cDescodificaMes(ByVal cCodigo)
    Dim mBDBaseDados As Database
    Dim recTABMESES As Recordset
    Dim cSql
    
    cSql = "SELECT COD_MES,NOME FROM TABMESES"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    Set recTABMESES = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
        
    recTABMESES.FindFirst "COD_MES='" & cCodigo & "'"
    
    If recTABMESES.NoMatch Then
        cDescodificaMes = ""
    Else
        cDescodificaMes = recTABMESES!Nome
    End If
    recTABMESES.Close
    Set recTABMESES = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Function lNovoNumSocio() As Long
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recNumSocio As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT NUM_SOCIO FROM SOCIOS"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recNumSocio = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    ' define qual o proximo número de utente
    With recNumSocio
        If .EOF And .BOF Then
            lNovoNumSocio = "1"
        Else
            .MoveLast
            lNovoNumSocio = str(recNumSocio!NUM_SOCIO + 1)
        End If
    End With
   ' fecha a tabela
    recNumSocio.Close
    Set recNumSocio = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Function lNovoNumFuncionario() As Long
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recNumFuncionario As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT NUM_FUNCIONARIO FROM FUNCIONARIOS"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recNumFuncionario = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    ' define qual o proximo número de utente
    With recNumFuncionario
        If .EOF And .BOF Then
            lNovoNumFuncionario = "1"
        Else
            .MoveLast
            lNovoNumFuncionario = str(recNumFuncionario!NUM_FUNCIONARIO + 1)
        End If
    End With
   ' fecha a tabela
    recNumFuncionario.Close
    Set recNumFuncionario = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Sub CarregacboSalas(ByRef Combo As SSDBCombo, ByVal cInst)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABSALAS As Recordset
    Dim cSql
    
    Combo.Redraw = False
    ' remove tudo o que a combo tiver
    Combo.RemoveAll
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT NOME, VALENCIA FROM TABSALAS"
    If Trim$(cInst) <> vbNullString Then
        cSql = cSql & " WHERE COD_INST='" & cCodificaInstituicao(cInst) & "'"
    End If
    cSql = cSql & " ORDER BY NOME;"
    
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABSALAS = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    'carrega a combo
    If Not recTABSALAS.EOF And Not recTABSALAS.BOF Then
        While Not recTABSALAS.EOF
            Combo.AddItem recTABSALAS!Nome
            recTABSALAS.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTABSALAS.Close
    Set recTABSALAS = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing

    Combo.Redraw = True
End Sub
Public Function cCodificaSala(ByVal cCodigo_Inst, ByVal cCodigo_Sala)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABSALAS As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT NOME,COD_INST,COD_SALA FROM TABSALAS"
    
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABSALAS = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    ' procura o nome
    recTABSALAS.FindFirst "COD_INST='" & cCodigo_Inst & "' AND NOME='" & cCodigo_Sala & "'"
    ' se não encontra retorna
    If recTABSALAS.NoMatch Then
        cCodificaSala = ""
    Else
        cCodificaSala = recTABSALAS!COD_SALA
    End If
    ' fecha a tabela
    recTABSALAS.Close
    Set recTABSALAS = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Function cDescodificaSala(ByVal cCodigo_Inst, ByVal cCodigo_Sala)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABSALAS As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT NOME,COD_INST,COD_SALA FROM TABSALAS"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABSALAS = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    ' procura o código
    recTABSALAS.FindFirst "COD_INST='" & cCodigo_Inst & "' AND COD_SALA='" & cCodigo_Sala & "'"
    ' se não encontra retorna
    If recTABSALAS.NoMatch Then
        cDescodificaSala = ""
    Else
        cDescodificaSala = recTABSALAS!Nome
    End If
    ' fecha a tabela
    recTABSALAS.Close
    Set recTABSALAS = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function
Public Function cDescodificaValencia(ByVal cCodigo_Inst, ByVal cCodigo_Sala)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABSALAS As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT VALENCIA,NOME,COD_INST,COD_SALA FROM TABSALAS"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABSALAS = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    ' procura o código
    recTABSALAS.FindFirst "COD_INST='" & cCodigo_Inst & "' AND COD_SALA='" & cCodigo_Sala & "'"
    ' se não encontra retorna
    If recTABSALAS.NoMatch Then
        cDescodificaValencia = ""
    Else
        cDescodificaValencia = vFiltraCamposNulos(recTABSALAS!VALENCIA)
    End If
    ' fecha a tabela
    recTABSALAS.Close
    Set recTABSALAS = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Function cDescodificaNomeEducadora(ByVal cCodigo_Inst, ByVal cCodigo_Sala)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABSALAS As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT NOME_EDUCADORA,COD_INST,COD_SALA FROM TABSALAS"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABSALAS = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    ' procura o código
    recTABSALAS.FindFirst "COD_INST='" & cCodigo_Inst & "' AND COD_SALA='" & cCodigo_Sala & "'"
    ' se não encontra retorna
    If recTABSALAS.NoMatch Then
        cDescodificaNomeEducadora = ""
    Else
        cDescodificaNomeEducadora = vFiltraCamposNulos(recTABSALAS!NOME_EDUCADORA)
    End If
    ' fecha a tabela
    recTABSALAS.Close
    Set recTABSALAS = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Sub CarregacboInstituicao(ByRef Combo As SSDBCombo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABINSTITUICAO As Recordset
    Dim cSql
    
    Combo.Redraw = False
    ' remove tudo o que a combo tiver
    Combo.RemoveAll
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT NOME FROM TABINSTITUICAO ORDER BY NOME"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABINSTITUICAO = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    'carrega a combo
    If Not recTABINSTITUICAO.EOF And Not recTABINSTITUICAO.BOF Then
        While Not recTABINSTITUICAO.EOF
            Combo.AddItem recTABINSTITUICAO!Nome
            recTABINSTITUICAO.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTABINSTITUICAO.Close
    Set recTABINSTITUICAO = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    
    Combo.Redraw = True
End Sub
Public Sub CarregacboCategoria(ByRef Combo As SSDBCombo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABCATEGORIA As Recordset
    Dim cSql
    
    Combo.Redraw = False
    ' remove tudo o que a combo tiver
    Combo.RemoveAll
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT DESCRICAO FROM TABCATEGORIA ORDER BY DESCRICAO"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABCATEGORIA = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    'carrega a combo
    If Not recTABCATEGORIA.EOF And Not recTABCATEGORIA.BOF Then
        While Not recTABCATEGORIA.EOF
            Combo.AddItem recTABCATEGORIA!DESCRICAO
            recTABCATEGORIA.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTABCATEGORIA.Close
    Set recTABCATEGORIA = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    
    Combo.Redraw = True
End Sub

Public Sub CarregacboSector(ByRef Combo As SSDBCombo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABSECTOR As Recordset
    Dim cSql
    
    Combo.Redraw = False
    ' remove tudo o que a combo tiver
    Combo.RemoveAll
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT DESCRICAO FROM TABSECTOR ORDER BY DESCRICAO"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABSECTOR = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    'carrega a combo
    If Not recTABSECTOR.EOF And Not recTABSECTOR.BOF Then
        While Not recTABSECTOR.EOF
            Combo.AddItem recTABSECTOR!DESCRICAO
            recTABSECTOR.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTABSECTOR.Close
    Set recTABSECTOR = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    
    Combo.Redraw = True
End Sub
Public Sub CarregacboEstadoCivil(ByRef Combo As SSDBCombo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABESTADOCIVIL As Recordset
    Dim cSql
    
    Combo.Redraw = False
    ' remove tudo o que a combo tiver
    Combo.RemoveAll
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT DESCRICAO FROM TABESTADOCIVIL ORDER BY DESCRICAO"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABESTADOCIVIL = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    'carrega a combo
    If Not recTABESTADOCIVIL.EOF And Not recTABESTADOCIVIL.BOF Then
        While Not recTABESTADOCIVIL.EOF
            Combo.AddItem recTABESTADOCIVIL!DESCRICAO
            recTABESTADOCIVIL.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTABESTADOCIVIL.Close
    Set recTABESTADOCIVIL = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    
    Combo.Redraw = True
End Sub

Public Function cCodificaInstituicao(ByVal cCodigo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABINSTITUICAO As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT NOME,COD_INST FROM TABINSTITUICAO"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABINSTITUICAO = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    ' procura o nome
    recTABINSTITUICAO.FindFirst "NOME='" & cCodigo & "'"
    ' se não encontra retorna
    If recTABINSTITUICAO.NoMatch Then
        cCodificaInstituicao = ""
    Else
        cCodificaInstituicao = recTABINSTITUICAO!COD_INST
    End If
    ' fecha a tabela
    recTABINSTITUICAO.Close
    Set recTABINSTITUICAO = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function
Public Function cCodificaCategoria(ByVal cCodigo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABCATEGORIA As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT DESCRICAO,COD_CATEGORIA FROM TABCATEGORIA"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABCATEGORIA = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    ' procura o nome
    recTABCATEGORIA.FindFirst "DESCRICAO='" & cCodigo & "'"
    ' se não encontra retorna
    If recTABCATEGORIA.NoMatch Then
        cCodificaCategoria = ""
    Else
        cCodificaCategoria = recTABCATEGORIA!COD_CATEGORIA
    End If
    ' fecha a tabela
    recTABCATEGORIA.Close
    Set recTABCATEGORIA = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Function cCodificaSector(ByVal cCodigo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABSECTOR As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT DESCRICAO,COD_SECTOR FROM TABSECTOR"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABSECTOR = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    ' procura o nome
    recTABSECTOR.FindFirst "DESCRICAO='" & cCodigo & "'"
    ' se não encontra retorna
    If recTABSECTOR.NoMatch Then
        cCodificaSector = ""
    Else
        cCodificaSector = recTABSECTOR!COD_SECTOR
    End If
    ' fecha a tabela
    recTABSECTOR.Close
    Set recTABSECTOR = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Function cCodificaEstadoCivil(ByVal cCodigo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABESTADOCIVIL As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT DESCRICAO,COD_ESTADO_CIVIL FROM TABESTADOCIVIL"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABESTADOCIVIL = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    ' procura o nome
    recTABESTADOCIVIL.FindFirst "DESCRICAO='" & cCodigo & "'"
    ' se não encontra retorna
    If recTABESTADOCIVIL.NoMatch Then
        cCodificaEstadoCivil = ""
    Else
        cCodificaEstadoCivil = recTABESTADOCIVIL!COD_ESTADO_CIVIL
    End If
    ' fecha a tabela
    recTABESTADOCIVIL.Close
    Set recTABESTADOCIVIL = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Function lNovoNumUtente() As Long
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recNumUtente As Recordset
    Dim cSql
        
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT NUM_UTENTE FROM UTENTES ORDER BY NUM_UTENTE"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recNumUtente = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    ' define qual o proximo número de utente
    With recNumUtente
        If .EOF And .BOF Then
            lNovoNumUtente = 1
        Else
            .MoveLast
            lNovoNumUtente = recNumUtente!NUM_UTENTE + 1
        End If
    End With
    ' fecha a tabela
    recNumUtente.Close
    Set recNumUtente = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function
Public Function cNIF(ByVal lNum_Utente)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recNumUtente As Recordset
    Dim cSql
        
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT NUM_CONTRIBUINTE FROM UTENTES WHERE NUM_UTENTE=" & lNum_Utente
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recNumUtente = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    recNumUtente.MoveLast
    cNIF = recNumUtente!NUM_CONTRIBUINTE
    ' fecha a tabela
    recNumUtente.Close
    Set recNumUtente = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Function cNIF_CAIF(ByVal lNum_Utente)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recNumUtente As Recordset
    Dim cSql
        
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT NUM_CONTRIBUINTE FROM UTENTES_IDOSOS WHERE NUM_UTENTE=" & lNum_Utente
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recNumUtente = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    recNumUtente.MoveLast
    cNIF_CAIF = recNumUtente!NUM_CONTRIBUINTE
    ' fecha a tabela
    recNumUtente.Close
    Set recNumUtente = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Function lNovoNumUtenteCAIF() As Long
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recNumUtente As Recordset
    Dim cSql
        
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT NUM_UTENTE FROM UTENTES_IDOSOS ORDER BY NUM_UTENTE"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recNumUtente = mBDBaseDados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    ' define qual o proximo número de utente
    With recNumUtente
        If .EOF And .BOF Then
            lNovoNumUtenteCAIF = 1
        Else
            .MoveLast
            lNovoNumUtenteCAIF = recNumUtente!NUM_UTENTE + 1
        End If
    End With
    ' fecha a tabela
    recNumUtente.Close
    Set recNumUtente = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Function cDescodificaInstituicao(ByVal cCodigo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABINSTITUICAO As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT NOME,COD_INST FROM TABINSTITUICAO"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABINSTITUICAO = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    ' procura o código
    recTABINSTITUICAO.FindFirst "COD_INST='" & cCodigo & "'"
    ' se não encontra retorna
    If recTABINSTITUICAO.NoMatch Then
        cDescodificaInstituicao = ""
    Else
        cDescodificaInstituicao = recTABINSTITUICAO!Nome
    End If
    ' fecha a tabela
    recTABINSTITUICAO.Close
    Set recTABINSTITUICAO = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function
Public Function cDescodificaCategoria(ByVal cCodigo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABCATEGORIA As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT DESCRICAO,COD_CATEGORIA FROM TABCATEGORIA"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABCATEGORIA = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    ' procura o código
    recTABCATEGORIA.FindFirst "COD_CATEGORIA='" & cCodigo & "'"
    ' se não encontra retorna
    If recTABCATEGORIA.NoMatch Then
        cDescodificaCategoria = ""
    Else
        cDescodificaCategoria = recTABCATEGORIA!DESCRICAO
    End If
    ' fecha a tabela
    recTABCATEGORIA.Close
    Set recTABCATEGORIA = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Function cDescodificaSector(ByVal cCodigo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABSECTOR As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT DESCRICAO,COD_SECTOR FROM TABSECTOR"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABSECTOR = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    ' procura o código
    recTABSECTOR.FindFirst "COD_SECTOR='" & cCodigo & "'"
    ' se não encontra retorna
    If recTABSECTOR.NoMatch Then
        cDescodificaSector = ""
    Else
        cDescodificaSector = recTABSECTOR!DESCRICAO
    End If
    ' fecha a tabela
    recTABSECTOR.Close
    Set recTABSECTOR = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function

Public Function cDescodificaEstadoCivil(ByVal cCodigo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABESTADOCIVIL As Recordset
    Dim cSql
    
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT DESCRICAO,COD_ESTADO_CIVIL FROM TABESTADOCIVIL"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABESTADOCIVIL = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    ' procura o código
    recTABESTADOCIVIL.FindFirst "COD_ESTADO_CIVIL='" & cCodigo & "'"
    ' se não encontra retorna
    If recTABESTADOCIVIL.NoMatch Then
        cDescodificaEstadoCivil = ""
    Else
        cDescodificaEstadoCivil = recTABESTADOCIVIL!DESCRICAO
    End If
    ' fecha a tabela
    recTABESTADOCIVIL.Close
    Set recTABESTADOCIVIL = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
End Function


Public Sub CalcHoras()
'    Horas      DateDiff("n", Date & " 15:30", Date & " 19:45") \ 60
'    Minutos    DateDiff("n", Date & " 15:30", Date & " 19:45") Mod 60
End Sub
