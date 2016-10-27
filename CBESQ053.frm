VERSION 5.00
Begin VB.Form frmMensalidadesRecalcular 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recalcular Mensalidades"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "2003"
      Top             =   1485
      Width           =   1200
   End
   Begin VB.CommandButton cmdRecalcular 
      Caption         =   "&Recalcular"
      Height          =   900
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "2002"
      Top             =   1485
      Width           =   1200
   End
   Begin VB.TextBox txtAtencao 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   1200
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "CBESQ053.frx":0000
      Top             =   120
      Width           =   4200
   End
End
Attribute VB_Name = "frmMensalidadesRecalcular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSRecalculaMensalidade As Workspace
Dim mBDRecalculaMensalidade As Database

Dim tBDAberta

Dim cSql
Dim iRespMsgbox

Dim recUTENTES As Recordset

Dim lProxMensalidade_Base
Dim lProxMensalidade
Dim dProx_Comparticipacao
Dim cProx_ESCALAO
Dim dProx_PCT

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSRecalculaMensalidade = DBEngine.CreateWorkspace("RecalculaMensalidade", gUtilizador.Nome, gUtilizador.Password)
    Set mBDRecalculaMensalidade = mWSRecalculaMensalidade.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Recalcula Mensalidade-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdRecalcular_Click()
    Dim mProcessamento As Processamento
    
    Set mProcessamento = New Processamento
    
    ' pede confirmação se quer continuar
    iRespMsgbox = MsgBox("Confirma que quer Recalcular as Mensalidades Automáticas dos Utentes !!!", vbQuestion + vbYesNo, _
                        "Recalcular Mensalidades")
    ' se resposta não sai
    If iRespMsgbox = vbNo Then
        GoTo SairDoProcedimento
    End If
    
On Error GoTo TrataErro
    
' Vai a UTENTES
    cSql = "SELECT * FROM UTENTES WHERE TIPO_MENSALIDADE = TRUE "
    cSql = cSql & " AND COD_INST <> '999'"

    Set recUTENTES = mBDRecalculaMensalidade.OpenRecordset(cSql, dbOpenDynaset)
    
    If Not (recUTENTES.EOF Or recUTENTES.BOF) Then
       recUTENTES.MoveFirst
       While Not recUTENTES.EOF
'           mWSRecalculaMensalidade.BeginTrans
           recUTENTES.Edit

           Call CalculaMensalidadeAutomatica
           recUTENTES.Fields("PROX_MENSALIDADE_BASE") = lProxMensalidade_Base
           recUTENTES.Fields("PROX_MENSALIDADE") = lProxMensalidade
           recUTENTES.Fields("PROX_ESCALAO") = cProx_ESCALAO
           recUTENTES.Fields("PROX_PCT") = dProx_PCT

           recUTENTES.Update
'           mWSRecalculaMensalidade.CommitTrans
           recUTENTES.MoveNext
       Wend
    End If
    
    recUTENTES.Close
    Set recUTENTES = Nothing
    
    ' actualizou as Mensalidades
    MsgBox "Recalculo de Mensalidades concluída com sucesso !!!", vbInformation + vbOKOnly, "Recalculo de Mensalidades"
    
    GoTo SairDoProcedimento

TrataErro:
    mWSRecalculaMensalidade.Rollback
    Call ErrosGerais("Recalculo de Mensalidades", Err.Number, Err.Description)
    Resume SairDoProcedimento

SairDoProcedimento:
End Sub

Public Sub CalculaMensalidadeAutomatica()
    Dim recTABMENSALIDADE As Recordset
    Dim recTABMES As Recordset
    Dim cSql
    Dim lValorPerCapita
    Dim lValorRendimentos
    Dim lValorDespesas
    Dim lValorDespesasLimitadas
    Dim lValorProxMensalidadeFinal
    Dim tEscalaoMaximo
    
    ' Calcula Rendimentos
    lValorRendimentos = (recUTENTES.Fields("R_TD_1").Value + recUTENTES.Fields("R_P_1").Value + _
                    recUTENTES.Fields("R_PA_1").Value + recUTENTES.Fields("R_TI_1").Value + _
                    recUTENTES.Fields("R_R_1").Value + recUTENTES.Fields("R_RSI_1").Value + _
                    recUTENTES.Fields("R_SD_1").Value + recUTENTES.Fields("R_AF_1").Value + _
                    recUTENTES.Fields("R_O_1").Value) + (recUTENTES.Fields("R_TD_2").Value + _
                    recUTENTES.Fields("R_P_2").Value + recUTENTES.Fields("R_PA_2").Value + _
                    recUTENTES.Fields("R_TI_2").Value + recUTENTES.Fields("R_R_2").Value + _
                    recUTENTES.Fields("R_RSI_2").Value + recUTENTES.Fields("R_SD_2").Value + _
                    recUTENTES.Fields("R_AF_2").Value + recUTENTES.Fields("R_O_2").Value)

    ' Calcula Despesas
    lValorDespesas = (recUTENTES.Fields("D_IRS_1").Value + recUTENTES.Fields("D_SS_1").Value + _
                        recUTENTES.Fields("D_O_1").Value) + _
                     (recUTENTES.Fields("D_IRS_2").Value + recUTENTES.Fields("D_SS_2").Value + _
                        recUTENTES.Fields("D_O_2").Value)
    ' Calcula Despesas Limitadas
    ' 20150515 foi acrescentado o CAMPO IMI que representa o ERPI
    lValorDespesasLimitadas = (recUTENTES.Fields("D_JEH_1").Value + recUTENTES.Fields("D_SC_1").Value + _
                                recUTENTES.Fields("D_T_1").Value + recUTENTES.Fields("D_IMI_1").Value + _
                                recUTENTES.Fields("D_JEH_2").Value + recUTENTES.Fields("D_SC_2").Value + _
                                recUTENTES.Fields("D_T_2").Value + recUTENTES.Fields("D_IMI_2").Value)
    
    If lValorDespesasLimitadas > 6060 Then
        lValorDespesasLimitadas = 6060
    End If
    
    ' Calcula Valor Per-Capita segundo formula dada pelo CBESQ
'     20150515 foi retirado o CAMPO IMI que representa o ERPI
'                        recUTENTES.Fields("D_IMI_1").Value + recUTENTES.Fields("D_O_1").Value)
'                        recUTENTES.Fields("D_IMI_2").Value + recUTENTES.Fields("D_O_2").Value)
'    lValorPerCapita = (((recUTENTES.Fields("R_TD_1").Value + recUTENTES.Fields("R_P_1").Value + _
'                        recUTENTES.Fields("R_PA_1").Value + recUTENTES.Fields("R_TI_1").Value + _
'                        recUTENTES.Fields("R_R_1").Value + recUTENTES.Fields("R_RSI_1").Value + _
'                        recUTENTES.Fields("R_SD_1").Value + recUTENTES.Fields("R_AF_1").Value + _
'                        recUTENTES.Fields("R_O_1").Value) + (recUTENTES.Fields("R_TD_2").Value + _
'                        recUTENTES.Fields("R_P_2").Value + recUTENTES.Fields("R_PA_2").Value + _
'                        recUTENTES.Fields("R_TI_2").Value + recUTENTES.Fields("R_R_2").Value + _
'                        recUTENTES.Fields("R_RSI_2").Value + recUTENTES.Fields("R_SD_2").Value + _
'                        recUTENTES.Fields("R_AF_2").Value + recUTENTES.Fields("R_O_2").Value)) - _
'                        ((recUTENTES.Fields("D_IRS_1").Value + recUTENTES.Fields("D_SS_1").Value + _
'                        recUTENTES.Fields("D_O_1").Value) + _
'                        (recUTENTES.Fields("D_IRS_2").Value + recUTENTES.Fields("D_SS_2").Value + _
'                        recUTENTES.Fields("D_O_2").Value) + _
'                        lValorDespesasLimitadas)) / IIf(recUTENTES.Fields("AGREGADO").Value = 0, 1 * 12, recUTENTES.Fields("AGREGADO").Value * 12)
     lValorPerCapita = (lValorRendimentos - (lValorDespesas + lValorDespesasLimitadas)) / IIf(recUTENTES.Fields("AGREGADO").Value = 0, 1 * 12, recUTENTES.Fields("AGREGADO").Value * 12)
                        
    'Verifica se é para por no Escalão Máximo
    If (lValorRendimentos + lValorDespesas + lValorDespesasLimitadas) > 0 Then
        tEscalaoMaximo = False
    Else
        tEscalaoMaximo = True
    End If
                        
    ' Vai a TABMENSALIDADES buscar o valor da Mensalidade Base
    If tEscalaoMaximo Then
        cSql = "SELECT * FROM TABMENSALIDADE WHERE COD_INST = '" & recUTENTES.Fields("COD_INST").Value & "'"
        cSql = cSql & " AND COD_SALA = '" & recUTENTES.Fields("COD_SALA").Value & "'"
        cSql = cSql & " AND COD_MENSALIDADE = '006'"
    Else
        cSql = "SELECT * FROM TABMENSALIDADE WHERE VALOR_MAX >= " & Int(lValorPerCapita) & " AND " & _
            "VALOR_MIN <= " & Int(lValorPerCapita)
    cSql = cSql & " AND COD_INST = '" & recUTENTES.Fields("COD_INST").Value & "'"
    cSql = cSql & " AND COD_SALA = '" & recUTENTES.Fields("COD_SALA").Value & "'"
    End If
    Set recTABMENSALIDADE = mBDRecalculaMensalidade.OpenRecordset(cSql, dbOpenSnapshot)

    cProx_ESCALAO = recTABMENSALIDADE!COD_MENSALIDADE
    If tEscalaoMaximo Then
        dProx_PCT = 100
    Else
        dProx_PCT = recTABMENSALIDADE!PCT
    End If
    
    If tEscalaoMaximo Then
        lProxMensalidade_Base = Round(recTABMENSALIDADE!MENS_MAX, 2)
    Else
        lProxMensalidade_Base = Round(lValorPerCapita * (recTABMENSALIDADE!PCT / 100), 2)
    End If

    ' se o calculo for inferior ao MENS_MIN passa a MENS_MIN
    If lProxMensalidade_Base < recTABMENSALIDADE!MENS_MIN Then
        lProxMensalidade_Base = recTABMENSALIDADE!MENS_MIN
    ' se o calculo for superior ao MENS_MAX passa a MENS_MAX
    ElseIf lProxMensalidade_Base > recTABMENSALIDADE!MENS_MAX Then
        lProxMensalidade_Base = recTABMENSALIDADE!MENS_MAX
    End If
    
'   Desconto tem de ser aqui
    lValorProxMensalidadeFinal = lProxMensalidade_Base
    lProxMensalidade_Base = lValorProxMensalidadeFinal - (lValorProxMensalidadeFinal * (recUTENTES.Fields("COMPARTICIPACAO").Value / 100))

' Vai a TABMESES buscar o valor da percentagem a adicionar a Mensalidade Base
    cSql = "SELECT PERCENTAGEMMENSAL,PERCENTAGEMMENSALATL FROM TABMESES WHERE COD_MES = '" & Format$(Month(Date), "00") & "'"
    Set recTABMES = mBDRecalculaMensalidade.OpenRecordset(cSql, dbOpenSnapshot)

    ' Se Sala for ATL
    ' Divide a Mensalidade Base por 11 Mês de Agosto
    If (recUTENTES.Fields("COD_INST").Value = "001" And recUTENTES.Fields("COD_SALA").Value = "008") Or _
        (recUTENTES.Fields("COD_INST").Value = "002" And recUTENTES.Fields("COD_SALA").Value = "006") Then
        lProxMensalidade = Round(lProxMensalidade_Base / recTABMES!PERCENTAGEMMENSALATL, 2)
    ' Se Sala for Outras Valências
    ' Divide a Mensalidade Base por 10 Mês de Julho
    Else
        lProxMensalidade = Round(lProxMensalidade_Base / recTABMES!PERCENTAGEMMENSAL, 2)
    End If

    dProx_Comparticipacao = 0

    recTABMENSALIDADE.Close
    Set recTABMENSALIDADE = Nothing
End Sub

Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Private Sub Form_Load()
    CenterMe Me
    LoadResStrings Me
    
    Call AlteraWindowList(Me.Caption)
    
    Me.Show
    DoEvents
    
    tBDAberta = tAbreBD
    
    txtAtencao.Text = " Atenção" & vbCrLf & _
                    " Vai Recalcular as Mensalidades Automáticas dos Utentes."
    
    
    End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSRecalculaMensalidade.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSRecalculaMensalidade = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmMensalidadesRecalcular = Nothing
End Sub





