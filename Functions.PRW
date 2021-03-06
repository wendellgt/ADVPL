#INCLUDE "RWMAKE.CH"
#INCLUDE "TOPCONN.CH"  
#INCLUDE "PROTHEUS.CH"


/*/{Protheus.doc} UsrRetNome
@type function
@systemOper Windows
@author Wendell Tavares
@since 03/03/2010
@version 1.0
@description Retorna o nome do usuario com padrao no cadastro
@param [cCodUser], string, Codigo do Usu�rio
@return string, Nome do Usu�rio
@table Sigapss
@link https://github.com/wendellgt
/*/

User Function UsrRetNom(cCodUser)
	Local cAlias := Alias()
	Local cSavOrd := IndexOrd()
	Local cNome
	Default cCodUser := '000000'
	
	PswOrder(1)
	If	!Empty(cCodUser) .And. PswSeek(cCodUser)
		cNome := PswRet(1)[1][4]
	Else
		cNome := SPACE(15)
	EndIf
	
	dbSelectArea(cAlias)
	dbSetOrder(cSavOrd)

Return(cNome)


/*/{Protheus.doc} zCriaPar
Fun��o para cria��o de par�metros (SX6)
@type function
@author Atilio
@since 12/11/2015
@version 1.0
    @param aPars, Array, Array com os par�metros do sistema
    @example aAdd(aPars, {"MV_X_EMAIL", "C", "e-Mail da empresa", "email@teste.com"} )
    		 u_zCriaPar(aPars)
    @see https://terminaldeinformacao.com
    @obs Abaixo a estrutura do array:
        [01] - Par�metro (ex.: "MV_X_TST")
        [02] - Tipo (ex.: "C")
        [03] - Descri��o (ex.: "Par�metro Teste")
        [04] - Conte�do (ex.: "123;456;789")
/*/
 
User Function zCriaPar(aPars)
    Local nAtual        := 0
    Local aArea        := GetArea()
    Local aAreaX6        := SX6->(GetArea())
    Default aPars        := {}
     
    DbSelectArea("SX6")
    SX6->(DbGoTop())
     
    //Percorrendo os par�metros e gerando os registros
    For nAtual := 1 To Len(aPars)
        //Se n�o conseguir posicionar no par�metro cria
        If !(SX6->(DbSeek(xFilial("SX6")+aPars[nAtual][1])))
            RecLock("SX6",.T.)
                //Geral
                X6_VAR        :=    aPars[nAtual][1]
                X6_TIPO    :=    aPars[nAtual][2]
                X6_PROPRI    :=    "U"
                //Descri��o
                X6_DESCRIC    :=    aPars[nAtual][3]
                X6_DSCSPA    :=    aPars[nAtual][3]
                X6_DSCENG    :=    aPars[nAtual][3]
                //Conte�do
                X6_CONTEUD    :=    aPars[nAtual][4]
                X6_CONTSPA    :=    aPars[nAtual][4]
                X6_CONTENG    :=    aPars[nAtual][4]
            SX6->(MsUnlock())
        EndIf
    Next
     
    RestArea(aAreaX6)
    RestArea(aArea)
Return

/*/{Protheus.doc} zPutSX1
Fun��o para criar Grupo de Perguntas
@author Atilio
@since 09/06/2017
@version 1.0
@type function
    @param cGrupo,    characters, Grupo de Perguntas       (ex.: X_TESTE)
    @param cOrdem,    characters, Ordem da Pergunta        (ex.: 01, 02, 03, ...)
    @param cTexto,    characters, Texto da Pergunta        (ex.: Produto De, Produto At�, Data De, ...)
    @param cMVPar,    characters, MV_PAR?? da Pergunta     (ex.: MV_PAR01, MV_PAR02, MV_PAR03, ...)
    @param cVariavel, characters, Vari�vel da Pergunta     (ex.: MV_CH0, MV_CH1, MV_CH2, ...)
    @param cTipoCamp, characters, Tipo do Campo            (C = Caracter, N = Num�rico, D = Data)
    @param nTamanho,  numeric,    Tamanho da Pergunta      (M�ximo de 60)
    @param nDecimal,  numeric,    Tamanho de Decimais      (M�ximo de 9)
    @param cTipoPar,  characters, Tipo do Par�metro        (G = Get, C = Combo, F = Escolha de Arquivos, K = Check Box)
    @param cValid,    characters, Valida��o da Pergunta    (ex.: Positivo(), u_SuaFuncao(), ...)
    @param cF3,       characters, Consulta F3 da Pergunta  (ex.: SB1, SA1, ...)
    @param cPicture,  characters, M�scara do Par�metro     (ex.: @!, @E 999.99, ...)
    @param cDef01,    characters, Primeira op��o do combo
    @param cDef02,    characters, Segunda op��o do combo
    @param cDef03,    characters, Terceira op��o do combo
    @param cDef04,    characters, Quarta op��o do combo
    @param cDef05,    characters, Quinta op��o do combo
    @param cHelp,     characters, Texto de Help do par�metro
    @obs Fun��o foi criada, pois a partir de algumas vers�es do Protheus 12, a fun��o padr�o PutSX1 n�o funciona (por medidas de seguran�a)
    @example Abaixo um exemplo de como criar um grupo de perguntas
     
    cPerg    := "X_TST"
     
    cValid   := ""
    cF3      := ""
    cPicture := ""
    cDef01   := ""
    cDef02   := ""
    cDef03   := ""
    cDef04   := ""
    cDef05   := ""
     
    u_zPutSX1(cPerg, "01", "Produto De?",       "MV_PAR01", "MV_CH0", "C", TamSX3('B1_COD')[01], 0, "G", cValid,       "SB1", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o produto inicial")
    u_zPutSX1(cPerg, "02", "Produto At�?",      "MV_PAR02", "MV_CH1", "C", TamSX3('B1_COD')[01], 0, "G", cValid,       "SB1", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o produto final")
    u_zPutSX1(cPerg, "03", "A partir da Data?", "MV_PAR03", "MV_CH2", "D", 08,                   0, "G", cValid,       cF3,   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a data inicial a ser considerada")
    u_zPutSX1(cPerg, "04", "M�dia maior que?",  "MV_PAR04", "MV_CH3", "N", 09,                   2, "G", "Positivo()", cF3,   "@E 999,999.99", cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a m�dia de atraso que ser� considerada")
    u_zPutSX1(cPerg, "05", "Tipo de Saldos?",   "MV_PAR05", "MV_CH4", "N", 01,                   0, "C", cValid,       cF3,   cPicture,        "Todos", "Maior que 0", "Menor que 0", "Zerados", cDef05, "Informe o tipo de saldo a ser considerado")
    u_zPutSX1(cPerg, "06", "Tipos de Produto?", "MV_PAR06", "MV_CH5", "C", 60,                   0, "K", cValid,       cF3,   cPicture,        "PA",    "PI",          "MP",          cDef04,    cDef05, "Informe os tipos de produto que ser�o considerados")
    u_zPutSX1(cPerg, "07", "Caminho de Log?",   "MV_PAR07", "MV_CH6", "C", 60,                   0, "F", cValid,       cF3,   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o caminho para gera��o do log")
/*/
 
User Function zPutSX1(cGrupo, cOrdem, cTexto, cMVPar, cVariavel, cTipoCamp, nTamanho, nDecimal, cTipoPar, cValid, cF3, cPicture, cDef01, cDef02, cDef03, cDef04, cDef05, cHelp)
    Local aArea       := GetArea()
    Local cChaveHelp  := ""
    Local nPreSel     := 0
    Default cGrupo    := Space(10)
    Default cOrdem    := Space(02)
    Default cTexto    := Space(30)
    Default cMVPar    := Space(15)
    Default cVariavel := Space(6)
    Default cTipoCamp := Space(1)
    Default nTamanho  := 0
    Default nDecimal  := 0
    Default cTipoPar  := "G"
    Default cValid    := Space(60)
    Default cF3       := Space(6)
    Default cPicture  := Space(40)
    Default cDef01    := Space(15)
    Default cDef02    := Space(15)
    Default cDef03    := Space(15)
    Default cDef04    := Space(15)
    Default cDef05    := Space(15)
    Default cHelp     := ""
     
    //Se tiver Grupo, Ordem, Texto, Par�metro, Vari�vel, Tipo e Tamanho, continua para a cria��o do par�metro
    If !Empty(cGrupo) .And. !Empty(cOrdem) .And. !Empty(cTexto) .And. !Empty(cMVPar) .And. !Empty(cVariavel) .And. !Empty(cTipoCamp) .And. nTamanho != 0
         
        //Defini��o de vari�veis
        cGrupo     := PadR(cGrupo, Len(SX1->X1_GRUPO), " ")           //Adiciona espa�os a direita para utiliza��o no DbSeek
        cChaveHelp := "P." + AllTrim(cGrupo) + AllTrim(cOrdem) + "."  //Define o nome da pergunta
        cMVPar     := Upper(cMVPar)                                   //Deixa o MV_PAR tudo em mai�sculo
        nPreSel    := Iif(cTipoPar == "C", 1, 0)                      //Se for Combo, o pr�-selecionado ser� o Primeiro
        cDef01     := Iif(cTipoPar == "F", "56", cDef01)              //Se for File, muda a defini��o para ser tanto Servidor quanto Local
        nTamanho   := Iif(nTamanho > 60, 60, nTamanho)                //Se o tamanho for maior que 60, volta para 60 - Limita��o do Protheus
        nDecimal   := Iif(nDecimal > 9,  9,  nDecimal)                //Se o decimal for maior que 9, volta para 9
        nDecimal   := Iif(cTipoPar == "N", nDecimal, 0)               //Se n�o for par�metro do tipo num�rico, ser� 0 o Decimal
        cTipoCamp  := Upper(cTipoCamp)                                //Deixa o tipo do Campo em mai�sculo
        cTipoCamp  := Iif(! cTipoCamp $ 'C;D;N;', 'C', cTipoCamp)     //Se o tipo do Campo n�o estiver entre Caracter / Data / Num�rico, ser� Caracter
        cTipoPar   := Upper(cTipoPar)                                 //Deixa o tipo do Par�metro em mai�sculo
        cTipoPar   := Iif(Empty(cTipoPar), 'G', cTipoPar)             //Se o tipo do Par�metro estiver em branco, ser� um Get
        nTamanho   := Iif(cTipoPar == "C", 1, nTamanho)               //Se for Combo, o tamanho ser� 1
     
        DbSelectArea('SX1')
        SX1->(DbSetOrder(1)) // Grupo + Ordem
     
        //Se n�o conseguir posicionar, a pergunta ser� criada
        If ! SX1->(DbSeek(cGrupo + cOrdem))
            RecLock('SX1', .T.)
                X1_GRUPO   := cGrupo
                X1_ORDEM   := cOrdem
                X1_PERGUNT := cTexto
                X1_PERSPA  := cTexto
                X1_PERENG  := cTexto
                X1_VAR01   := cMVPar
                X1_VARIAVL := cVariavel
                X1_TIPO    := cTipoCamp
                X1_TAMANHO := nTamanho
                X1_DECIMAL := nDecimal
                X1_GSC     := cTipoPar
                X1_VALID   := cValid
                X1_F3      := cF3
                X1_PICTURE := cPicture
                X1_DEF01   := cDef01
                X1_DEFSPA1 := cDef01
                X1_DEFENG1 := cDef01
                X1_DEF02   := cDef02
                X1_DEFSPA2 := cDef02
                X1_DEFENG2 := cDef02
                X1_DEF03   := cDef03
                X1_DEFSPA3 := cDef03
                X1_DEFENG3 := cDef03
                X1_DEF04   := cDef04
                X1_DEFSPA4 := cDef04
                X1_DEFENG4 := cDef04
                X1_DEF05   := cDef05
                X1_DEFSPA5 := cDef05
                X1_DEFENG5 := cDef05
                X1_PRESEL  := nPreSel
                 
                //Se tiver Help da Pergunta
                If !Empty(cHelp)
                    X1_HELP    := ""
                     
                    fPutHelp(cChaveHelp, cHelp)
                EndIf
            SX1->(MsUnlock())
        EndIf
    EndIf
     
    RestArea(aArea)
Return