#include "Protheus.ch"


/*/{Protheus.doc} Frel0002.prw
@type function
@systemOper Windows
@author Wendell Tavares
@since 01/08/2019
@version 1.0
@description Relatorio de parametros para gera��o dos impostos de servico
@link https://github.com/wendellgt

/*/


User Function Frel0002()

	cPergunte := "R0002"
	criaPerg(cPergunte)
	 
	pergunte(cPergunte,.T.)

	Processa( {|| U_ERel0002() }, "Aguarde...", "Gerando Relat�rio...",.F.)

Return


User Function ERel0002()
 
Local aCabec 	:= {} 
Local aDados 	:= {}
Local aRet 		:= {}
Local nl
Local aParame	:= {}
Local aCampos	:= {}
Local aDados	:= {}
Local aRet		:= {}
LOcal cQuery 	:= 'SELECT '

//Verifica se tem o excel instalado 
If !ApOleClient("MSExcel") 
	MsgAlert("Microsoft Excel n�o instalado!") 
	Return 
EndIf

/*       
cType:
	C = Caracter
	D = Data
	N = Numerico (Totaliza)
	M = Moeda	 (Totaliza)
*/  
//Cabe�alho da planilha
AADD(aCabec, {"TIPO"				,"C"})
AADD(aCabec, {"PARAMETRO"			,"C"})
AADD(aCabec, {"CONTEUDO"			,"C"})
AADD(aCabec, {"ESPERADO"			,"C"})
AADD(aCabec, {"DESCRICAO"			,"C"})
AADD(aCabec, {"OBSERVACAO"			,"C"}) 


//Campos obrigatorios
Aadd(aCampos, { 'TODOS', 'A2_COD', 		'', 	'', 		'C�digo do Fornecedor' })
Aadd(aCampos, { 'TODOS', 'A2_NATUREZ', 	'', 	'', 		'C�digo da Natureza' })
Aadd(aCampos, { 'TODOS', 'B1_COD', 		'', 	'', 		'C�digo do Produto' })
Aadd(aCampos, { 'TODOS', 'F4_CODIGO',	'', 	'', 		'C�digo da TES' })



//***************************************************************************************
//					IMPOSTO INSS
//***************************************************************************************
If (MV_PAR04 == 1)

	//Parametros
	Aadd(aParame, {'INSS', 'MV_COMIINS', 	'N', 		'Indica se o INSS ser� considerado para pagamento de comiss�o Informar (S) ou (N)' })
	Aadd(aParame, {'INSS', 'MV_FORINSS', 	'C�digo', 	'Fornecedor padr�o para t�tulos de INSS' })
	Aadd(aParame, {'INSS', 'MV_INSS', 		'INSS', 	'Natureza de T�tulos de Pagamento de INSS' })
	Aadd(aParame, {'INSS', 'MV_VLRETIN', 	'25.00', 	'Valor m�nimo para dispensa de reten��o de INSS.' })
	Aadd(aParame, {'INSS', 'MV_LIMINSS', 	'354.08', 	'Valor limite de reten��o para o INSS de pessoa f�sica' })
	Aadd(aParame, {'INSS', 'MV_INSIRF', 	'1', 		'Define se o valor do INSS deve ser abatido da base de calculo do IRRF no Financeiro. �1�, Abate o valor' })
	Aadd(aParame, {'INSS', 'MV_RNDINS', 	'.F.', 		'Informe o crit�rio de arredondamento da INSS. Asop��es v�lidas s�o: .T. arredonda, .F. trunca.titulo principal. 1=Emiss�o,  2=Vencto Real' })
	Aadd(aParame, {'INSS', 'MV_ACMINSS', 	'1', 		'Define se a acumulatividade dos valores de INSS levar�o em conta a data de Emiss�o ou Vencimento do de 1=Emissao,  2=Vencto Real' })
	Aadd(aParame, {'INSS', 'MV_RATINSS', 	'.T.', 		'Indica se o valor do INSS deve ser ou n�o rateado pelo n�mero de t�tulos a pagar em uma nota fiscal' })
	Aadd(aParame, {'INSS', 'MV_MININSS', 	'29.00', 	'Valor m�nimo para reten��o de INSS. Caso o valor de INSS calculado no per�odo seja inferior, o INSS n�o ser� retido.' })

	//Campos INSS
	Aadd(aCampos, { 'INSS', 'A2_RECINSS', 	'S', 	'Calcula INSS', 		'Verificador para c�lculo ou n�o de INSS para titulos deste fornecedor. (S=sim, N=n�o)' })
	Aadd(aCampos, { 'INSS', 'ED_CALCINS', 	'S', 	'Calcula INSS', 		'(S=sim, N=n�o)' })
	Aadd(aCampos, { 'INSS', 'ED_PERCINS', 	'11', 	'Percentual de INSS', 	'Percentual para c�lculo de INSS em titulos com esta natureza' })
	Aadd(aCampos, { 'INSS', 'ED_DEDINSS', 	'1', 	'Dedu��o INSS', 		'Informe "1" para que o valor do INSS seja deduzido do valor da nota/t�tulo quando houver reten��o desta contribui��o. Informe "2" em caso contr�rio' })
	Aadd(aCampos, { 'INSS', 'B1_INSS', 		'S', 	'Calcula INSS', 		'(S=sim, N=n�o)' })

Endif

//***************************************************************************************
//					IMPOSTO PCC
//***************************************************************************************
If (MV_PAR05 == 1)

	//Parametros
	Aadd(aParame, {'PCC', 'MV_TXPIS', 		'0,65', 	'indicam as al�quotas utilizada para PIS via apura��o.' })
	Aadd(aParame, {'PCC', 'MV_TXCOFIN', 	'3', 		'indicam as al�quotas utilizada para COFINS via apura��o.' })
	Aadd(aParame, {'PCC', 'MV_RF10925', 	'26/07/2004','Data de referencia inicial para que os novos procedimentos quanto a reten��o de PIS/COFINS/CSLL seja m aplicados.' })
	Aadd(aParame, {'PCC', 'MV_VL10925', 	'5000', 	'Valor Maximo de pagamentos no per�odo para dispensa da reten��o de PIS/COFINS/CSLL' })
	Aadd(aParame, {'PCC', 'MV_RT10925', 	'1', 		'Modo padr�o de opera��o do sistema quanto a reten��o de PIS/COFINS/CSLL. (1=Calculado p/sistema 2=Efetua Reten��o sempre, 3=N�o efetua reten��o)' })
	Aadd(aParame, {'PCC', 'MV_AC10925', 	'1', 		'Controle de permiss�o para o usu�rio alterar modo de reten��o na janela �calculo de reten��o� 1=Permite altera��o, 2= N�o permite altera��o' })
	Aadd(aParame, {'PCC', 'MV_AB10925', 	'2', 		'Modo de reten��o do PIS,COFINS e CSLL p/ C.Receber 1=Verif.reten��o pelo valor da nota emitida 2=Verif.reten��o p/soma notas emitidas no per�odo' })
	Aadd(aParame, {'PCC', 'MV_MT10925', 	'1', 		'Verificar base de reten��o pelo total ou apenas pelos t�tulos que retiveram impostos 1, Todos 2, Somente os que retiveram Pis etc' })
	Aadd(aParame, {'PCC', 'MV_BX10925', 	'2', 		'Define momento do tratamento da reten��o dos impostos Pis Cofins e Csll 1 Na Baixa ou 2 Na Emiss�o' })
	Aadd(aParame, {'PCC', 'MV_BS10925', 	'2', 		'Indica se o calculo da reten��o ser� sobre a base do PIS/COFINS/CSLL ou pelo valor total da duplicata. 1=Valor da base, 2=Valor total' })
	Aadd(aParame, {'PCC', 'MV_LJ10925', 	'2', 		'Considerar para verifica��o de valores de reten��o os valores de todas as lojas do cliente Op��es: 1 Loja atual ou 2 Todas as lojas' })
	Aadd(aParame, {'PCC', 'MV_TXCSLL', 		'1', 		'indicam as al�quotas utilizada para CSLL via apura��o.' })
	
	//Campos
	Aadd(aCampos, { 'PCC', 'A2_RECCOFI', 	'2', 	'Recolhe Cofins', 		'Informe se o Fornecedor � respons�vel pelo recolhimento da COFINS. Caso n�o seja, o sistema far� a reten��o do tributo. (1=sim, 2=n�o)' })
	Aadd(aCampos, { 'PCC', 'A2_RECPIS', 	'2', 	'Recolhe PIS', 			'Informe se o Fornecedor � respons�vel pelo recolhimento do PIS. Caso n�o seja, o sistema far� a reten��o do tributo. (1=sim, 2=n�o)' })
	Aadd(aCampos, { 'PCC', 'A2_RECCSLL', 	'2', 	'Recolhe Csll', 		'Informe se o Fornecedor � respons�vel pelo recolhimento do CSLL. Caso n�o seja, o sistema far� a reten��o do tributo. (1=sim, 2=n�o)' })
	Aadd(aCampos, { 'PCC', 'ED_CALCPIS', 	'S', 	'Calcula PIS', 			'(S=sim, N=n�o)' })
	Aadd(aCampos, { 'PCC', 'ED_PERCPIS', 	'0,65', 'Percentual de PIS', 	'Percentual para c�lculo de PIS em titulos com esta natureza' })
	Aadd(aCampos, { 'PCC', 'ED_CALCCOF', 	'S', 	'Calcula Cofins', 		'(S=sim, N=n�o)' })
	Aadd(aCampos, { 'PCC', 'ED_PERCCOF', 	'3', 	'Percentual de Cofins', 'Percentual para c�lculo de COFINS em titulos com esta natureza' })
	Aadd(aCampos, { 'PCC', 'ED_CALCCSL', 	'S', 	'Calcula CSLL', 		'(S=sim, N=n�o)' })
	Aadd(aCampos, { 'PCC', 'ED_PERCCSL', 	'1', 	'Percentual de CSLL', 	'Percentual para c�lculo de CSLL em titulos com esta natureza' })
	Aadd(aCampos, { 'PCC', 'B1_PIS', 		'1', 	'Retem PIS', 			'(1=sim, 2=n�o)' })
	Aadd(aCampos, { 'PCC', 'B1_COFINS', 	'1', 	'Retem Cofins', 		'(1=sim, 2=n�o)' })
	Aadd(aCampos, { 'PCC', 'B1_CSLL', 		'1', 	'Retem CSLL', 			'(1=sim, 2=n�o)' })
	Aadd(aCampos, { 'PCC', 'F4_PISCOF', 	'3', 	'Calcula Pis e Cofins', '(1=PIS, 2=COFINS, 3=AMBOS, 4=N�O CONSIDERA)' })
	Aadd(aCampos, { 'PCC', 'F4_PISCRED', 	'3', 	'Credito Pis e Cofins', '(1=Credita, 2=Debita, 3=N�o Calcula, 4=Calcula, 5=Exclus�o de Base)' })


Endif

//***************************************************************************************
//					IMPOSTO IRRF
//***************************************************************************************
If (MV_PAR06 == 1)

	//Parametros
	Aadd(aParame, {'IRRF', 'MV_ALIQIRF', 	'1.5', 		'Al�quota de IRRF para t�tulos c/reten��o na fonte.' })
	Aadd(aParame, {'IRRF', 'MV_VLRETIR', 	'10.00', 	'Valor m�nimo para dispensa de reten��o de IRRF' })
	Aadd(aParame, {'IRRF', 'MV_IMPADT', 	'S', 		'Define utiliza��o da gera��o de IRRF em adiantamento a Pagar' })
	Aadd(aParame, {'IRRF', 'MV_NATNDF', 	'NDF', 		'Natureza do titulo NDF gerado via apura��o de IRRF' })
	Aadd(aParame, {'IRRF', 'MV_PRZIRRF', 	'1', 		'N�mero de dias para pagamento do IRRF apos a emiss�o do titulo' })
	Aadd(aParame, {'IRRF', 'MV_VENCIRF', 	'V', 		'Indica se o t�tulo de IRRF ser� gerado a partir da data de (E) Emiss�o, (V) Vencimento ou (C) Contabiliza��o' })
	Aadd(aParame, {'IRRF', 'MV_RATIRRF', 	'F', 		'Indica se o valor do IRRF deve ser ou n�o rateado pelo numero de t�tulos a pagar em uma nota fiscal de compra (T = Rateia, F = N�o Rateia)' })
	Aadd(aParame, {'IRRF', 'MV_RNDIRF', 	'.F.', 		'Informe o crit�rio de arredondamento da IRRF. Op��es v�lidas s�o: .T. arredonda, .F. trunca.' })
	Aadd(aParame, {'IRRF', 'MV_AGLIMPJ', 	'2', 		'Define forma de verifica��o da base do IRRF1=Apenas Filial Corrente(DEFAULT)' })
	Aadd(aParame, {'IRRF', 'MV_IRMP232', 	'2', 		'Define se a empresa ter� IRRF retido na forma da MP 232. Valores Poss�veis 1-Sim, 2-N�o' })
	Aadd(aParame, {'IRRF', 'MV_VCTIRPF', 	'V', 		'Indica se o titulo de IRRF (Pessoa F�sica) ser� gerado data de (E) Emiss�o, (V) Vencimento ou (C) Contabiliza��o' })
	Aadd(aParame, {'IRRF', 'MV_ACMIRPF', 	'1', 		'Indica qual data ser� considerada para compor a base de c�lculo acumulada de Imposto de Renda para pessoa f�sica.' })
	Aadd(aParame, {'IRRF', 'MV_ACMIRPJ', 	'1', 		'Indica qual data ser� considerada para compor a base de c�lculo acumulada de Imposto de Renda para pessoa jur�dica.' })
	Aadd(aParame, {'IRRF', 'MV_IRF',		'XXX',		'Natureza utilizada para Imposto de Renda' })
	Aadd(aParame, {'IRRF', 'MV_ACMIRRF', 	'1',		'Define se acumula o valor minimo do IRRF. 1 = Acumula (Default) ou 2 = Nao Acumula' })

	//Campos
	Aadd(aCampos, { 'IRRF', 'ED_CALCIRF', 	'S', 	'Calcula IRRF', 		'(S=sim, N=n�o)' })
	Aadd(aCampos, { 'IRRF', 'ED_PERCIRF', 	'1,5', 	'Percentual IRRF', 		'Porcentual de IRRF. Utilizado para base do c�lculo dos titulos de IRRF. Caso n�o seja informado utiliza o opar�metro MV_ALIQIRF.' })
	Aadd(aCampos, { 'IRRF', 'B1_IRRF', 		'S', 	'Calcula IRRF', 		'(S=sim, N=n�o)' })
	Aadd(aCampos, { 'IRRF', 'B1_REDIRRF', 	'0,00', 'Redu��o de IRRF', 		'Caso haja redu��o' })
	Aadd(aCampos, { 'IRRF', 'A2_MINIRF',   	'X',	'Val.Min.Ret IRRF', 	'Dispensa da verifica��o do valor m�nimo para reten��o de IRRF. (1 = Sim) ret�m para qualquer valor, (2 = N�o) Respeita MV_ACMIRRF'})


Endif


//***************************************************************************************
//					IMPOSTO ISS
//***************************************************************************************
If (MV_PAR07 == 1)

	//Parametros
	Aadd(aParame, {'ISS', 'MV_DESCISS', 	'.T.', 		'Informa ao sistema se o ISS devera ser descontado do valor do titulo financeiro caso o cliente for respons�vel pelo recolhimento' })
	Aadd(aParame, {'ISS', 'MV_TPABISS', 	'2', 		'Se par�metro igual a 1 indica se ser� efetuado um desconto na duplicata quando o cliente recolhe ISS se igual a 2 ser� gerado um titulo de abatimento.' })
	Aadd(aParame, {'ISS', 'MV_MRETISS', 	'1', 		'Modo de reten��o do ISS nas aquisi��es de servi�os 1 = na emiss�o do titulo principal,  2 = na baixa do titulo principal' })
	Aadd(aParame, {'ISS', 'MV_VRETISS', 	'0,00', 	'Valor m�nimo para dispensa de reten��o de ISS' })
	Aadd(aParame, {'ISS', 'MV_MODRISS', 	'1', 		'Indica qual o modo de rentencao do ISS ( �1� � por titulo,  �2� � mensal vencimento,  �3� � por base)' })
	Aadd(aParame, {'ISS', 'MV_ALIQISS', 	'5.00', 	'Al�quota do ISS em casos de presta��o de servi�os, usando percentuais definidos pelo munic�pio.' })
	Aadd(aParame, {'ISS', 'MV_DEDISS', 		'2', 		'Define o momento do tratamento da dedu��o do iss (Na baixa ou emiss�o do t�tulo). 1= Na baixa ,  2= Na emiss�o' })
	Aadd(aParame, {'ISS', 'MV_TPNFISS', 	'E', 		'Informe os tipos de documento que deveram realizar o calculo de ISS somente pela Natureza. Utilize: E- para entrada e S-para sa�da.' })
	Aadd(aParame, {'ISS', 'MV_R51CFOP', 	'', 		'Indica quais CFOPs n�o dever�o ser lan�ados no Registro Tipo 51.' })
	
	//Campos ISS
	Aadd(aCampos, { 'ISS', 'A2_RECISS', 	'N', 	'Recolhe ISS ?', 		'Indica se o fornecedor recolhe ou nao o ISS. Se recolher dever� ser preenchido com S, caso contr�rio dever� ser preenchido com N ou branco.' })
	Aadd(aCampos, { 'ISS', 'ED_CALCISS', 	'S', 	'Calcula ISS', 			'(S=sim, N=n�o)' })
	Aadd(aCampos, { 'ISS', 'B1_ALIQISS', 	'', 	'Aliq. de ISS', 		'Informa ao sistema que este produto se refere a Servi�os, utilizando a al�quota para c�lculo de ISS.  (0 = MV_ALIQISS)' })
	Aadd(aCampos, { 'ISS', 'B1_CODISS', 	'', 	'Cod.Serv.ISS', 		'C�digo de Servi�o do ISS, utilizado para discriminar a opera��o perante o munic�pio tributador.' })
	Aadd(aCampos, { 'ISS', 'F4_ISS', 		'S', 	'Calcula ISS', 			'(S=sim, N=n�o)' })
	Aadd(aCampos, { 'ISS', 'F4_LFISS', 		'T', 	'Livro Fiscal ISS', 	'Livro Fiscal ISS. "T" para ISS tributado, "I" para ISS isento, "O" para ISS outras, "N" n�o lan�ar no Livro Fiscal.' })
		

Endif


//Consulta dos Parametros
For i:=1 to Len(aParame)

	//Trata valor logico
	xVal := SuperGetMv(aPArame[i][2])
	IF (VALTYPE(xVal) == 'L')
		IF (xVal)
			cVal := '.T.'
		Else
			cVal := '.F.'
		Endif
	Else
		cVal := xVal
	Endif
	//Grava no retorno
	Aadd(aDados, {aPArame[i][1], aPArame[i][2], cVal, aPArame[i][3], '', aPArame[i][4]  })
	xVal := Nil
Next


//Consulta dos Campos
For c:= 1 to Len(aCampos)
	cQuery += aCampos[c][2] + ", "
Next

cQuery += " A2_LOJA " 
cQuery += " FROM " + retSqlName("SA2") + " AS SA2"
cQuery += "	INNER JOIN " + retSqlName("SB1") + " AS SB1 ON SB1.D_E_L_E_T_ <> '*' AND B1_COD = '" + MV_PAR02 + "'"
cQuery += "	INNER JOIN " + retSqlName("SF4") + " AS SF4 ON SF4.D_E_L_E_T_ <> '*' AND F4_CODIGO = '" + MV_PAR03 + "'"
cQuery += "	LEFT JOIN  " + retSqlName("SED") + " AS SED ON SED.D_E_L_E_T_ <> '*' AND ED_CODIGO = A2_NATUREZ "
cQuery += " WHERE SA2.D_E_L_E_T_ <> '*' AND A2_COD = '" + MV_PAR01 + "'" 


aRet	:= u_QryArr(changeQuery(cQuery))

For nQ := 1 to Len(aRet)
	For nl := 1 to Len(aCampos)
		
		Aadd(aDados, {	aCampos[nl][1],;
						aCampos[nl][2],;
						aRet[nQ][nl],;
						aCampos[nl][3],;
						aCampos[nl][4],;
						aCampos[nl][5] })
	Next
Next 

//Gera o Relat�rio
U_GD2Excel(aCabec,aDados) 
Return



//*********************************************************
//		CRIA AS PERGUNTAS DO RELATORIO
//*********************************************************

Static Function criaPerg(cPerg)

     
    cValid   := ""
    cF3      := ""
    cPicture := ""
    cDef01   := ""
    cDef02   := ""
    cDef03   := ""
    cDef04   := ""
    cDef05   := ""

	u_zPutSX1(cPerg, "01", "Fornecedor",       	"MV_PAR01", "MV_CH0", "C", 6,	0,	"G", cValid,       "SA2", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o c�digo do Fornecedor")
	u_zPutSX1(cPerg, "02", "Produto",       	"MV_PAR02", "MV_CH1", "C", 15,	0,	"G", cValid,       "SB1", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o c�digo do Produto")
	u_zPutSX1(cPerg, "03", "Tipo de Entrada",	"MV_PAR03", "MV_CH2", "C", 3,	0,	"G", cValid,       "SF4", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o c�digo da TES")                                                                   
	u_zPutSX1(cPerg, "04", "INSS ?",       		"MV_PAR04", "MV_CH3", "C", 1,	0,	"C", cValid,       cF3,   cPicture,        "1=SIM",	"2=NAO",       cDef03,        cDef04,    cDef05, "V�lida INSS?")
	u_zPutSX1(cPerg, "05", "PCC ?",       		"MV_PAR05", "MV_CH4", "C", 1,	0,	"C", cValid,       cF3,   cPicture,        "1=SIM",	"2=NAO",       cDef03,        cDef04,    cDef05, "V�lida PIS, COFINS, CSLL?")
	u_zPutSX1(cPerg, "06", "IRRF ?",       		"MV_PAR06", "MV_CH5", "C", 1,	0,	"C", cValid,       cF3,   cPicture,        "1=SIM",	"2=NAO",       cDef03,        cDef04,    cDef05, "V�lida IRRF?")
	u_zPutSX1(cPerg, "07", "ISS ?",       		"MV_PAR07", "MV_CH6", "C", 1,	0,	"C", cValid,       cF3,   cPicture,        "1=SIM",	"2=NAO",       cDef03,        cDef04,    cDef05, "V�lida ISS?")


Return

