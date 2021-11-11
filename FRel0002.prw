#include "Protheus.ch"


/*/{Protheus.doc} Frel0002.prw
@type function
@systemOper Windows
@author Wendell Tavares
@since 01/08/2019
@version 1.0
@description Relatorio de parametros para geração dos impostos de servico
@link https://github.com/wendellgt

/*/


User Function Frel0002()

	cPergunte := "R0002"
	criaPerg(cPergunte)
	 
	pergunte(cPergunte,.T.)

	Processa( {|| U_ERel0002() }, "Aguarde...", "Gerando Relatório...",.F.)

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
	MsgAlert("Microsoft Excel não instalado!") 
	Return 
EndIf

/*       
cType:
	C = Caracter
	D = Data
	N = Numerico (Totaliza)
	M = Moeda	 (Totaliza)
*/  
//Cabeçalho da planilha
AADD(aCabec, {"TIPO"				,"C"})
AADD(aCabec, {"PARAMETRO"			,"C"})
AADD(aCabec, {"CONTEUDO"			,"C"})
AADD(aCabec, {"ESPERADO"			,"C"})
AADD(aCabec, {"DESCRICAO"			,"C"})
AADD(aCabec, {"OBSERVACAO"			,"C"}) 


//Campos obrigatorios
Aadd(aCampos, { 'TODOS', 'A2_COD', 		'', 	'', 		'Código do Fornecedor' })
Aadd(aCampos, { 'TODOS', 'A2_NATUREZ', 	'', 	'', 		'Código da Natureza' })
Aadd(aCampos, { 'TODOS', 'B1_COD', 		'', 	'', 		'Código do Produto' })
Aadd(aCampos, { 'TODOS', 'F4_CODIGO',	'', 	'', 		'Código da TES' })



//***************************************************************************************
//					IMPOSTO INSS
//***************************************************************************************
If (MV_PAR04 == 1)

	//Parametros
	Aadd(aParame, {'INSS', 'MV_COMIINS', 	'N', 		'Indica se o INSS será considerado para pagamento de comissão Informar (S) ou (N)' })
	Aadd(aParame, {'INSS', 'MV_FORINSS', 	'Código', 	'Fornecedor padrão para títulos de INSS' })
	Aadd(aParame, {'INSS', 'MV_INSS', 		'INSS', 	'Natureza de Títulos de Pagamento de INSS' })
	Aadd(aParame, {'INSS', 'MV_VLRETIN', 	'25.00', 	'Valor mínimo para dispensa de retenção de INSS.' })
	Aadd(aParame, {'INSS', 'MV_LIMINSS', 	'354.08', 	'Valor limite de retenção para o INSS de pessoa física' })
	Aadd(aParame, {'INSS', 'MV_INSIRF', 	'1', 		'Define se o valor do INSS deve ser abatido da base de calculo do IRRF no Financeiro. “1”, Abate o valor' })
	Aadd(aParame, {'INSS', 'MV_RNDINS', 	'.F.', 		'Informe o critério de arredondamento da INSS. Asopþ§es válidas são: .T. arredonda, .F. trunca.titulo principal. 1=Emissão,  2=Vencto Real' })
	Aadd(aParame, {'INSS', 'MV_ACMINSS', 	'1', 		'Define se a acumulatividade dos valores de INSS levarão em conta a data de Emissão ou Vencimento do de 1=Emissao,  2=Vencto Real' })
	Aadd(aParame, {'INSS', 'MV_RATINSS', 	'.T.', 		'Indica se o valor do INSS deve ser ou não rateado pelo n•mero de títulos a pagar em uma nota fiscal' })
	Aadd(aParame, {'INSS', 'MV_MININSS', 	'29.00', 	'Valor mínimo para retenção de INSS. Caso o valor de INSS calculado no período seja inferior, o INSS não será retido.' })

	//Campos INSS
	Aadd(aCampos, { 'INSS', 'A2_RECINSS', 	'S', 	'Calcula INSS', 		'Verificador para cálculo ou não de INSS para titulos deste fornecedor. (S=sim, N=não)' })
	Aadd(aCampos, { 'INSS', 'ED_CALCINS', 	'S', 	'Calcula INSS', 		'(S=sim, N=não)' })
	Aadd(aCampos, { 'INSS', 'ED_PERCINS', 	'11', 	'Percentual de INSS', 	'Percentual para cálculo de INSS em titulos com esta natureza' })
	Aadd(aCampos, { 'INSS', 'ED_DEDINSS', 	'1', 	'Dedução INSS', 		'Informe "1" para que o valor do INSS seja deduzido do valor da nota/título quando houver retenção desta contribuição. Informe "2" em caso contrário' })
	Aadd(aCampos, { 'INSS', 'B1_INSS', 		'S', 	'Calcula INSS', 		'(S=sim, N=não)' })

Endif

//***************************************************************************************
//					IMPOSTO PCC
//***************************************************************************************
If (MV_PAR05 == 1)

	//Parametros
	Aadd(aParame, {'PCC', 'MV_TXPIS', 		'0,65', 	'indicam as alíquotas utilizada para PIS via apuração.' })
	Aadd(aParame, {'PCC', 'MV_TXCOFIN', 	'3', 		'indicam as alíquotas utilizada para COFINS via apuração.' })
	Aadd(aParame, {'PCC', 'MV_RF10925', 	'26/07/2004','Data de referencia inicial para que os novos procedimentos quanto a retenção de PIS/COFINS/CSLL seja m aplicados.' })
	Aadd(aParame, {'PCC', 'MV_VL10925', 	'5000', 	'Valor Maximo de pagamentos no período para dispensa da retenção de PIS/COFINS/CSLL' })
	Aadd(aParame, {'PCC', 'MV_RT10925', 	'1', 		'Modo padrão de operação do sistema quanto a retenção de PIS/COFINS/CSLL. (1=Calculado p/sistema 2=Efetua Retenção sempre, 3=Não efetua retenção)' })
	Aadd(aParame, {'PCC', 'MV_AC10925', 	'1', 		'Controle de permissão para o usuário alterar modo de retenção na janela “calculo de retenção” 1=Permite alteração, 2= Não permite alteração' })
	Aadd(aParame, {'PCC', 'MV_AB10925', 	'2', 		'Modo de retenção do PIS,COFINS e CSLL p/ C.Receber 1=Verif.retenção pelo valor da nota emitida 2=Verif.retenção p/soma notas emitidas no período' })
	Aadd(aParame, {'PCC', 'MV_MT10925', 	'1', 		'Verificar base de retenção pelo total ou apenas pelos títulos que retiveram impostos 1, Todos 2, Somente os que retiveram Pis etc' })
	Aadd(aParame, {'PCC', 'MV_BX10925', 	'2', 		'Define momento do tratamento da retenção dos impostos Pis Cofins e Csll 1 Na Baixa ou 2 Na Emissão' })
	Aadd(aParame, {'PCC', 'MV_BS10925', 	'2', 		'Indica se o calculo da retenção será sobre a base do PIS/COFINS/CSLL ou pelo valor total da duplicata. 1=Valor da base, 2=Valor total' })
	Aadd(aParame, {'PCC', 'MV_LJ10925', 	'2', 		'Considerar para verificação de valores de retenção os valores de todas as lojas do cliente Opções: 1 Loja atual ou 2 Todas as lojas' })
	Aadd(aParame, {'PCC', 'MV_TXCSLL', 		'1', 		'indicam as alíquotas utilizada para CSLL via apuração.' })
	
	//Campos
	Aadd(aCampos, { 'PCC', 'A2_RECCOFI', 	'2', 	'Recolhe Cofins', 		'Informe se o Fornecedor é responsável pelo recolhimento da COFINS. Caso não seja, o sistema fará a retenção do tributo. (1=sim, 2=não)' })
	Aadd(aCampos, { 'PCC', 'A2_RECPIS', 	'2', 	'Recolhe PIS', 			'Informe se o Fornecedor é responsável pelo recolhimento do PIS. Caso não seja, o sistema fará a retenção do tributo. (1=sim, 2=não)' })
	Aadd(aCampos, { 'PCC', 'A2_RECCSLL', 	'2', 	'Recolhe Csll', 		'Informe se o Fornecedor é responsável pelo recolhimento do CSLL. Caso não seja, o sistema fará a retenção do tributo. (1=sim, 2=não)' })
	Aadd(aCampos, { 'PCC', 'ED_CALCPIS', 	'S', 	'Calcula PIS', 			'(S=sim, N=não)' })
	Aadd(aCampos, { 'PCC', 'ED_PERCPIS', 	'0,65', 'Percentual de PIS', 	'Percentual para cálculo de PIS em titulos com esta natureza' })
	Aadd(aCampos, { 'PCC', 'ED_CALCCOF', 	'S', 	'Calcula Cofins', 		'(S=sim, N=não)' })
	Aadd(aCampos, { 'PCC', 'ED_PERCCOF', 	'3', 	'Percentual de Cofins', 'Percentual para cálculo de COFINS em titulos com esta natureza' })
	Aadd(aCampos, { 'PCC', 'ED_CALCCSL', 	'S', 	'Calcula CSLL', 		'(S=sim, N=não)' })
	Aadd(aCampos, { 'PCC', 'ED_PERCCSL', 	'1', 	'Percentual de CSLL', 	'Percentual para cálculo de CSLL em titulos com esta natureza' })
	Aadd(aCampos, { 'PCC', 'B1_PIS', 		'1', 	'Retem PIS', 			'(1=sim, 2=não)' })
	Aadd(aCampos, { 'PCC', 'B1_COFINS', 	'1', 	'Retem Cofins', 		'(1=sim, 2=não)' })
	Aadd(aCampos, { 'PCC', 'B1_CSLL', 		'1', 	'Retem CSLL', 			'(1=sim, 2=não)' })
	Aadd(aCampos, { 'PCC', 'F4_PISCOF', 	'3', 	'Calcula Pis e Cofins', '(1=PIS, 2=COFINS, 3=AMBOS, 4=NÃO CONSIDERA)' })
	Aadd(aCampos, { 'PCC', 'F4_PISCRED', 	'3', 	'Credito Pis e Cofins', '(1=Credita, 2=Debita, 3=Não Calcula, 4=Calcula, 5=Exclusão de Base)' })


Endif

//***************************************************************************************
//					IMPOSTO IRRF
//***************************************************************************************
If (MV_PAR06 == 1)

	//Parametros
	Aadd(aParame, {'IRRF', 'MV_ALIQIRF', 	'1.5', 		'Alíquota de IRRF para títulos c/retenção na fonte.' })
	Aadd(aParame, {'IRRF', 'MV_VLRETIR', 	'10.00', 	'Valor mínimo para dispensa de retenção de IRRF' })
	Aadd(aParame, {'IRRF', 'MV_IMPADT', 	'S', 		'Define utilização da geração de IRRF em adiantamento a Pagar' })
	Aadd(aParame, {'IRRF', 'MV_NATNDF', 	'NDF', 		'Natureza do titulo NDF gerado via apuração de IRRF' })
	Aadd(aParame, {'IRRF', 'MV_PRZIRRF', 	'1', 		'Número de dias para pagamento do IRRF apos a emissão do titulo' })
	Aadd(aParame, {'IRRF', 'MV_VENCIRF', 	'V', 		'Indica se o título de IRRF será gerado a partir da data de (E) Emissão, (V) Vencimento ou (C) Contabilização' })
	Aadd(aParame, {'IRRF', 'MV_RATIRRF', 	'F', 		'Indica se o valor do IRRF deve ser ou não rateado pelo numero de títulos a pagar em uma nota fiscal de compra (T = Rateia, F = Não Rateia)' })
	Aadd(aParame, {'IRRF', 'MV_RNDIRF', 	'.F.', 		'Informe o critério de arredondamento da IRRF. Opções válidas são: .T. arredonda, .F. trunca.' })
	Aadd(aParame, {'IRRF', 'MV_AGLIMPJ', 	'2', 		'Define forma de verificação da base do IRRF1=Apenas Filial Corrente(DEFAULT)' })
	Aadd(aParame, {'IRRF', 'MV_IRMP232', 	'2', 		'Define se a empresa terá IRRF retido na forma da MP 232. Valores Possíveis 1-Sim, 2-Não' })
	Aadd(aParame, {'IRRF', 'MV_VCTIRPF', 	'V', 		'Indica se o titulo de IRRF (Pessoa Física) será gerado data de (E) Emissão, (V) Vencimento ou (C) Contabilização' })
	Aadd(aParame, {'IRRF', 'MV_ACMIRPF', 	'1', 		'Indica qual data será considerada para compor a base de cálculo acumulada de Imposto de Renda para pessoa física.' })
	Aadd(aParame, {'IRRF', 'MV_ACMIRPJ', 	'1', 		'Indica qual data será considerada para compor a base de cálculo acumulada de Imposto de Renda para pessoa jurídica.' })
	Aadd(aParame, {'IRRF', 'MV_IRF',		'XXX',		'Natureza utilizada para Imposto de Renda' })
	Aadd(aParame, {'IRRF', 'MV_ACMIRRF', 	'1',		'Define se acumula o valor minimo do IRRF. 1 = Acumula (Default) ou 2 = Nao Acumula' })

	//Campos
	Aadd(aCampos, { 'IRRF', 'ED_CALCIRF', 	'S', 	'Calcula IRRF', 		'(S=sim, N=não)' })
	Aadd(aCampos, { 'IRRF', 'ED_PERCIRF', 	'1,5', 	'Percentual IRRF', 		'Porcentual de IRRF. Utilizado para base do cálculo dos titulos de IRRF. Caso não seja informado utiliza o oparâmetro MV_ALIQIRF.' })
	Aadd(aCampos, { 'IRRF', 'B1_IRRF', 		'S', 	'Calcula IRRF', 		'(S=sim, N=não)' })
	Aadd(aCampos, { 'IRRF', 'B1_REDIRRF', 	'0,00', 'Redução de IRRF', 		'Caso haja redução' })
	Aadd(aCampos, { 'IRRF', 'A2_MINIRF',   	'X',	'Val.Min.Ret IRRF', 	'Dispensa da verificação do valor mínimo para retenção de IRRF. (1 = Sim) retém para qualquer valor, (2 = Não) Respeita MV_ACMIRRF'})


Endif


//***************************************************************************************
//					IMPOSTO ISS
//***************************************************************************************
If (MV_PAR07 == 1)

	//Parametros
	Aadd(aParame, {'ISS', 'MV_DESCISS', 	'.T.', 		'Informa ao sistema se o ISS devera ser descontado do valor do titulo financeiro caso o cliente for responsável pelo recolhimento' })
	Aadd(aParame, {'ISS', 'MV_TPABISS', 	'2', 		'Se parâmetro igual a 1 indica se será efetuado um desconto na duplicata quando o cliente recolhe ISS se igual a 2 será gerado um titulo de abatimento.' })
	Aadd(aParame, {'ISS', 'MV_MRETISS', 	'1', 		'Modo de retenção do ISS nas aquisições de serviços 1 = na emissão do titulo principal,  2 = na baixa do titulo principal' })
	Aadd(aParame, {'ISS', 'MV_VRETISS', 	'0,00', 	'Valor mínimo para dispensa de retenção de ISS' })
	Aadd(aParame, {'ISS', 'MV_MODRISS', 	'1', 		'Indica qual o modo de rentencao do ISS ( “1” – por titulo,  “2” – mensal vencimento,  “3” – por base)' })
	Aadd(aParame, {'ISS', 'MV_ALIQISS', 	'5.00', 	'Alíquota do ISS em casos de prestação de serviços, usando percentuais definidos pelo município.' })
	Aadd(aParame, {'ISS', 'MV_DEDISS', 		'2', 		'Define o momento do tratamento da dedução do iss (Na baixa ou emissão do título). 1= Na baixa ,  2= Na emissão' })
	Aadd(aParame, {'ISS', 'MV_TPNFISS', 	'E', 		'Informe os tipos de documento que deveram realizar o calculo de ISS somente pela Natureza. Utilize: E- para entrada e S-para saída.' })
	Aadd(aParame, {'ISS', 'MV_R51CFOP', 	'', 		'Indica quais CFOPs não deverão ser lançados no Registro Tipo 51.' })
	
	//Campos ISS
	Aadd(aCampos, { 'ISS', 'A2_RECISS', 	'N', 	'Recolhe ISS ?', 		'Indica se o fornecedor recolhe ou nao o ISS. Se recolher deverá ser preenchido com S, caso contrário deverá ser preenchido com N ou branco.' })
	Aadd(aCampos, { 'ISS', 'ED_CALCISS', 	'S', 	'Calcula ISS', 			'(S=sim, N=não)' })
	Aadd(aCampos, { 'ISS', 'B1_ALIQISS', 	'', 	'Aliq. de ISS', 		'Informa ao sistema que este produto se refere a Serviços, utilizando a alíquota para cálculo de ISS.  (0 = MV_ALIQISS)' })
	Aadd(aCampos, { 'ISS', 'B1_CODISS', 	'', 	'Cod.Serv.ISS', 		'Código de Serviço do ISS, utilizado para discriminar a operação perante o município tributador.' })
	Aadd(aCampos, { 'ISS', 'F4_ISS', 		'S', 	'Calcula ISS', 			'(S=sim, N=não)' })
	Aadd(aCampos, { 'ISS', 'F4_LFISS', 		'T', 	'Livro Fiscal ISS', 	'Livro Fiscal ISS. "T" para ISS tributado, "I" para ISS isento, "O" para ISS outras, "N" não lançar no Livro Fiscal.' })
		

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

//Gera o Relatório
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

	u_zPutSX1(cPerg, "01", "Fornecedor",       	"MV_PAR01", "MV_CH0", "C", 6,	0,	"G", cValid,       "SA2", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o código do Fornecedor")
	u_zPutSX1(cPerg, "02", "Produto",       	"MV_PAR02", "MV_CH1", "C", 15,	0,	"G", cValid,       "SB1", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o código do Produto")
	u_zPutSX1(cPerg, "03", "Tipo de Entrada",	"MV_PAR03", "MV_CH2", "C", 3,	0,	"G", cValid,       "SF4", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o código da TES")                                                                   
	u_zPutSX1(cPerg, "04", "INSS ?",       		"MV_PAR04", "MV_CH3", "C", 1,	0,	"C", cValid,       cF3,   cPicture,        "1=SIM",	"2=NAO",       cDef03,        cDef04,    cDef05, "Válida INSS?")
	u_zPutSX1(cPerg, "05", "PCC ?",       		"MV_PAR05", "MV_CH4", "C", 1,	0,	"C", cValid,       cF3,   cPicture,        "1=SIM",	"2=NAO",       cDef03,        cDef04,    cDef05, "Válida PIS, COFINS, CSLL?")
	u_zPutSX1(cPerg, "06", "IRRF ?",       		"MV_PAR06", "MV_CH5", "C", 1,	0,	"C", cValid,       cF3,   cPicture,        "1=SIM",	"2=NAO",       cDef03,        cDef04,    cDef05, "Válida IRRF?")
	u_zPutSX1(cPerg, "07", "ISS ?",       		"MV_PAR07", "MV_CH6", "C", 1,	0,	"C", cValid,       cF3,   cPicture,        "1=SIM",	"2=NAO",       cDef03,        cDef04,    cDef05, "Válida ISS?")


Return

