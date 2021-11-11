#include "Protheus.ch"


/*/{Protheus.doc} FRel0001.PRW
@type function
@systemOper Windows
@author Wendell Tavares
@since 31/07/2019
@version 1.0
@description Relatório de Usuarios do Sistema contendo dados e ultimo acesso, utiliza a biblioteca GD2Excel para exportar o relatório
@table sigapss 
@link https://github.com/wendellgt
/*/

 
User Function FRel0001()

	Processa( {|| U_ERel0001() }, "Aguarde...", "Gerando Relatório...",.F.)

Return


User Function ERel0001()
 
Local aCabec := {} 
Local aDados := {}
Local aRet := {}
Local nl
Local aAcesso := {}
Local cBloq

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
AADD(aCabec, {"CODIGO"				,"C"})
AADD(aCabec, {"USUARIO"				,"C"})
AADD(aCabec, {"NOME"				,"C"})
AADD(aCabec, {"E-MAIL"				,"C"})
AADD(aCabec, {"ACESSOS"				,"N"}) 
AADD(aCabec, {"BLOQUEADO"			,"C"})
AADD(aCabec, {"DT ULT ACESSO"		,"D"})
AADD(aCabec, {"HORA"				,"C"})
AADD(aCabec, {"COMPUTADOR"			,"C"})


//Carrega todos os usuário
aRet := AllUsers() 

For nl := 1 to Len(aRet)
	//Verifica o ultimo acesso do usuario	       
	aAcesso := FWUsrUltLog(aRet[nl][1][1])
	
	IIF(aRet[nl][1][17], cBloq := "1", cBloq := "0")
	
	Aadd(aDados, {	aRet[nl][1][1],;
							aRet[nl][1][2],;
							aRet[nl][1][4],;
							aRet[nl][1][14],;
							aRet[nl][1][15],;
							cBloq,;
							aAcesso[1],;
							aAcesso[2],;
							aAcesso[4] })
Next
 
U_GD2Excel(aCabec,aDados) 
Return
