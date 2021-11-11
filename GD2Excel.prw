# include "Protheus.ch"
 

/*/
Funcao: GDToExcel
Autor: Marinaldo de Jesus 
Data: 01/06/2013
Descricao: Mostrar os Dados no Excel
Sintaxe: StaticCall(NDJLIB001,GDToExcel,aHeader,aCols,cWorkSheet,cTable,lTotalize,lPicture)
/*/

User Function GD2Excel(aHeader,aCols,cWorkSheet,cTable,lTotalize,lPicture)


Local oFWMSExcel := FWMSExcel():New()

Local oMsExcel

Local aCells

Local cType
Local cColumn

Local cFile
Local cFileTMP

Local cPicture

Local lTotal

Local nRow
Local nRows
Local nField
Local nFields

Local nAlign
Local nFormat

Local uCell

DEFAULT cWorkSheet := "Plan1"
DEFAULT cTable := "RELATORIO TOTVS"
DEFAULT lTotalize := .T.
DEFAULT lPicture := .F.

BEGIN SEQUENCE

oFWMSExcel:AddworkSheet(cWorkSheet)
oFWMSExcel:AddTable(cWorkSheet,cTable)
        
/*       
cType:
	C = Caracter
	D = Data
	N = Numerico (Totaliza)
	M = Moeda	 (Totaliza)
*/                         

nFields := Len( aHeader )
For nField := 1 To nFields
cType := aHeader[nField][2]
nAlign := IF(cType=="C",1,IF(cType=="N",3,2))
nFormat := IF(cType=="D",4,IF(cType=="M",3,IF(cType=="N",2,1))) 
cColumn := aHeader[nField][1]
lTotal := ( lTotalize .and. (cType == "M" .or. cType == "N") )
oFWMSExcel:AddColumn(@cWorkSheet,@cTable,@cColumn,@nAlign,@nFormat,@lTotal)
Next nField

aCells := Array(nFields)

nRows := Len( aCols )
For nRow := 1 To nRows
For nField := 1 To nFields
uCell := aCols[nRow][nField]
IF ( lPicture )
cPicture := aHeader[nField][__AHEADER_PICTURE__]
IF .NOT.( Empty(cPicture) )
uCell := Transform(uCell,cPicture)
EndIF
EndIF
aCells[nField] := uCell
Next nField
oFWMSExcel:AddRow(@cWorkSheet,@cTable,aClone(aCells))
Next nRow

oFWMSExcel:Activate()

cFile := ( CriaTrab( NIL, .F. ) + ".xml" )

While File( cFile )
cFile := ( CriaTrab( NIL, .F. ) + ".xml" )
End While

oFWMSExcel:GetXMLFile( cFile )
oFWMSExcel:DeActivate()

IF .NOT.( File( cFile ) )
cFile := ""
BREAK
EndIF

cFileTMP := ( GetTempPath() + cFile )
IF .NOT.( __CopyFile( cFile , cFileTMP ) )
fErase( cFile )
cFile := ""
BREAK
EndIF

fErase( cFile )

cFile := cFileTMP

IF .NOT.( File( cFile ) )
cFile := ""
BREAK
EndIF

IF .NOT.( ApOleClient("MsExcel") )
BREAK
EndIF

oMsExcel := MsExcel():New()
oMsExcel:WorkBooks:Open( cFile )
oMsExcel:SetVisible( .T. )
oMsExcel := oMsExcel:Destroy()

END SEQUENCE

oFWMSExcel := FreeObj( oFWMSExcel )

Return( cFile )