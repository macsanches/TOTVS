#Include "TOTVS.ch"    
#include "Topconn.ch"  
#include "RESTFUL.CH"

WSRESTFUL GetVisita DESCRIPTION "API REST Protheus para consulta de visitas" FORMAT APPLICATION_JSON
WSDATA IdVisita	       AS CHARACTER  OPTIONAL

WSMETHOD GET GetVisita;
	DESCRIPTION "API utilizada para efetuar a consulta de visitas";
	WSSYNTAX "/rest/visitas/?{IdVisita}";
	PATH "/rest/visitas/";
	TTALK "ConsultaVisita";
	PRODUCES APPLICATION_JSON
END WSRESTFUL


WSMETHOD GET GetVisita HEADERPARAM IdVisita WSSERVICE GetVisita
	Local oResponse		:=	Nil
	Local aResponse		:=	{}
	Local oVisita		
	Local lRet			:=	.T.
	Local cVisita         := ""
	Local cQuery        := ""
	Local cAliasSPY     := ""

	if ( ValType( self:IdVisita  ) == "C" .and. !Empty( self:IdVisita  ) )
   		cVisita := self:IdVisita
   		cVisita := cVisita
	ENDIF

	cQuery := "SELECT * "
	cQuery += "	FROM " + RetSqlName("SPY") + " SPY " + Chr(10) + Chr (13) 
	cQuery += "WHERE SPY.D_E_L_E_T_='' "+ Chr(10) + Chr(13)
	cQuery += "AND PY_FILIAL ='"+ xFilial("SM0") + "' AND PY_NUMERO='"+ cVisita + "' "+ Chr(10) + Chr(13)
	cQuery += "ORDER BY PY_NUMERO"

	conout("testapi: "+ cQuery)

	While .T.
    	cAliasSPY := GetNextAlias()
    	If !TCCanOpen(cAliasSPY) .And. Select(cAliasSPY) == 0
        	Exit
    	EndIf
	EndDo

	dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSPY,.F.,.T.) 
	DbSelectArea(cAliasSPY)
	
	(cAliasSPY)->(DbGoTop())

	IF (cAliasSPY)->(EOF())
   		(cAliasSPY)->(DbcloseArea())
   		oResponse := JsonObject():New()
   			oResponse["consultarResultado"]	:= {}
   			self:SetResponse( oResponse:ToJson() )
   		FreeObj( oResponse )
   		oResponse := Nil
   		Return( lRet )
	ELSE
   		While !(cAliasSPY)->(EOF())
        	oVisita  := nil
        	oVisita  := JsonObject():New()
        	oVisita["Filial"]   	:= ALLTRIM((cAliasSPY)->PY_FILIAL)
        	oVisita["Numero"]   	:= (cAliasSPY)->PY_NUMERO
        	oVisita["nomeVis"]  	:= ALLTRIM((cAliasSPY)->PY_NOME)
        	oVisita["contato"]  	:= ALLTRIM((cAliasSPY)->PY_CONTATO)
        	oVisita["cc"]       	:= ALLTRIM((cAliasSPY)->PY_DESCCC)
        	oVisita["empresa"]  	:= (cAliasSPY)->PY_NOMEMP
        	oVisita["dtvisita"]		:= alltrim((cAliasSPY)->PY_DTVISIT)
        	oVisita["entrada"]  	:= (cAliasSPY)->PY_DATAE
			oVisita["saida"]    	:= (cAliasSPY)->PY_DATAS

			if PY_TIPOVIS == "1"
				oVisita["TipoVis"]	:= "Negocios"
			else
				oVisita["TipoVis"]	:= "Particular"
			EndIf

			if PY_CLASSIF == "1"
				oVisita["TipoVis"] := "Agendada"
			else
				oVisita["TipoVis"]:= "Nao agendada"
			EndIf

        	aadd(aResponse,oVisita)

        	(cAliasSPY)->(dbskip())
    	Enddo
    	(cAliasSPY)->(DbcloseArea())

    	oResponse := JsonObject():New()
    	oResponse["consultarResultado"]	:= aResponse
    	self:SetResponse(EncodeUTF8(oResponse:ToJson()))
	ENDIF

	FreeObj(oResponse)
	oResponse := Nil
Return(lRet)
