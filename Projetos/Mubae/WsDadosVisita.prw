#Include "TOTVS.ch"    
#include "Topconn.ch"  
#include "RESTFUL.CH"

WSRESTFUL visitas DESCRIPTION "API REST Protheus para consulta de visitas" FORMAT APPLICATION_JSON
WSDATA IdVisita	       AS STRING

WSMETHOD GET DESCRIPTION "API utilizada para efetuar a consulta de visitas" WSSYNTAX "/visitas/?{IdVisita}"
WSMETHOD POST DESCRIPTION "Sincronizacao de Dados via POST" WSSYNTAX "/visitas/?"

END WSRESTFUL


WSMETHOD GET HEADERPARAM IdVisita WSSERVICE visitas
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

	cQuery := "SELECT * FROM SPY010 SPY " 
	cQuery += "INNER JOIN SPW010 SPW ON PY_VISITA = PW_VISITA "
	cQuery += "WHERE SPY.D_E_L_E_T_ = '' AND SPW.D_E_L_E_T_ = '' "
	cQuery += "AND PY_NUMERO = '"+ cVisita + "' "
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
        	oVisita["Filial"]   	:= Alltrim((cAliasSPY)->PY_FILIAL)
        	oVisita["Numero"]   	:= (cAliasSPY)->PY_NUMERO
        	oVisita["nomeVis"]  	:= Alltrim((cAliasSPY)->PW_NOMFULL)
        	oVisita["cpf"]  		:= Alltrim((cAliasSPY)->PW_CPF)
        	oVisita["empresa"]  	:= Alltrim((cAliasSPY)->PY_NOMEMP)
        	oVisita["dtvisita"]		:= Alltrim((cAliasSPY)->PY_DTVISIT)
        	oVisita["entrada"]  	:= (cAliasSPY)->PY_DATAE
			oVisita["saida"]    	:= (cAliasSPY)->PY_DATAS

			If PY_TIPOVIS == "1"
				oVisita["tipovis"]	:= "Negocios"
			ElseIf PY_TIPOVIS == "2"
				oVisita["tipovis"]	:= "Embarque"
			ElseIf PY_TIPOVIS == "3"
				oVisita["tipovis"]	:= "Entrega"
			ElseIf PY_TIPOVIS == "4"
				oVisita["tipovis"]	:= "Emb/Ent"
			ElseIf PY_TIPOVIS == "5"
				oVisita["tipovis"]	:= "Ent.Geral"
			EndIf

			if PY_CLASSIF == "1"
				oVisita["classif"] := "Agendada"
			else
				oVisita["classif"]:= "Nao agendada"
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


WSMETHOD POST WSRECEIVE RECEIVE WSSERVICE visitas

	Local lRet      := .T.         // Recebe o Retorno 
	Local cBody     := ''          // Recebe o conteudo do Rest
	Local oJson     := NIL         // Recebe o JSON de Entrada
	Local oJsonRet  := NIL         // Recebe o JSON de Saida
	Local cErr 		:= Space(0)
	Local cJson		:= Space(0)
	Local i 		:= 0
	Local ret

	// Pega o conteudo JSON da transação Rest
	cBody := Self:GetContent() 
	::SetContentType("application/json")

	oJson := JsonObject():new()
	ret := oJson:FromJson(cBody)

	conout("conteudo: " + cBody)

	If ValType(ret) == "C"
		SetRestFault(400,cErr)
		Return(.F.)
	EndIf

	//Deserializando o JSON
	cJson := oJson:toJson()

	//Exibe o conteúdo do JSON deserializado no console
	ConOut("", "JSON deserializado:", cJson)
  
	aRet := {}
	aAtuVis := oJson:GetJsonObject("Assinatura")

	For i := 1 to len(aAtuVis)

		cNumVis := aAtuVis[i]['CodVis']
		cAssinatura := aAtuVis[i]['Assinatura']
		
		dbSelectArea("SPY")
		dbSetOrder(4)

		If SPY->(dbSeek(xFilial() + cNumVis))
			RecLock("SPY",.F.)
				SPY->PY_XXASSDG := cAssinatura
			MsUnlock()

			Aadd(aRet,JsonObject():new())
			nPos := Len(aRet)
			aRet[nPos]['Visita'] := cNumVis
			aRet[nPos]['Result'] := "Assinado"
		Else
			Aadd(aRet,JsonObject():new())
			nPos := Len(aRet)
			aRet[nPos]['Visita'] := cNumVis
			aRet[nPos]['Result'] := "VisitaNaoEncontrada"
		Endif
	Next i
  
	// Monta Objeto JSON de retorno
   	oJsonRet := NIL
   	oJsonRet := JsonObject():new()
    oJsonRet['AtualizacaoVis'] := aRet
  
   // Devolve o retorno para o Rest
   ::SetResponse(oJsonRet:toJSON())
        
   FreeObj(oJsonRet)
   FreeObj(oJson)
Return(lRet)
