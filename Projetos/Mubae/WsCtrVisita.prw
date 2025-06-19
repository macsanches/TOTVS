#include "protheus.ch"
#include "totvs.ch"
#include "restful.ch"
#Include "RwMake.Ch"  
#include "topconn.ch"
#Include 'FWMVCDef.ch'
#Include "TBIConn.ch"
#include "Directry.ch"

	WSRESTFUL WsPedOp DESCRIPTION "Servico REST para integracao Protheus"
		WSDATA AtualizacaoPV AS STRING

		WSMETHOD POST DESCRIPTION "Sincronizacao de Dados via POST" WSSYNTAX "/WsPedAss"

	END WSRESTFUL

WSMETHOD POST WSRECEIVE RECEIVE WSSERVICE WsVisAssinatura

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

		If dbSeek(xFilial + cNumVis)
			RecLock("SC6",.F.)
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

