/******************************************************************************
* Programa.........:   TOPSKA										          *
* Módulo...........:   PCP      	                                          *
* Autor............:   Guilherme Carvalho 			                          *
* Solicitante......:   TopTools - SKA        			                      *
* Objetivo.........:   Integração com o sistema SKA para receber informações  *
*                      SKA realiza POST no Protheus.                          *
******************************************************************************/
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

		WSMETHOD POST DESCRIPTION "Sincronizacao de Dados via POST" WSSYNTAX "/WsPedOp"

	END WSRESTFUL

WSMETHOD POST WSRECEIVE RECEIVE WSSERVICE WsPedOp

Local lRet      := .T.         // Recebe o Retorno 
Local cBody     := ''          // Recebe o conteudo do Rest
Local oJson     := NIL         // Recebe o JSON de Entrada
Local oJsonRet  := NIL         // Recebe o JSON de Saida
Local cErr 		:= Space(0)
Local cJson		:= Space(0)
Local i 		:= 0
Local ret


// Pega o conteudo JSON da transação Rest
cBody := Self:GetContent() //::GetContent()
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

//cTeste := oJson[1]
//ConOut("tipo var: "+VALTYPE( cTeste ))

//ConOut("len: "+str(len(oJson)))

aAtuPV:= oJson:GetJsonObject("AtualizacaoPV")

For i := 1 to len(aAtuPV)

	cNumOP := aAtuPV[i]['CodOP']
	cDataReal := aAtuPV[i]['DataEntregaReal']
	cVersao := aAtuPV[i]['VersaoOP']
	cDataReal := SUBSTR(cDataReal,1,4)+SUBSTR(cDataReal,6,2)+SUBSTR(cDataReal,9,2)
	
	IF SUBSTRING(cNumOP,11,3) <> '001'
		LOOP
	ENDIF

	//ConOut("Gui: "+cNumOP+" - "+cDataReal+" - "+cVersao)
	
	dbSelectArea("SC6")
	dbSetOrder(7)	// INDICE POR NUMERO DE OP (C6_NUMOP+C6_ITEMOP)

	If dbSeek(SUBSTRING(cNumOP,1,10))
		RecLock("SC6",.F.)
		SC6->C6_XXDTENR := Stod(cDataReal)
		MsUnlock()
		Aadd(aRet,JsonObject():new())
		nPos := Len(aRet)
		aRet[nPos]['CodOP'] := cNumOP
		aRet[nPos]['Result'] := "C6Ok"
	Else
		Aadd(aRet,JsonObject():new())
		nPos := Len(aRet)
		aRet[nPos]['CodOP'] := cNumOP
		aRet[nPos]['Result'] := "PvNaoEncontrado"
	Endif
Next i

  
   // Monta Objeto JSON de retorno
   oJsonRet := NIL
   oJsonRet := JsonObject():new()
   /*oJsonRet['Propriedade1'] := EncodeUTF8("Retorno 1", "cp1252") 
   oJsonRet['Propriedade2'] := 10 
   oJsonRet['Propriedade3'] := .T.*/
    oJsonRet['AtualizacaoPV'] := aRet
  
   // Devolve o retorno para o Rest
   ::SetResponse(oJsonRet:toJSON())
        
   FreeObj(oJsonRet)
   FreeObj(oJson)

Return(lRet)
