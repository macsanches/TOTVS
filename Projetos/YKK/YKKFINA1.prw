#INCLUDE "rwmake.ch"
#INCLUDE "protheus.ch"
#INCLUDE "fwbrowse.ch"
#INCLUDE "fwmvcdef.ch"
#INCLUDE "totvs.ch"
#INCLUDE "tbiconn.ch"
#INCLUDE "tcbrowse.ch"


User Function YKKFINA1()
    // cria estrutura das pastas
    lEstr := YKKFINAe()

    // estrutura Ok.
    IF lEstr
        oProcess := MsNewProcess():New({|| YKKFINAl(oProcess)}, "Lendo Arquivos...", "Aguarde...", .T.)
        oProcess:Activate()
    ENDIF
Return

/*/ YKKFINAl
Efetua leitura dos arquivos nas Pastas 
@type       StaticFunction
@author     iVan Oliveira - Exitto
@since      10/05/2024
@version    1.0
@param      [_cArquivo]  , character, Nome local do arquivo a ser lido. 
@return     Logical      , Se foi executado com sucesso (.T.), ou .f. para erro     
@example    YKKFINAl
/*/ 
Static Function YKKFINAl(oRegua)

    Local lRet      := .T.
    Local nArqs     := 0
    Local nArqL     := 0
    Local cNomPasta := '\SapConcur\Disponiveis\'
    Local aErros    := {}
    Local cOpArq    := ''
    Local aTmpSE2   := {}
    Local nQtIt     := 0
    Local nQtArqPr  := 0

    // Retorna arquivos do diretorio SapConcur
    aFiles := Directory(cNomPasta + "*.txt", "D",,.T., 1)
    
    IF !empty(aFiles)
        IF MsgYesNo("Serão processado(s): " + StrZero(len(aFiles),3)+ ", arquivo(s) disponíveis na pasta: " + cNomPasta + Chr(13) + Chr(10), "Confirma?")
            // Verificando e validando arquivos a importar.
            oRegua:SetRegua1(Len(aFiles))
            
            For nArqs := 1 to len(aFiles)

            oRegua:IncRegua1("Arq.: " + Left(Alltrim(aFiles[nArqs,01]),20))

            // Validando os arquivos.
            aItProc := YKKFINAa(cNomPasta + aFiles[nArqs,01], @aErros, @cOpArq, oRegua,aFiles[nArqs,01]) 

            // Nenhum Erro encontrado passa para inclusão título
            if !empty(aItProc)
                nQtArqPr++
                Begin Transaction

                    cArquivo := cNomPasta + aFiles[nArqs,01]
                    //Ordena o Array por Filial + Representante + Natureza +Sequencial
                    // _aRet, { _cArquivo, _cFilProc, _cIdSAP, _cCodFor, _cLojaFor, _cCodCC, _cCtaCTB,  _nValTitRat, _cTipoPgto, _nSomFor, _cSeqItem, StrZero(_nLin,4), _cCtaCred, _cNatureza} )
                    aSort(aItProc, , , {|x, y| x[02] + x[03] + x[04] + x[12] + x[09] < y[02] + y[03] + y[04] + y[12] + x[09] })
                    // Separando Multi-Thread
                    cIndSep := aItProc[01][02] + aItProc[01][03] + aItProc[01][04] + aItProc[01][09] 
                    lURegOk := .F.
                    cArqSDir:= aFiles[nArqs,01]
                    for nQtIt :=1 to len(aItProc)

                        if  cIndSep <> aItProc[nQtIt][02] + aItProc[nQtIt][03] + aItProc[nQtIt][04] + aItProc[01][09] 

                            // se foi coletado registros.
                            if !empty(aTmpSE2)
                                // Tentativa de inclusão título
                                if !u_YKKFINA2(aTmpSE2,@aErros, cArquivo, aItProc[nQtIt][12])
                                    DisarmTransaction()
                                    AAdd( aErros, { cArquivo, aItProc[nQtIt][02], StrZero(nQtIt,4), cIndSep, '- Erro Autoexecução(FINA050), já demonstrado. O Arquivo não será movido até a resolução do problema !'} )
                                    nArqL++
                                    Exit
                                Endif
                            else
                                AAdd( aErros, { cArquivo, aItProc[nQtIt][02], StrZero(nQtIt,4), cIndSep, '- Ocorreu erro de seleção de registros do arquivo, informe o suporte esta ocorrência e tente novamente !'} )
                                nArqL++
                                Exit
                            Endif

                            cIndSep := aItProc[nQtIt][02] + aItProc[nQtIt][03] + aItProc[nQtIt][04] + aItProc[01][09]  
                            aTmpSE2 := {}

                            // Se último registro não foi incluido ainda (faz parte de outro bloco.)
                            if nQtIt == len(aItProc) .and. !lURegOk
                                nQtIt--
                            ElseIf nQtIt < len(aItProc)
                                nQtIt--
                            Endif

                        Else 
                            aadd(aTmpSE2, aItProc[nQtIt])  
                            // Já adicionado último quando conjunto for igual 
                            if nQtIt == len(aItProc)
                                cIndSep := ''
                                lURegOk := .T.
                                nQtIt--
                            Endif
                        Endif
                    Next

                End Transaction
                
                // Mover arquivo para Processados ou Rejeitados
                if AScan(aErros, {|x| AllTrim(x[01]) == cArqSDir }) == 0 
                    FErase( '\SapConcur\Processados\' + aFiles[nArqs,01] )
                    nStatus := frename(cArquivo, '\SapConcur\Processados\' + aFiles[nArqs,01] )
                    if nStatus == -1
                        MsgStop('Ocorreu um erro na tentativa de mover o arquivo : ' + _cArquivo + ;
                        ' para pasta de Processados. Informar o código do erro ao suporte e solicitar sua remoção manual. Erro: '+str(ferror(),4))
                    Endif
                Endif
            // Realiza a movimentação do Arquivo por falha
            Else
                // Mover para Rejeitados.
                cArquivo := cNomPasta + aFiles[nArqs,01]
                if cOpArq == 'R'
                    FErase( '\SapConcur\Rejeitados\' + aFiles[nArqs,01] )
                    nStatus := frename(cArquivo, '\SapConcur\Rejeitados\' + aFiles[nArqs,01] )
                    if nStatus == -1
                        MsgStop('Ocorreu um erro na tentativa de mover o arquivo : ' + cArquivo + ;
                        ' para pasta de Rejeitados. Informar o código do erro ao suporte e solicitar sua remoção manual. Erro: '+str(ferror(),4))
                    Endif
                    
                // Mover processados.
                ElseIf cOpArq == 'P'
                    FErase( '\SapConcur\Processados\' + aFiles[nArqs,01] )
                    nStatus := frename(cArquivo, '\SapConcur\Processados\' + aFiles[nArqs,01] )
                    if nStatus == -1
                        MsgStop('Ocorreu um erro na tentativa de mover o arquivo : ' + cArquivo + ;
                        ' para pasta de Processados. Informar o código do erro ao suporte e solicitar sua remoção manual. Erro: '+str(ferror(),4))
                    Endif
                ElseIf cOpArq == 'D'
                    // Deixar na Pasta de Disponiveis, pois poderá ser repeocessado novamente.
                Endif
            Endif
        Next
       
        // se houveram erros, imprime relatório.
        if !empty(aErros)

            aTmp := {}
            aEval(aErros,{|x| if(ascan(aTmp, alltrim(x[01])) == 0, aadd( aTmp, alltrim(x[01]) ),  ) } )

            MsgAlert("Foram lidos: " + StrZero(_nArqs-1,3) + " arquivos, sendo: " + StrZero(len(aTmp),3) + " com ressalvas e serão impressos a seguir, favor analisar e prosseguir com os ajustes necessários !","Final Processamento" )   
            // se houveram erros, imprime relatório.
            _cTitulo := "Sap Concur: LOG de Erros "
            _cDescri := " "
            _cReport := dtos ( date() ) + '_' + StrTran(Time(),':','')
            _aFilt   := {}
            _aTitulo := {}

            // Cabeçalho do relatório
            aadd( aTitulo, { 'Arquivo'       , 'COLUNA01', "@!", 200 } )
            aadd( aTitulo, { 'Fil.SAP'       , 'COLUNA02', "@!", 020 } )
            aadd( aTitulo, { 'nº.Linha'      , 'COLUNA03', "@!", 030 } )
            aadd( aTitulo, { 'Referência'    , 'COLUNA04', "@!", 070 } )
            aadd( aTitulo, { 'Descrição Erro', 'COLUNA05', "@!", 500 } )
            // Gerando relat para conf.
            u_YKKFINR1( cTitulo , cDescri , cReport , aFilt , aTitulo , aErros )
        Elseif nQtArqPr>0
            MsgInfo("Importações realizadas com sucesso. !","Final Processamento" )   
        Endif
    Else
        _nArqL:=-1
    Endif
Endif
        
// Se nenhum foi encontrado.
if len(aFiles) == 0  
    MsgAlert('Não existem arquivos disponíveis a importar na pasta: \SapConcur\Disponiveis\ ! ', 'YKKFINA1')
    lRet := .F.
ElseIf nArqL <0
    MsgAlert('Operação de importação cancelada pelo usuário, os arquivos permanecerão na pasta: \SapConcur\Disponiveis\ ! ', 'YKKFINA1')  
    lRet := .F.
ElseIf nQtArqPr==0
    MsgAlert('Foram encontrado(s) arquivo(s) sem o leiaute padrão e foram removidos para pasta: \YKKFin\Rejeitados ! ', 'YKKFINA1')  
    lRet := .F.
Endif

Return lRet

/*/ YKKFINAe
Criação da estrutura das Pastas de Leitura
@type       StaticFunction
@author     iVan Oliveira - Exitto
@since      10/05/2024
@version    1.0
@param      [_cArquivo]  , character, Nome local do arquivo a ser lido. 
@return     Logical      , Se foi executado com sucesso (.T.), ou .f. para erro     
@example    YKKFINAe
/*/ 
Static Function YKKFINAe()
    Local aEstr := { '\YKKFin\', '\YKKFin\Disponiveis\', '\YKKFin\Processados\', '\YKKFin\Rejeitados\' }
    Local nIt   := 0
    Local lRet  := .T.
 
    // Verificando a existência das Pastas
    For nIt := 1 to len(aEstr)
        If !ExistDir( aEstr[nIt] )
            nRet := MakeDir( aEstr[nIt] )
            If nRet != 0
                lRet  := .F.
                MsgStop( "Não foi possível criar a pasta: " +  aEstr[nIt] + ". Cod. Erro: " + cValToChar( FError() ) )
            Endif
        Endif
    Next
Return lRet

/*/ YKKFINAa
Carregar informações do arquivo a processar.
@type       StaticFunction
@author     iVan Oliveira - Exitto
@since      10/05/2024
@version    1.0
@param      [_cArquivo] , character , Nome local do arquivo a ser lido. 
@param      [_aRetErr]  , Array     , Itens com erro.
@param      [_cMovArq ] , character , Rótulo para movimentação de arquivos nas pastas(<D>isponiveis, <P>rocessados, <R>ejeitados)
@param      [_ObjR]     , Object    , Obj das barras de processamento
@param      [_cArqSDir] , character , Nome do arquivo a ser lido. 
@return     Array       , Array com o conteúdo do item a ser gerado financeiro
@example    YKKFINAa(_cNomPasta + _aFiles[_nArqs,01], @_aErros, @_cOpArq, _oRegua,_cArqSDir) 
/*/ 
Static Function YKKFINAa(cArquivo, aRetErr, cMovArq, ObjR, cArqSDir )
    Local nLin    := 0
    Local nSomFor := 0
    Local nTamCC  := TamSX3("CTT_CUSTO")[1]
    Local nTamCTA := TamSX3("CT1_CONTA")[1]
    Local nTamCFOR:= TamSX3("A2_COD")[1]
    Local nTamLj  := TamSX3("A2_LOJA")[1]
    Local nTamISAP:= TamSX3("E2_XIDSAP")[1]
    Local nTamFSE2:= TamSX3("E2_FILIAL")[1]
    Local nTamNtr := TamSX3("ED_CODIGO")[1]
    Local aRet    := {}
    Local lLinInc := .F.
    Local nQtDupl := 0
    Local aIdSap  := {}

    //Definindo o arquivo a ser lido
    oFile := FWFileReader():New(cArquivo)

    //Se o arquivo pode ser aberto
    If (oFile:Open())
        // Retorna todas as linhas do arquivo em um Array unidimensional
        aTmp := oFile:getAllLines()

        //Fecha o arquivo e processa o Array
        oFile:Close()

        // abrir arquivo
        ObjR:SetRegua2(len(aTmp))

        //Lendo as informações oriundas do arquivo.
        cItemProc := 'EXTRACT'
        cGrupoFor := ''
        
        For nLin := 1 to len(aTmp)
            // Carregando as Informações para carga de arquivo
            aString := Separa (aTmp[nLin],"|")

            // Somente arquivos com as colunas em leiaute
            If (len(aString)==5 .and. nLin == 1) .or. (len(aString)>300 .and. nLin > 1)
                cItemProc  := if(nLin == 1, 'EXTRACT', Alltrim(Upper(aString[SAP_REPORT_ID])))

                // Movimenta a Régua
                ObjR:IncRegua2("Registro: " + cItemProc)

                // Somente Linha Detalhe
                If left(Upper(aTmp[nLin]),6) == 'DETAIL'

                    // Tamanho de colunas pelo leiaute do arquivo.
                    If len(aString) == 399

                        cFilProc   := Padr(Alltrim(aString[SAP_ORG_UNIT_2_FL]),nTamFSE2)
                        cSeqItem   := Alltrim(Upper(aString[SAP_SEQUENCE_NUMBER]))	
                        cIdSAP     := Padr(Alltrim(Upper(aString[SAP_REPORT_ID])),nTamISAP)
                        cCodFor    := Padr(Alltrim(aString[SAP_EMP_ID_FOR]),nTamCFOR)
                        cCodCC     := Padr(Alltrim(aString[SAP_ORG_UNIT_3_CC]),nTamCC)
                        cFilDest   := Padr(Alltrim(aString[SAP_ORG_UNIT_FL_DEST]),nTamFSE2)
                        cCodCCD    := Padr(Alltrim(aString[SAP_ORG_UNIT_CC_DEST]),nTamCC)
                        cCtaCTB    := Padr(StrTran(Alltrim(aString[SAP_PAYMENT_CODE_CTA]),'.',''),nTamCTA)
                        nValTitRat := Val(aString[SAP_TOTAL_APPROVED_AMOUNT])
                        cTipoPgto  := Alltrim(Upper(aString[SAP_NET_RECLAIM_ADJAMOUNT])) // Definir se o arquivo é de cartão ou adiantamento.
                        cNatureza  := Padr(Alltrim(aString[SAP_NATUREZA]),nTamNtr)
                        cNroCartao := StrTran(Alltrim(Upper(aString[SAP_BILLING_AMOUNT])),'X','' )
                        cLojaFor   := Padr('01', nTamLj)

                        // Soma total por fornecedor
                        if cGrupoFor#cCodFor + cNatureza
                            nSomFor   := nValTitRat
                            cGrupoFor := cCodFor + cNatureza
                        Else
                            nSomFor += nValTitRat
                        Endif

                        // Validar se o item já foi importado.(Campo custom. SE2 SAP ID)
                        if !SE2->(MsSeek( cFilProc + cIdSAP + cTipoPgto))
                            if cFilProc<>cFilDest
                                if 'ITAU' $alltrim(cTipoPgto)
                                    nPosID := ascan(aForIntDiv,'ITAU')
                                    if nPosID == 0
                                        if AScan(aRetErr, {|x| x[01] == cArqSDir  .and. x[04] == cTipoPgto }) == 0 
                                            AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cTipoPgto, '- Não foi Definido um fornecedor Itaú para Conceito de Inter Divisão(ver parâmetro: FS_X_SAPID) !'} )
                                        Endif
                                    Else
                                        cCodFor := Left(StrTran(aForIntDiv[nPosID], 'ITAU', '' ), nTamCFOR)
                                        cLojaFor:= Padr(Right(alltrim(aForIntDiv[nPosID]),nTamLj), nTamLj)
                                    Endif
                                ElseIf   'BRADESCO' $alltrim(cTipoPgto)
                                    if nPosID := ascan(aForIntDiv,'BRADESCO') == 0
                                        if AScan(aRetErr, {|x| x[01] == cArqSDir  .and. x[04] == cTipoPgto }) == 0 
                                            AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cTipoPgto, '- Não foi Definido um fornecedor Bradesco para Conceito de Inter Divisão(ver parâmetro: FS_X_SAPID) !'}) 
                                        Endif
                                    Else
                                        cCodFor := Left(StrTran(aForIntDiv[nPosID], 'BRADESCO', '' ), nTamCFOR)
                                        cLojaFor:= Padr(Right(alltrim(aForIntDiv[nPosID]),nTamLj), nTamLj)
                                    Endif
                                Else
                                    if AScan(aRetErr, {|x| x[01] == cArqSDir  .and. x[04] == cTipoPgto }) == 0 
                                        AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cTipoPgto, '- Banco indefinido para Conceito de Inter Divisão(ver parâmetro: FS_X_SAPID) !'} )
                                    Endif
                                Endif
                                
                                // Tradução de Colunas Contas - Se Encontrou parâmetro realizar a tradução.
                                DbSelectArea("SX6") 
                                DbSetOrder(1)  
                                
                                aTmpTrad := {}
                                
                                If DbSeek( FwxFilial("SX6") + "FX_X_SAP" + Alltrim(cFilDest)) 
                                    aTmpTrad := StrTokArr(alltrim(SX6->X6_CONTEUD), '|')
                                    // Localizando o item se possuir tradutor na filial destino
                                    nPosCC := ascan(aTmpTrad, Alltrim(cCtaCTB ))

                                    if nPosCC > 0
                                        nPosCCC := AT( '-',  aTmpTrad[nPosCC] ) 
                                        if nPosCCC>0
                                            cCtaCTB := Padr(Alltrim(Substr(aTmpTrad[nPosCC],nPosCCC+1,len(aTmpTrad[nPosCC]))),nTamCTA)
                                        Endif
                                    Endif
                                EndIf

                                // Controle para todos pares irem com mesmo cod fornecedor (Conceito de Inter Divisão)
                                if AScan(aIdSap, {|x| x[01] == cIdSAP }) == 0 
                                    AAdd(aIdSap, {cIdSAP, cCodFor, cLojaFor })
                                Endif
                            Endif
 
                            // Se os códigos abaixo não foram relatados, procurar pela existência e/ou bloqueio.
                            if AScan(aRetErr, {|x| x[01] == cArqSDir  .and. x[04] == cCodFor }) == 0 
                                // Validar Fornecedor  --> Compartilhado.
                                If !SA2->(MsSeek(FWxFilial("SA2")  + cCodFor + cLojaFor ))
                                    AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cCodFor, '- Fornecedor não localizado !' } )
                                Else
                                    // Bloqueado?
                                    cLojaFor := SA2->A2_LOJA
                                    cCtaCred := SA2->A2_CONTA
                                    if SA2->A2_MSBLQL == '1'
                                        AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cCodFor, '- Fornecedor bloqueado para utilização !'} )

                                    Endif
                                    If Empty(SA2->A2_CONTA)
                                        AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cCodFor, '- Conta Fornecedor sem preenchimento !'} )

                                    Endif
                                endif
                            
                            Endif

                            if AScan(aRetErr, {|x| x[01] == cArqSDir  .and. x[04]== cCtaCTB } ) == 0 
                                // Validar Conta Ctb.   --> Compartilhado.
                                If !CT1->( MsSeek(FWxFilial("CT1") + cCtaCTB ))
                                    AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cCtaCTB, '- Conta Contábil não localizada !' } )

                                Else
                                    // Bloqueada?
                                    if CT1->CT1_BLOQ == '1'
                                        AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cCtaCTB, '- Conta Contábil bloqueada para utilização !'} )
                                    Endif
                                Endif
                            Endif

                            if AScan(aRetErr, {|x| x[01] == cArqSDir  .and. x[04]== cCodCC })  == 0 .or. ;
                               AScan(aRetErr, {|x| x[01] == cArqSDir  .and. x[04]== cCodCCD }) == 0 

                                // Validar CC. --> Exclusivo
                                If !CTT->( MsSeek(cFilProc + cCodCC ))
                                    AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cCodCC, '- Centro Custo não localizado !'  } )

                                Else
                                    // Bloqueado?
                                    if CTT->CTT_BLOQ == '1'
                                        AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cCodCC, '- Centro Custo bloqueado para utilização !' } )
                                    Endif
                                Endif

                                 // Validar CC. destino  
                                If !CTT->( MsSeek(cFilProc + cCodCCD ))
                                    AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cCodCCD, '- Centro Custo Destino(Conceito de Inter Divisão) não localizado !'  } )

                                Else
                                    // Bloqueado?
                                    if CTT->CTT_BLOQ == '1'
                                        AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cCodCCD, '- Centro Custo(Conceito de Inter Divisão) bloqueado para utilização !' } )
                                    Endif
                                Endif
                        
                            Endif

                            if AScan(aRetErr, {|x| x[01] == cArqSDir  .and. x[04] == cNatureza }) == 0 
                            // Validar natureza  --> Compartilhado.

                                If !SED->(MsSeek(FWxFilial("SED") + cNatureza ))
                                    AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cNatureza, '- Cód.Natureza não localizado !' } )
                                Else
                                    if SED->ED_MSBLQL == '1'
                                        AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cNatureza, '- Natureza bloqueada para utilização !'} )

                                    Endif
                                endif
                            
                            Endif
                            // Se houve algum erro anterior, não permitir executar.
                            if AScan(aRetErr, {|x| AllTrim(x[01]) == cArqSDir }) > 0
                                nValTitRat := 0
                            Endif
                        Else
                            if AScan(aRetErr, {|x| x[01] == cArqSDir  .and. x[04] == cIdSAP }) == 0 
                                AAdd( aRetErr, { cArqSDir, cFilProc, StrZero(nLin,4), cIdSAP, '- Este item já foi incluído no Financeiro Protheus!'} )
                            Endif
                            nValTitRat := 0
                            nQtDupl++
                        Endif

                    Else
                         lLinInc := .T.
                         aRet    := {}
                        Exit 
                    Endif

                    // Sem erro na linha, adicionar no array de retorno
                    if  nValTitRat#0

                        // Controle para todos pares irem com mesmo cod fornecedor (Conceito de Inter Divisão)
                        nPosID := AScan(aIdSap, {|x| x[01] == cIdSAP }) 
                        if nPosID > 0
                            if cCodFor#aIdSap[nPosID][02]
                                cCodFor := aIdSap[nPosID][02]
                                cLojaFor := aIdSap[nPosID][03]
                            Endif         
                        Endif
                        AAdd( aRet, {  cArqSDir, cFilProc, cIdSAP, cCodFor, cLojaFor, cCodCC, cCtaCTB,  nValTitRat, cTipoPgto, ;
                                        nSomFor, cSeqItem, StrZero(nLin,4), cCtaCred, cNatureza, cFilDest, cCodCCD, cNroCartao } )
                    Endif
                Endif
            else
                lLinInc := .T.
                Exit 
            Endif
        Next
        
        // Caso não encontre erros e não identificado no arquivo itens aptos a serem importaos, arquivo será colocado na pasta de rejeitado.
        if AScan(aRetErr, {|x| AllTrim(x[01]) == cArqSDir }) == 0 .and. Empty(aRet) .or. lLinInc
        

        // Se foi encontrado algum  erro no arquivo, não executar itens, deixando o arquivo na mesma pasta aguardando as devidas correções.
            _cMovArq := 'R'
        ElseIf  AScan(aRetErr, {|x| AllTrim(x[01]) == cArqSDir }) > 0 .and. Empty(aRet)  .and. (nQtDupl < len(aTmp)-1 ) .or.;
                AScan(aRetErr, {|x| AllTrim(x[01]) == cArqSDir }) > 0 .and. !Empty(aRet) .and. (nQtDupl < len(aTmp)-1 )
            aRet    := {}
            cMovArq := 'D'
        // Se todos os itens foram incluídos já, remover para 
        ElseIf AScan(aRetErr, {|x| AllTrim(x[01]) == cArqSDir }) > 0 .and. (len(aTmp)-1 == nQtDupl)
            cMovArq := 'P'
        Endif

    Endif
Else
    // marca o arquivo para ser movido para pasta de rejeitados
    AAdd( aRetErr, { cArquivo, cFilAnt, StrZero(nLin,4), '-', '- Não foi possível realizar abertura do arquivo, o arquivo será movido a pasta de Rejeição  !'} )
    cMovArq := 'R'
EndIf

Return aRet
