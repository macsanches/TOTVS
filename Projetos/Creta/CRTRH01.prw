#INCLUDE "TOTVS.CH"
#INCLUDE "TBICONN.CH"
#INCLUDE "TOPCONN.CH"

/*/{Protheus.doc} CRTRH01
Relatorio de status dos colaboradores
@type function
@version 1.0.1 
@author Marcos A Sanches
@since 5/4/2025
/*/
User Function CRTRH01()
    Private oReport

    Processa({|| GeraPlanilha()},"Gerando planilha...", , , , )
Return                           

Static Function GeraPlanilha()
    Local aArea             := GetArea()
    Local cQuery            := ""
    Local oFWMsExcel    
    Local oExcel
    Local cArquivo          := GetTempPath()+'zTstExc1.xml'
    Local nSequencia        := 1
    Local cCidade           := ""
    Local Fone              := ""
    Local cTpConta          := ""
    Local cCusto            := ""
    Local cDepto            := ""
    Local cCargo            := ""
    Local cEvento           := ""
    Local cBruto            := 0
    Local cLiquido          := 0
    Local cInsalubridade    := 0
    Local cVT               := 0
    Local cVA               := 0
    Local nPos              := 0
    Local nHoras            := 0
    Local cInicial          := ""
    Local cFinal            := ""

    Pergunte("CRTRH01", .T.)
 
    //Pegando os dados
    cQuery := " SELECT RA_FILIAL,RA_MAT,RA_NOMECMP,CONVERT(CHAR(10),RA_NASC,103) AS RA_NASC,RA_LOGRDSC,RA_LOGRNUM,RA_BAIRRO,RA_CODMUN,RA_CEP, RA_TELEFON AS RA_FONE, RA_DDDFONE AS RA_DDD"
    cQuery += " ,RA_RG,RA_CIC,RA_PIS,SUBSTRING(RA_BCDEPSA,1,3) AS RA_BANCO, SUBSTRING(RA_BCDEPSA,4,10) AS RA_AGENCIA,RA_RGUF "
    cQuery += " ,RA_CTDEPSA, RA_TPCTSAL "
    cQuery += " ,RA_ADMISSA,RA_CC,RA_DEPTO,RA_HRSMES"
    cQuery += " ,CASE WHEN RA_SITFOLH = ' ' THEN 'Ativo' WHEN RA_SITFOLH = 'D' THEN 'Desligado' WHEN RA_SITFOLH = 'F' THEN 'Ferias' ELSE 'Nao informado' END AS RA_SITFOLH "
    cQuery += " ,RA_CODFUNC"
    cQuery += " FROM "+RetSqlName("SRA")+" SRA "
    cQuery += " WHERE SRA.D_E_L_E_T_ = '' "  

    //Filial
    If MV_PAR01 <> "" .And. MV_PAR02 <> ""
        cQuery += " AND RA_FILIAL BETWEEN '" + MV_PAR01 + "' AND '" + MV_PAR02 + "'"
    Endif

    //Centro de custo
    IF !Empty(MV_PAR04) .And. !Empty(MV_PAR05)
        cQuery += " AND RA_CC BETWEEN '" + MV_PAR05 + "' AND '" + MV_PAR06 + "'"
    Endif

    //Matricula
    If !Empty(MV_PAR06) .And. !Empty(MV_PAR07)
        cQuery += " AND RA_MAT BETWEEN '" + MV_PAR07 + "' AND '" + MV_PAR08 + "'"
    Endif

    //Situacao
    If !Empty(MV_PAR08)
        //cQuery += " AND RA_NOMECMP LIKE '%" + MV_PAR09 + "%'"
    Endif

    //Categorias
    If !Empty(MV_PAR09)
        //cQuery += " AND RA_CODFUNC BETWEEN '" + MV_PAR10 + "' AND '" + MV_PAR11 + "'"
    Endif

    //Data de admissão
    If !Empty(MV_PAR10)  .And. !Empty(MV_PAR11)
        cQuery += " AND RA_ADMISSA BETWEEN '" + Dtos(MV_PAR10) + "' AND '" + Dtos(MV_PAR11) + "'"
    Endif

    cQuery += " ORDER BY RA_MAT "

   //salva o código sql na pasta TEMP para consultas no seu SGBD
    memoWrite("\TEMP\RELSRA.sql",cQuery)     
    TCQuery cQuery New Alias "TRB"

    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New()
    oFWMsExcel:AddworkSheet("Geral")
                   
    //Criando a Tabela
    oFWMsExcel:AddTable("Geral","Colaborador")
    oFWMsExcel:AddColumn("Geral","Colaborador","Nº",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Matricula",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Nome completo",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Data de Nascimento",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Endereço",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Num",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Bairro",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Cidade",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Cep",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","DDD + Telefone",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","RG",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","CPF",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","PIS",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Banco",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Agencia",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Conta",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Tipo conta",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Data de admissão",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Centro de custo",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Departamento",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Situacao",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Data do envento",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Substituto",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Cargo",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Carga horaria",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Inicial",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Final",1)
    oFWMsExcel:AddColumn("Geral","Colaborador","Sal.Bruto",3)
    oFWMsExcel:AddColumn("Geral","Colaborador","Sal.Liquido",3)
    oFWMsExcel:AddColumn("Geral","Colaborador","Insalubridade",3)
    oFWMsExcel:AddColumn("Geral","Colaborador","VT",3)
    oFWMsExcel:AddColumn("Geral","Colaborador","VA",3)
    oFWMsExcel:AddColumn("Geral","Colaborador","Total",3)
    
    //Criando as Linhas... Enquanto não for fim da query
    While !(TRB->(EoF()))
         if Select("TMPAux") > 0
            TMPAux->(dbCloseArea())
        Endif

        cQuery := "SELECT TOP 100 * FROM " + RetSqlName("CC2") + " WHERE D_E_L_E_T_ = '' AND CC2_CODMUN = '" + TRB->RA_CODMUN + "'"
        TCQuery cQuery New Alias "TMPAux"

        cCidade := TMPAux->CC2_MUN

        If Len(Alltrim(TRB->RA_FONE)) <= 9
            Fone := "(" + TRB->RA_DDD + ") " + SUBSTRING(TRB->RA_FONE,1,5) + "-" + SUBSTRING(TRB->RA_FONE,6,4)
        Elseif Len(Alltrim(TRB->RA_FONE)) > 10
            Fone := "(" + SUBSTRING(TRB->RA_FONE,1,2) + ") " + SUBSTRING(TRB->RA_FONE,3,5) + "-" + SUBSTRING(TRB->RA_FONE,8,4)
        Else
            Fone := "(" + SUBSTRING(TRB->RA_FONE,1,2) + ") " + SUBSTRING(TRB->RA_FONE,3,5) + "-" + SUBSTRING(TRB->RA_FONE,8,4)
        Endif

        IF TRB->RA_TPCTSAL == "1"
            cTpConta := "Conta corrente"
        ELSEIF TRB->RA_TPCTSAL == "2"
            cTpConta := "Conta poupança"
        ELSE
            cTpConta := "Não informado"
        Endif
        
        TMPAux->(dbCloseArea())
        
        cQuery := "SELECT * FROM " + RetSqlName("CTT") + " WHERE D_E_L_E_T_ = '' AND CTT_CUSTO = '" + TRB->RA_CC + "'"
        TCQuery cQuery New Alias "TMPAux"

        cCusto := TMPAux->CTT_DESC01

        TMPAux->(dbCloseArea())

        cQuery := "SELECT * FROM " + RetSqlName("SQB") + " WHERE D_E_L_E_T_ = '' AND QB_DEPTO = '" + TRB->RA_DEPTO + "'"
        TCQuery cQuery New Alias "TMPAux"

        cDepto := TMPAux->QB_DESCRIC

        TMPAux->(dbCloseArea())

        cQuery := "SELECT * FROM " + RetSqlName("SRJ") + " WHERE D_E_L_E_T_ = '' AND RJ_FUNCAO = '" + TRB->RA_CODFUNC + "'"
        TCQuery cQuery New Alias "TMPAux"

        cCargo := TMPAux->RJ_DESC

        TMPAux->(dbCloseArea())

        cQuery := "SELECT * FROM " + RetSqlName("SR8") + " WHERE D_E_L_E_T_ = '' AND R8_MAT = '" + TRB->RA_MAT + "' AND R8_FILIAL = '" + TRB->RA_FILIAL + "'"
        TCQuery cQuery New Alias "TMPAux"

        cEvento := TMPAux->R8_DATA

        TMPAux->(dbCloseArea())
        
        if !Empty(MV_PAR03)
            cQuery := "SELECT * FROM " + RetSqlName("RFQ") + " WHERE D_E_L_E_T_ = '' AND RFQ_PERIOD = '" + MV_PAR03 + "'"

            TCQuery cQuery New Alias "TMPAux"

            IF TMPAux->RFQ_STATUS == "1"
                //Mes aberto - SRC
                cQuery := "SELECT * FROM " + RetSqlName("SRC") + " WHERE D_E_L_E_T_ = '' AND RC_FILIAL = '" + TRB->RA_FILIAL + "' AND RC_MAT = '" + TRB->RA_MAT + "' AND RC_PERIODO = '" + MV_PAR03 + "'"
                TCQuery cQuery New Alias "TMPfol"
                
                While !(TMPfol->(EoF()))
                    IF TMPfol->RC_PD == "101"
                        cBruto := TMPfol->RC_VALOR
                    Elseif TMPfol->RC_PD == "799"
                        cliquido := TMPfol->RC_VALOR
                    Elseif TMPfol->RC_PD == "010" .OR. TMPfol->RC_PD == "011" .OR. TMPfol->RC_PD == "012"
                        cInsalubridade := TMPfol->RC_VALOR
                    Elseif TMPfol->RC_PD == "784"
                        cVT := TMPfol->RC_VALOR
                    Elseif TMPfol->RC_PD == "998"
                        cVA := TMPfol->RC_VALOR
                    Endif

                    TMPfol->(dbSkip())
                Enddo   
                TMPfol->(dbCloseArea())
            Else
                //Mes fechado - SRD
                cQuery := "SELECT * FROM " + RetSqlName("SRD") + " WHERE D_E_L_E_T_ = '' AND RD_FILIAL = '" + TRB->RA_FILIAL + "' AND RD_MAT = '" + TRB->RA_MAT + "' AND RD_PERIODO = '" + MV_PAR03 + "'"
                TCQuery cQuery New Alias "TMPfol"
                
                While !(TMPfol->(EoF()))
                    IF TMPfol->RD_PD == "101"
                        cBruto := TMPfol->RD_VALOR
                    Elseif TMPfol->RD_PD == "799"
                        cliquido := TMPfol->RD_VALOR
                    Elseif TMPfol->RD_PD == "010" .OR. TMPfol->RD_PD == "011" .OR. TMPfol->RD_PD == "012"
                        cInsalubridade := TMPfol->RD_VALOR
                    Elseif TMPfol->RD_PD == "784"
                        cVT := TMPfol->RD_VALOR
                    Elseif TMPfol->RD_PD == "998"
                        cVA := TMPfol->RD_VALOR
                    Endif

                    TMPfol->(dbSkip())
                Enddo   
                TMPfol->(dbCloseArea())
            Endif 
        Endif

        IF nPos := AT("Diurno",cCargo) <> 0
            nHoras := 84
            cInicial := "07:00"
            cFinal := "19:00"
        ELSEIF nPos := AT("Noturno",cCargo) <> 0
            nHoras := 84
            cInicial := "19:00"
            cFinal := "07:00"
        Else
            nHoras := 220
            cInicial := "07:00"
            cFinal := "17:00"
        Endif

        oFWMsExcel:AddRow("Geral","Colaborador",{;
                                    nSequencia,;
                                    TRB->RA_MAT,;
                                    TRB->RA_NOMECMP,;
                                    SUBSTRING(TRB->RA_NASC,7,2) + '/' + SUBSTRING(TRB->RA_NASC,5,2) + '/' + SUBSTRING(TRB->RA_NASC,1,4),;
                                    TRB->RA_LOGRDSC,;
                                    TRB->RA_LOGRNUM,;
                                    TRB->RA_BAIRRO,;
                                    cCidade,;
                                    SUBSTRING(TRB->RA_CEP,1,5) + "-" + SUBSTRING(TRB->RA_CEP,6,3),;
                                    Fone,;
                                    SUBSTRING(TRB->RA_RG,1,2) + "." + SUBSTRING(TRB->RA_RG,3,3) + "." + SUBSTRING(TRB->RA_RG,6,3) + "-" + SUBSTRING(TRB->RA_RG,9,2) + " - " + TRB->RA_RGUF,;
                                    SUBSTRING(TRB->RA_CIC,1,3) + "." + SUBSTRING(TRB->RA_CIC,4,3) + "." + SUBSTRING(TRB->RA_CIC,7,3) + "-" + SUBSTRING(TRB->RA_CIC,10,2),;
                                    TRB->RA_PIS,;
                                    TRB->RA_BANCO,;
                                    TRB->RA_AGENCIA,;
                                    TRB->RA_CTDEPSA,;
                                    cTpConta,;
                                    SUBSTRING(TRB->RA_ADMISSA,7,2) + "/" + SUBSTRING(TRB->RA_ADMISSA,5,2) + "/" + SUBSTRING(TRB->RA_ADMISSA,1,4),;
                                    cCusto,;
                                    cDepto,;
                                    TRB->RA_SITFOLH,;
                                    SUBSTRING(cEvento,7,2) + "/" + SUBSTRING(cEvento,5,2) + "/" + SUBSTRING(cEvento,1,4),;
                                    Space(10),;
                                    cCargo,;
                                    nHoras,;
                                    cInicial,;
                                    cFinal,;
                                    cBruto,;
                                    cliquido,;
                                    cInsalubridade,;
                                    cVT,;
                                    cVA,;
                                    cliquido + cVA + cVT})

        TRB->(DbSkip())
        nSequencia++
    EndDo

    TRB->(DbCloseArea())
    nSequencia := 1
    
    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel:AddworkSheet("Lic.Materinadade")

    //Criando a Tabela - secundaria
    oFWMsExcel:AddTable("Lic.Materinadade","Colaborador")
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Nº",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Matricula",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Nome completo",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Data de Nascimento",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Endereço",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Num",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Bairro",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Cidade",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Cep",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","DDD + Telefone",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","RG",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","CPF",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","PIS",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Banco",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Agencia",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Conta",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Tipo conta",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Data de admissão",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Centro de custo",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Departamento",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Situacao",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Data do envento",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Substituto",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Cargo",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Carga horaria",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Inicial",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Final",1)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Sal.Bruto",3)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Sal.Liquido",3)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Insalubridade",3)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","VT",3)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","VA",3)
    oFWMsExcel:AddColumn("Lic.Materinadade","Colaborador","Total",3)

    //Pegando os dados
    cQuery := " SELECT RA_FILIAL,RA_MAT,RA_NOMECMP,CONVERT(CHAR(10),RA_NASC,103) AS RA_NASC,RA_LOGRDSC,RA_LOGRNUM,RA_BAIRRO,RA_CODMUN,RA_CEP, RA_TELEFON AS RA_FONE, RA_DDDFONE AS RA_DDD"
    cQuery += " ,RA_RG,RA_CIC,RA_PIS,SUBSTRING(RA_BCDEPSA,1,3) AS RA_BANCO, SUBSTRING(RA_BCDEPSA,4,10) AS RA_AGENCIA,RA_RGUF "
    cQuery += " ,RA_CTDEPSA, RA_TPCTSAL "
    cQuery += " ,RA_ADMISSA,RA_CC,RA_DEPTO,RA_HRSMES"
    cQuery += " ,CASE WHEN RA_SITFOLH = ' ' THEN 'Ativo' WHEN RA_SITFOLH = 'D' THEN 'Desligado' WHEN RA_SITFOLH = 'F' THEN 'Ferias' ELSE 'Nao informado' END AS RA_SITFOLH "
    cQuery += " ,RA_CODFUNC"
    cQuery += " FROM "+RetSqlName("SRA") +" SRA INNER JOIN " +RetSqlName("SR8")+" SR8 ON SRA.RA_MAT = SR8.R8_MAT AND SRA.RA_FILIAL = SR8.R8_FILIAL"
    cQuery += " WHERE SRA.D_E_L_E_T_ = '' AND SR8.R8_TIPOAFA IN ('010','011','012','006','007')"
    
    //Filial
    If MV_PAR01 <> "" .And. MV_PAR02 <> ""
        cQuery += " AND RA_FILIAL BETWEEN '" + MV_PAR01 + "' AND '" + MV_PAR02 + "'"
    Endif

    //Centro de custo
    IF !Empty(MV_PAR04) .And. !Empty(MV_PAR05)
        cQuery += " AND RA_CC BETWEEN '" + MV_PAR05 + "' AND '" + MV_PAR06 + "'"
    Endif

    //Matricula
    If !Empty(MV_PAR06) .And. !Empty(MV_PAR07)
        cQuery += " AND RA_MAT BETWEEN '" + MV_PAR07 + "' AND '" + MV_PAR08 + "'"
    Endif

    //Situacao
    If !Empty(MV_PAR08)
        //cQuery += " AND RA_NOMECMP LIKE '%" + MV_PAR09 + "%'"
    Endif

    //Categorias
    If !Empty(MV_PAR09)
        //cQuery += " AND RA_CODFUNC BETWEEN '" + MV_PAR10 + "' AND '" + MV_PAR11 + "'"
    Endif

    //Data de admissão
    If !Empty(MV_PAR10)  .And. !Empty(MV_PAR11)
        cQuery += " AND RA_ADMISSA BETWEEN '" + Dtos(MV_PAR10) + "' AND '" + Dtos(MV_PAR11) + "'"
    Endif

    cQuery += " ORDER BY RA_MAT "

    memoWrite("\TEMP\RELSRA.sql",cQuery)     
    TCQuery cQuery New Alias "TRB"

    //Criando as Linhas... Enquanto não for fim da query
    While !(TRB->(EoF()))
         if Select("TMPAux") > 0
            TMPAux->(dbCloseArea())
        Endif

        cQuery := "SELECT TOP 100 * FROM " + RetSqlName("CC2") + " WHERE D_E_L_E_T_ = '' AND CC2_CODMUN = '" + TRB->RA_CODMUN + "'"
        TCQuery cQuery New Alias "TMPAux"

        cCidade := TMPAux->CC2_MUN

        If Len(Alltrim(TRB->RA_FONE)) <= 9
            Fone := "(" + TRB->RA_DDD + ") " + SUBSTRING(TRB->RA_FONE,1,5) + "-" + SUBSTRING(TRB->RA_FONE,6,4)
        Elseif Len(Alltrim(TRB->RA_FONE)) > 10
            Fone := "(" + SUBSTRING(TRB->RA_FONE,1,2) + ") " + SUBSTRING(TRB->RA_FONE,3,5) + "-" + SUBSTRING(TRB->RA_FONE,8,4)
        Else
            Fone := "(" + SUBSTRING(TRB->RA_FONE,1,2) + ") " + SUBSTRING(TRB->RA_FONE,3,5) + "-" + SUBSTRING(TRB->RA_FONE,8,4)
        Endif

        IF TRB->RA_TPCTSAL == "1"
            cTpConta := "Conta corrente"
        ELSEIF TRB->RA_TPCTSAL == "2"
            cTpConta := "Conta poupança"
        ELSE
            cTpConta := "Não informado"
        Endif
        
        TMPAux->(dbCloseArea())
        
        cQuery := "SELECT * FROM " + RetSqlName("CTT") + " WHERE D_E_L_E_T_ = '' AND CTT_CUSTO = '" + TRB->RA_CC + "'"
        TCQuery cQuery New Alias "TMPAux"

        cCusto := TMPAux->CTT_DESC01

        TMPAux->(dbCloseArea())

        cQuery := "SELECT * FROM " + RetSqlName("SQB") + " WHERE D_E_L_E_T_ = '' AND QB_DEPTO = '" + TRB->RA_DEPTO + "'"
        TCQuery cQuery New Alias "TMPAux"

        cDepto := TMPAux->QB_DESCRIC

        TMPAux->(dbCloseArea())

        cQuery := "SELECT * FROM " + RetSqlName("SRJ") + " WHERE D_E_L_E_T_ = '' AND RJ_FUNCAO = '" + TRB->RA_CODFUNC + "'"
        TCQuery cQuery New Alias "TMPAux"

        cCargo := TMPAux->RJ_DESC

        TMPAux->(dbCloseArea())

        cQuery := "SELECT * FROM " + RetSqlName("SR8") + " WHERE D_E_L_E_T_ = '' AND R8_MAT = '" + TRB->RA_MAT + "' AND R8_FILIAL = '" + TRB->RA_FILIAL + "'"
        TCQuery cQuery New Alias "TMPAux"

        cEvento := TMPAux->R8_DATA

        TMPAux->(dbCloseArea())
        
        if !Empty(MV_PAR03)
            cQuery := "SELECT * FROM " + RetSqlName("RFQ") + " WHERE D_E_L_E_T_ = '' AND RFQ_PERIOD = '" + MV_PAR03 + "'"

            TCQuery cQuery New Alias "TMPAux"

            IF TMPAux->RFQ_STATUS == "1"
                //Mes aberto - SRC
                cQuery := "SELECT * FROM " + RetSqlName("SRC") + " WHERE D_E_L_E_T_ = '' AND RC_FILIAL = '" + TRB->RA_FILIAL + "' AND RC_MAT = '" + TRB->RA_MAT + "' AND RC_PERIODO = '" + MV_PAR03 + "'"
                TCQuery cQuery New Alias "TMPfol"
                
                While !(TMPfol->(EoF()))
                    IF TMPfol->RC_PD == "101"
                        cBruto := TMPfol->RC_VALOR
                    Elseif TMPfol->RC_PD == "799"
                        cliquido := TMPfol->RC_VALOR
                    Elseif TMPfol->RC_PD == "010" .OR. TMPfol->RC_PD == "011" .OR. TMPfol->RC_PD == "012"
                        cInsalubridade := TMPfol->RC_VALOR
                    Elseif TMPfol->RC_PD == "784"
                        cVT := TMPfol->RC_VALOR
                    Elseif TMPfol->RC_PD == "998"
                        cVA := TMPfol->RC_VALOR
                    Endif

                    TMPfol->(dbSkip())
                Enddo   
                TMPfol->(dbCloseArea())
            Else
                //Mes fechado - SRD
                cQuery := "SELECT * FROM " + RetSqlName("SRD") + " WHERE D_E_L_E_T_ = '' AND RD_FILIAL = '" + TRB->RA_FILIAL + "' AND RD_MAT = '" + TRB->RA_MAT + "' AND RD_PERIODO = '" + MV_PAR03 + "'"
                TCQuery cQuery New Alias "TMPfol"
                
                While !(TMPfol->(EoF()))
                    IF TMPfol->RD_PD == "101"
                        cBruto := TMPfol->RD_VALOR
                    Elseif TMPfol->RD_PD == "799"
                        cliquido := TMPfol->RD_VALOR
                    Elseif TMPfol->RD_PD == "010" .OR. TMPfol->RD_PD == "011" .OR. TMPfol->RD_PD == "012"
                        cInsalubridade := TMPfol->RD_VALOR
                    Elseif TMPfol->RD_PD == "784"
                        cVT := TMPfol->RD_VALOR
                    Elseif TMPfol->RD_PD == "998"
                        cVA := TMPfol->RD_VALOR
                    Endif

                    TMPfol->(dbSkip())
                Enddo   
                TMPfol->(dbCloseArea())
            Endif 
        Endif

        IF nPos := AT("Diurno",cCargo) <> 0
            nHoras := 84
            cInicial := "07:00"
            cFinal := "19:00"
        ELSEIF nPos := AT("Noturno",cCargo) <> 0
            nHoras := 84
            cInicial := "19:00"
            cFinal := "07:00"
        Else
            nHoras := 220
            cInicial := "07:00"
            cFinal := "17:00"
        Endif

        oFWMsExcel:AddRow("Lic.Materinadade","Colaborador",{;
                                    nSequencia,;
                                    TRB->RA_MAT,;
                                    TRB->RA_NOMECMP,;
                                    SUBSTRING(TRB->RA_NASC,7,2) + '/' + SUBSTRING(TRB->RA_NASC,5,2) + '/' + SUBSTRING(TRB->RA_NASC,1,4),;
                                    TRB->RA_LOGRDSC,;
                                    TRB->RA_LOGRNUM,;
                                    TRB->RA_BAIRRO,;
                                    cCidade,;
                                    SUBSTRING(TRB->RA_CEP,1,5) + "-" + SUBSTRING(TRB->RA_CEP,6,3),;
                                    Fone,;
                                    SUBSTRING(TRB->RA_RG,1,2) + "." + SUBSTRING(TRB->RA_RG,3,3) + "." + SUBSTRING(TRB->RA_RG,6,3) + "-" + SUBSTRING(TRB->RA_RG,9,2) + " - " + TRB->RA_RGUF,;
                                    SUBSTRING(TRB->RA_CIC,1,3) + "." + SUBSTRING(TRB->RA_CIC,4,3) + "." + SUBSTRING(TRB->RA_CIC,7,3) + "-" + SUBSTRING(TRB->RA_CIC,10,2),;
                                    TRB->RA_PIS,;
                                    TRB->RA_BANCO,;
                                    TRB->RA_AGENCIA,;
                                    TRB->RA_CTDEPSA,;
                                    cTpConta,;
                                    SUBSTRING(TRB->RA_ADMISSA,7,2) + "/" + SUBSTRING(TRB->RA_ADMISSA,5,2) + "/" + SUBSTRING(TRB->RA_ADMISSA,1,4),;
                                    cCusto,;
                                    cDepto,;
                                    TRB->RA_SITFOLH,;
                                    SUBSTRING(cEvento,7,2) + "/" + SUBSTRING(cEvento,5,2) + "/" + SUBSTRING(cEvento,1,4),;
                                    Space(10),;
                                    cCargo,;
                                    nHoras,;
                                    cInicial,;
                                    cFinal,;
                                    cBruto,;
                                    cliquido,;
                                    cInsalubridade,;
                                    cVT,;
                                    cVA,;
                                    cliquido + cVA + cVT})

        TRB->(DbSkip())
        nSequencia++
    EndDo
    
    TRB->(DbCloseArea())
    nSequencia := 1

    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel:AddworkSheet("Afastamentos")

    //Criando a Tabela - secundaria
    oFWMsExcel:AddTable("Afastamentos","Colaborador")
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Nº",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Matricula",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Nome completo",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Data de Nascimento",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Endereço",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Num",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Bairro",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Cidade",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Cep",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","DDD + Telefone",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","RG",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","CPF",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","PIS",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Banco",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Agencia",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Conta",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Tipo conta",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Data de admissão",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Centro de custo",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Departamento",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Situacao",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Data do envento",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Substituto",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Cargo",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Carga horaria",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Inicial",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Final",1)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Sal.Bruto",3)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Sal.Liquido",3)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Insalubridade",3)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","VT",3)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","VA",3)
    oFWMsExcel:AddColumn("Afastamentos","Colaborador","Total",3)

    //Pegando os dados
    cQuery := " SELECT RA_FILIAL,RA_MAT,RA_NOMECMP,CONVERT(CHAR(10),RA_NASC,103) AS RA_NASC,RA_LOGRDSC,RA_LOGRNUM,RA_BAIRRO,RA_CODMUN,RA_CEP, RA_TELEFON AS RA_FONE, RA_DDDFONE AS RA_DDD"
    cQuery += " ,RA_RG,RA_CIC,RA_PIS,SUBSTRING(RA_BCDEPSA,1,3) AS RA_BANCO, SUBSTRING(RA_BCDEPSA,4,10) AS RA_AGENCIA,RA_RGUF "
    cQuery += " ,RA_CTDEPSA, RA_TPCTSAL "
    cQuery += " ,RA_ADMISSA,RA_CC,RA_DEPTO,RA_HRSMES"
    cQuery += " ,CASE WHEN RA_SITFOLH = ' ' THEN 'Ativo' WHEN RA_SITFOLH = 'D' THEN 'Desligado' WHEN RA_SITFOLH = 'F' THEN 'Ferias' ELSE 'Nao informado' END AS RA_SITFOLH "
    cQuery += " ,RA_CODFUNC"
    cQuery += " FROM "+RetSqlName("SRA") +" SRA INNER JOIN " +RetSqlName("SR8")+" SR8 ON SRA.RA_MAT = SR8.R8_MAT AND SRA.RA_FILIAL = SR8.R8_FILIAL"
    cQuery += " WHERE SRA.D_E_L_E_T_ = '' AND SR8.R8_TIPOAFA NOT IN ('003','004','005','013','014','016','017','018')"

    //Filial
    If MV_PAR01 <> "" .And. MV_PAR02 <> ""
        cQuery += " AND RA_FILIAL BETWEEN '" + MV_PAR01 + "' AND '" + MV_PAR02 + "'"
    Endif

    //Centro de custo
    IF !Empty(MV_PAR04) .And. !Empty(MV_PAR05)
        cQuery += " AND RA_CC BETWEEN '" + MV_PAR05 + "' AND '" + MV_PAR06 + "'"
    Endif

    //Matricula
    If !Empty(MV_PAR06) .And. !Empty(MV_PAR07)
        cQuery += " AND RA_MAT BETWEEN '" + MV_PAR07 + "' AND '" + MV_PAR08 + "'"
    Endif

    //Situacao
    If !Empty(MV_PAR08)
        //cQuery += " AND RA_NOMECMP LIKE '%" + MV_PAR09 + "%'"
    Endif

    //Categorias
    If !Empty(MV_PAR09)
        //cQuery += " AND RA_CODFUNC BETWEEN '" + MV_PAR10 + "' AND '" + MV_PAR11 + "'"
    Endif

    //Data de admissão
    If !Empty(MV_PAR10)  .And. !Empty(MV_PAR11)
        cQuery += " AND RA_ADMISSA BETWEEN '" + Dtos(MV_PAR10) + "' AND '" + Dtos(MV_PAR11) + "'"
    Endif

    cQuery += " ORDER BY RA_MAT "

    memoWrite("\TEMP\RELSRA.sql",cQuery)     
    TCQuery cQuery New Alias "TRB"

    //Criando as Linhas... Enquanto não for fim da query
    While !(TRB->(EoF()))
         if Select("TMPAux") > 0
            TMPAux->(dbCloseArea())
        Endif

        cQuery := "SELECT TOP 100 * FROM " + RetSqlName("CC2") + " WHERE D_E_L_E_T_ = '' AND CC2_CODMUN = '" + TRB->RA_CODMUN + "'"
        TCQuery cQuery New Alias "TMPAux"

        cCidade := TMPAux->CC2_MUN

        If Len(Alltrim(TRB->RA_FONE)) <= 9
            Fone := "(" + TRB->RA_DDD + ") " + SUBSTRING(TRB->RA_FONE,1,5) + "-" + SUBSTRING(TRB->RA_FONE,6,4)
        Elseif Len(Alltrim(TRB->RA_FONE)) > 10
            Fone := "(" + SUBSTRING(TRB->RA_FONE,1,2) + ") " + SUBSTRING(TRB->RA_FONE,3,5) + "-" + SUBSTRING(TRB->RA_FONE,8,4)
        Else
            Fone := "(" + SUBSTRING(TRB->RA_FONE,1,2) + ") " + SUBSTRING(TRB->RA_FONE,3,5) + "-" + SUBSTRING(TRB->RA_FONE,8,4)
        Endif

        IF TRB->RA_TPCTSAL == "1"
            cTpConta := "Conta corrente"
        ELSEIF TRB->RA_TPCTSAL == "2"
            cTpConta := "Conta poupança"
        ELSE
            cTpConta := "Não informado"
        Endif
        
        TMPAux->(dbCloseArea())
        
        cQuery := "SELECT * FROM " + RetSqlName("CTT") + " WHERE D_E_L_E_T_ = '' AND CTT_CUSTO = '" + TRB->RA_CC + "'"
        TCQuery cQuery New Alias "TMPAux"

        cCusto := TMPAux->CTT_DESC01

        TMPAux->(dbCloseArea())

        cQuery := "SELECT * FROM " + RetSqlName("SQB") + " WHERE D_E_L_E_T_ = '' AND QB_DEPTO = '" + TRB->RA_DEPTO + "'"
        TCQuery cQuery New Alias "TMPAux"

        cDepto := TMPAux->QB_DESCRIC

        TMPAux->(dbCloseArea())

        cQuery := "SELECT * FROM " + RetSqlName("SRJ") + " WHERE D_E_L_E_T_ = '' AND RJ_FUNCAO = '" + TRB->RA_CODFUNC + "'"
        TCQuery cQuery New Alias "TMPAux"

        cCargo := TMPAux->RJ_DESC

        TMPAux->(dbCloseArea())

        cQuery := "SELECT * FROM " + RetSqlName("SR8") + " WHERE D_E_L_E_T_ = '' AND R8_MAT = '" + TRB->RA_MAT + "' AND R8_FILIAL = '" + TRB->RA_FILIAL + "'"
        TCQuery cQuery New Alias "TMPAux"

        cEvento := TMPAux->R8_DATA

        TMPAux->(dbCloseArea())
        
        if !Empty(MV_PAR03)
            cQuery := "SELECT * FROM " + RetSqlName("RFQ") + " WHERE D_E_L_E_T_ = '' AND RFQ_PERIOD = '" + MV_PAR03 + "'"

            TCQuery cQuery New Alias "TMPAux"

            IF TMPAux->RFQ_STATUS == "1"
                //Mes aberto - SRC
                cQuery := "SELECT * FROM " + RetSqlName("SRC") + " WHERE D_E_L_E_T_ = '' AND RC_FILIAL = '" + TRB->RA_FILIAL + "' AND RC_MAT = '" + TRB->RA_MAT + "' AND RC_PERIODO = '" + MV_PAR03 + "'"
                TCQuery cQuery New Alias "TMPfol"
                
                While !(TMPfol->(EoF()))
                    IF TMPfol->RC_PD == "101"
                        cBruto := TMPfol->RC_VALOR
                    Elseif TMPfol->RC_PD == "799"
                        cliquido := TMPfol->RC_VALOR
                    Elseif TMPfol->RC_PD == "010" .OR. TMPfol->RC_PD == "011" .OR. TMPfol->RC_PD == "012"
                        cInsalubridade := TMPfol->RC_VALOR
                    Elseif TMPfol->RC_PD == "784"
                        cVT := TMPfol->RC_VALOR
                    Elseif TMPfol->RC_PD == "998"
                        cVA := TMPfol->RC_VALOR
                    Endif

                    TMPfol->(dbSkip())
                Enddo   
                TMPfol->(dbCloseArea())
            Else
                //Mes fechado - SRD
                cQuery := "SELECT * FROM " + RetSqlName("SRD") + " WHERE D_E_L_E_T_ = '' AND RD_FILIAL = '" + TRB->RA_FILIAL + "' AND RD_MAT = '" + TRB->RA_MAT + "' AND RD_PERIODO = '" + MV_PAR03 + "'"
                TCQuery cQuery New Alias "TMPfol"
                
                While !(TMPfol->(EoF()))
                    IF TMPfol->RD_PD == "101"
                        cBruto := TMPfol->RD_VALOR
                    Elseif TMPfol->RD_PD == "799"
                        cliquido := TMPfol->RD_VALOR
                    Elseif TMPfol->RD_PD == "010" .OR. TMPfol->RD_PD == "011" .OR. TMPfol->RD_PD == "012"
                        cInsalubridade := TMPfol->RD_VALOR
                    Elseif TMPfol->RD_PD == "784"
                        cVT := TMPfol->RD_VALOR
                    Elseif TMPfol->RD_PD == "998"
                        cVA := TMPfol->RD_VALOR
                    Endif

                    TMPfol->(dbSkip())
                Enddo   
                TMPfol->(dbCloseArea())
            Endif 
        Endif

        IF nPos := AT("Diurno",cCargo) <> 0
            nHoras := 84
            cInicial := "07:00"
            cFinal := "19:00"
        ELSEIF nPos := AT("Noturno",cCargo) <> 0
            nHoras := 84
            cInicial := "19:00"
            cFinal := "07:00"
        Else
            nHoras := 220
            cInicial := "07:00"
            cFinal := "17:00"
        Endif

        oFWMsExcel:AddRow("Afastamentos","Colaborador",{;
                                    nSequencia,;
                                    TRB->RA_MAT,;
                                    TRB->RA_NOMECMP,;
                                    SUBSTRING(TRB->RA_NASC,7,2) + '/' + SUBSTRING(TRB->RA_NASC,5,2) + '/' + SUBSTRING(TRB->RA_NASC,1,4),;
                                    TRB->RA_LOGRDSC,;
                                    TRB->RA_LOGRNUM,;
                                    TRB->RA_BAIRRO,;
                                    cCidade,;
                                    SUBSTRING(TRB->RA_CEP,1,5) + "-" + SUBSTRING(TRB->RA_CEP,6,3),;
                                    Fone,;
                                    SUBSTRING(TRB->RA_RG,1,2) + "." + SUBSTRING(TRB->RA_RG,3,3) + "." + SUBSTRING(TRB->RA_RG,6,3) + "-" + SUBSTRING(TRB->RA_RG,9,2) + " - " + TRB->RA_RGUF,;
                                    SUBSTRING(TRB->RA_CIC,1,3) + "." + SUBSTRING(TRB->RA_CIC,4,3) + "." + SUBSTRING(TRB->RA_CIC,7,3) + "-" + SUBSTRING(TRB->RA_CIC,10,2),;
                                    TRB->RA_PIS,;
                                    TRB->RA_BANCO,;
                                    TRB->RA_AGENCIA,;
                                    TRB->RA_CTDEPSA,;
                                    cTpConta,;
                                    SUBSTRING(TRB->RA_ADMISSA,7,2) + "/" + SUBSTRING(TRB->RA_ADMISSA,5,2) + "/" + SUBSTRING(TRB->RA_ADMISSA,1,4),;
                                    cCusto,;
                                    cDepto,;
                                    TRB->RA_SITFOLH,;
                                    SUBSTRING(cEvento,7,2) + "/" + SUBSTRING(cEvento,5,2) + "/" + SUBSTRING(cEvento,1,4),;
                                    Space(10),;
                                    cCargo,;
                                    nHoras,;
                                    cInicial,;
                                    cFinal,;
                                    cBruto,;
                                    cliquido,;
                                    cInsalubridade,;
                                    cVT,;
                                    cVA,;
                                    cliquido + cVA + cVT})

        TRB->(DbSkip())
        nSequencia++
    EndDo

    //Ativando o arquivo e gerando o xml
    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo)
        
    oExcel := MsExcel():New()       //Abre uma nova conexão com Excel
    oExcel:WorkBooks:Open(cArquivo) //Abre uma planilha
    oExcel:SetVisible(.T.)          //Visualiza a planilha
    oExcel:Destroy()                //Encerra o processo do gerenciador de tarefas
     
    TRB->(DbCloseArea())
    RestArea(aArea)
Return
