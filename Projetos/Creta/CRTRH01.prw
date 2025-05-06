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
    Local aArea         := GetArea()
    Local cQuery        := ""
    Local oFWMsExcel    
    Local oExcel
    Local cArquivo      := GetTempPath()+'zTstExc1.xml'
    Local nSequencia    := 1
    Local cCidade       := ""
    Local Fone          := ""
    Local cTpConta      := ""
    Local cCusto        := ""
    Local cDepto        := ""
    Local cCargo        := ""
    Local cEvento      := ""
 
    //Pegando os dados
    cQuery := " SELECT TOP 100 RA_FILIAL,RA_MAT,RA_NOMECMP,CONVERT(CHAR(10),RA_NASC,103) AS RA_NASC,RA_LOGRDSC,RA_LOGRNUM,RA_BAIRRO,RA_CODMUN,RA_CEP, RA_TELEFON AS RA_FONE, RA_DDDFONE AS RA_DDD"
    cQuery += " ,RA_RG,RA_CIC,RA_PIS,SUBSTRING(RA_BCDEPSA,1,3) AS RA_BANCO, SUBSTRING(RA_BCDEPSA,4,10) AS RA_AGENCIA,RA_RGUF "
    cQuery += " ,RA_CTDEPSA, RA_TPCTSAL "
    cQuery += " ,RA_ADMISSA,RA_CC,RA_DEPTO"
    cQuery += " ,CASE WHEN RA_SITFOLH = ' ' THEN 'Ativo' WHEN RA_SITFOLH = 'D' THEN 'Desligado' WHEN RA_SITFOLH = 'F' THEN 'Ferias' ELSE 'Nao informado' END AS RA_SITFOLH "
    cQuery += " ,RA_CODFUNC"
    cQuery += " FROM "+RetSqlName("SRA")+" SRA "
    cQuery += " WHERE SRA.D_E_L_E_T_ = '' "  
    cQuery += " ORDER BY RA_MAT "

   //salva o código sql na pasta TEMP para consultas no seu SGBD
    memoWrite("\TEMP\RELSRA.sql",cQuery)     
    TCQuery cQuery New Alias "TRB"

    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New()
    oFWMsExcel:AddworkSheet("Gestao dos colaboradores")
                   
    //Criando a Tabela
    oFWMsExcel:AddTable("Gestao dos colaboradores","Colaborador")
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Nº",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Matricula",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Nome completo",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Data de Nascimento",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Endereço",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Num",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Bairro",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Cidade",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Cep",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","DDD + Telefone",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","RG",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","CPF",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","PIS",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Banco",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Agencia",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Conta",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Tipo conta",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Data de admissão",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Centro de custo",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Departamento",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Situacao",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Data do envento",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Substituto",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Cargo",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Carga horaria",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Inicial",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Final",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Sal.Liquido",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Insalubridade",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","VT",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","VR",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Total",1)
    

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

        oFWMsExcel:AddRow("Gestao dos colaboradores","Colaborador",{;
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
                                                        Space(10),;
                                                        Space(10),;
                                                        Space(10),;
                                                        Space(10),;
                                                        Space(10),;
                                                        Space(10),;
                                                        Space(10),;
                                                        Space(10)})

        TRB->(DbSkip())
        nSequencia++
    EndDo
     
    //Ativando o arquivo e gerando o xml
    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo)
         
    //Abrindo o excel e abrindo o arquivo xml
    oExcel := MsExcel():New()             //Abre uma nova conexão com Excel
    oExcel:WorkBooks:Open(cArquivo)     //Abre uma planilha
    oExcel:SetVisible(.T.)                 //Visualiza a planilha
    oExcel:Destroy()                        //Encerra o processo do gerenciador de tarefas
     
    TRB->(DbCloseArea())
    RestArea(aArea)
Return
