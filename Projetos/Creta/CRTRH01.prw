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
 
    //Pegando os dados
    cQuery := " SELECT RA_MAT,RA_NOMECMP,RA_NASC,RA_LOGRDSC,RA_LOGRNUM,RA_BAIRRO,RA_CODMUN,RA_CEP, '(' + RA_DDDFONE + ')' + '-' + RA_TELEFON AS RA_FONE"
    cQuery += " ,RA_RG,RA_CIC,RA_PIS,SUBSTRING(RA_BCDEPSA,1,3) AS RA_BANCO, SUBSTRING(RA_BCDEPSA,4,10) AS RA_AGENCIA "
    cQuery += " ,CASE WHEN RA_CTDEPSA = '1' THEN 'Conta Corrente' WHEN RA_CTDEPSA = '2' THEN 'Conta Poupanca' ELSE 'Nao informado' END AS RA_CTDEPSA "
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
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Matricula",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Nome",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Dt.Nasc",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Logradouro",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Num",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Bairro",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Municipio",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Cep",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Telefone",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","RG",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","CPF",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","PIS",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Banco",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Agencia",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Tipo conta",1)
    oFWMsExcel:AddColumn("Gestao dos colaboradores","Colaborador","Admissao",1)
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
        oFWMsExcel:AddRow("Gestao dos colaboradores","Colaborador",{;
                                                        TRB->RA_MAT,;
                                                        TRB->RA_NOMECMP,;
                                                        TRB->RA_NASC,;
                                                        TRB->RA_LOGRDSC,;
                                                        TRB->RA_LOGRNUM,;
                                                        TRB->RA_BAIRRO,;
                                                        TRB->RA_CODMUN,;
                                                        TRB->RA_CEP,;
                                                        TRB->RA_FONE,;
                                                        TRB->RA_RG,;
                                                        TRB->RA_CIC,;
                                                        TRB->RA_PIS,;
                                                        TRB->RA_BANCO,;
                                                        TRB->RA_AGENCIA,;
                                                        TRB->RA_CTDEPSA,;
                                                        TRB->RA_ADMISSA,;
                                                        TRB->RA_CC,;
                                                        TRB->RA_DEPTO,;
                                                        TRB->RA_SITFOLH,;
                                                        Space(10),;
                                                        Space(10),;
                                                        TRB->RA_CODFUNC,;
                                                        Space(10),;
                                                        Space(10),;
                                                        Space(10),;
                                                        Space(10),;
                                                        Space(10),;
                                                        Space(10),;
                                                        Space(10),;
                                                        Space(10)})

        //Pulando Registro
        TRB->(DbSkip())
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
