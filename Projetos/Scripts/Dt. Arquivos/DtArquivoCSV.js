const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const diretorio = 'D:\\Clientes\\0046-CLAW\\';

// Cria um novo livro de trabalho (workbook) e planilha (worksheet)
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Dados');

// Define as colunas da planilha
worksheet.columns = [
    { header: 'Arquivo', key: 'arquivo', width: 30 },
    { header: 'Caminho Completo', key: 'caminho', width: 50 },
    { header: 'Data de Modificação', key: 'data', width: 20 }
];

// Função para percorrer o diretório recursivamente
function listarArquivos(diretorioBase) {
    const arquivos = fs.readdirSync(diretorioBase);

    arquivos.forEach(arquivo => {
        const caminhoCompleto = path.join(diretorioBase, arquivo);
        const stats = fs.statSync(caminhoCompleto);

        if (stats.isDirectory()) {
            // Se for diretório, chama a função recursivamente
            listarArquivos(caminhoCompleto);
        } else {
            // Se for arquivo, adiciona ao Excel
            worksheet.addRow({
                arquivo: arquivo,
                caminho: caminhoCompleto,
                data: stats.mtime.toISOString()
            });
        }
    });
}

// Chama a função para iniciar o processo
listarArquivos(diretorio);

// Salva o arquivo Excel
workbook.xlsx.writeFile('dados_modificacao.xlsx')
    .then(() => console.log('Exportação concluída com sucesso!'))
    .catch(err => console.error('Erro ao salvar o arquivo Excel:', err));