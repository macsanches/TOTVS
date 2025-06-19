# Caminho da pasta raiz
$pastaRaiz = "D:\Clientes\0046-CLAW"

# Lista todos os arquivos na pasta e subpastas
$arquivos = Get-ChildItem -Path $pastaRaiz -File -Recurse

# Para cada arquivo, pega nome e data de modificação
$resultado = foreach ($arquivo in $arquivos) {
    "{0}`t{1}" -f $arquivo.FullName, $arquivo.LastWriteTime
}

# Salva tudo em um .txt
$resultado | Out-File "D:\Clientes\0046-CLAW\relatorio-datas.txt" -Encoding UTF8
