# Caminho da pasta raiz
$pastaRaiz = "D:\Clientes\0046-CLAW"

# Lista todos os arquivos na pasta e subpastas
$arquivos = Get-ChildItem -Path $pastaRaiz -File -Recurse

# Cria um objeto com caminho e data de modificação
$dados = $arquivos | Select-Object @{Name="Caminho";Expression={$_.Name}},
                                   @{Name="DataModificacao";Expression={$_.LastWriteTime}}

# Exporta para CSV
$dados | Export-Csv -Path "D:\Clientes\0046-CLAW\saida.csv" -Encoding UTF8 -NoTypeInformation