# Boletim Econômico e Regulatório 

## O Projeto

O propósito desse projeto é a automação da extração, tratamento e apresentação de dados econômicos e financeiros que afetam os reajustes e revisões tarifarias dos serviços públicos regulados pela AGEPAR.

## Indicadores

O boletim agrega 4 índices de inflação que compõem uma cesta de índices utilizados para o próposito do reajuste tarifário. São eles:

- IPCA
- INPC
- IGP-DI
- IGP-M

Os indicadores são publicados mensalmente pelo IBGE(Instituto Brasileiro de Geografia e Estatistica) e FGV-IBRE(Instituto Brasileiro de Economia da Fundação Getúlio Vargas) e a série histórica podem ser encontrados em seus respectivos domínios.

Além da inflação, o boletim faz o acompanhamento preço do barril de petróleo brent e do combustível Diesel S10. Os dados de referência para a série histórica é publicado no IPEADATA e a série histórica do Diesel S10 é publicado pela ANP(Agência Nacional de Petróleo).

## Dependências

A versão do Python utilizada para o execução do projeto é a 3.9. Para scraping, tratamento, apresentação e manipulação de planilhas, foram utilizadas as seguintes bibliotecas:

    pandas
    numpy
    ssl
    datetime
    urllib
    matplotlib
    openpyxl

Os arquivos .xlsx presentes nesse repositório estão para fins de apresentação de dados em formato tabular. Todavia, não é essencial ao projeto.
