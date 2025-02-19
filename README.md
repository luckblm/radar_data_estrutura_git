# Radar de Indicadores do Estado do Pará

## Descrição

Este repositório contém scripts em R para a geração de planilhas Excel utilizadas na entrada de dados e na criação de uma base de dados unificada sobre os indicadores do produto **Radar de Indicadores do Estado do Pará**. O produto apresenta 51 indicadores para o Estado, desagregados por **Região de Integração** e **municípios**.

## Funcionalidades

- Cria planilhas Excel onde, dependendo do indicador, são geradas abas individuais para cada categoria.
- Permite a alimentação manual dos arquivos para posterior integração na base de dados central.
- Facilita a análise e visualização dos dados nos sistemas e dashboards desenvolvidos em **R, Power BI** ou para consultas diretas.

## Estrutura do Projeto

O repositório contém:

- **Scripts em R** para a geração dos arquivos base.
- **Tabelas modelo** para cada Região de Integração.
- **Exemplo de uso** para a estrutura de dados e a criação das planilhas Excel.

## Dependências

Para executar os scripts, é necessário ter instalados os seguintes pacotes no R:

```r
install.packages(c("tidyverse", "openxlsx"))
```

## Estrutura das Tabelas

Cada arquivo gerado contém os seguintes campos:

- **tematica**: Tema do indicador
- **indicador**: Nome do indicador
- **regiao**: Região de Integração
- **localidade**: Nome do município ou região
- **categoria1 a categoria5**: Categorias do indicador

## Contribuição

Se deseja contribuir para este projeto:

1. Fork este repositório
2. Crie uma branch para suas modificações (`git checkout -b minha-mudanca`)
3. Faça commit das alterações (`git commit -m 'Minha contribuição'`)
4. Envie para o repositório remoto (`git push origin minha-mudanca`)
5. Abra um Pull Request

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

