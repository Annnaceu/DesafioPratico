# Projeto de RPA - Web Scraping no Site Magazine Luiza

Este projeto tem como objetivo automatizar a coleta de informações de produtos no site da Magazine Luiza (Magalu) utilizando **Robotic Process Automation (RPA)**. A automação foi construída para extrair dados de notebooks, categorizá-los entre os melhores e piores produtos com base na quantidade de avaliações e, posteriormente, gerar relatórios em Excel e enviar por e-mail.

O código foi desenvolvido com **Python**, utilizando bibliotecas como **Selenium** e **BeautifulSoup** para o scraping de dados. O fluxo de trabalho da automação envolve:

1. **Abertura do navegador** para acessar a página do Magalu.
2. **Coleta de dados de notebooks**, como nome do produto, preço, quantidade de avaliações e nota.
3. **Filtragem dos produtos**, categorizando-os em "Melhores" e "Piores" com base na quantidade de avaliações.
4. **Geração de um arquivo Excel** com três planilhas: "Todos", "Melhores" e "Piores".
5. **Envio de relatório por e-mail** com os dados coletados, atendendo aos requisitos do cliente.

## Tecnologias Utilizadas

- **Python**: Linguagem principal do projeto.
- **Selenium**: Para automação de navegação no site.
- **BeautifulSoup**: Para extração de dados da página.
- **Pandas**: Para manipulação de dados e criação do relatório em Excel.
- **SMTP**: Para envio do relatório por e-mail.

## Como Executar

1. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/Projeto-RPA.git
