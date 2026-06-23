# Automação Tickets ABEC 🤖🎟️

Automação destinada para a captura de tickets no Citsmart da mantenedora **ABEC**, desenvolvida com Node.js. A coleta é feita diretamente na própria ferramenta, garantindo a extração automatizada das informações necessárias.

O processo inicia com o acesso ao Citsmart e autenticação do usuário, seguido da coleta dos números dos tickets. Registros já processados são desconsiderados com base no arquivo `ticketsIgnorados.txt`. Em seguida, cada ticket é acessado individualmente para a extração de dados como data de vencimento, forma de pagamento e valor. A partir de campos como **Número do pedido**, o sistema identifica se o ticket está vinculado a uma **OC**, **contrato** ou se requer a criação de uma **regularização**. Quando necessário, os dados são organizados no padrão de regularização e registrados no arquivo `regularizações.txt`, utilizando como apoio as informações presentes em `contasFinanceiras.txt`.

Por fim, todos os tickets processados são consolidados em uma **planilha**, que é gerada e aberta automaticamente ao final da execução.

---
## Variáveis de ambiente

Para executar este projeto é necessário adicionar as seguintes variáveis no seu arquivo `.env`, referentes ao login no Citsmart:

`USUARIO` e `SENHA`.

---
## Execução

Para executar a automação, é necessário instalar todos os arquivos presentes neste repositório. Após isso, abra o editor de código (ou similar) na pasta do projeto, certifique-se de ter o Node.js instalado e execute o seguinte comando no terminal para instalar as dependências:

```bash
    npm install
```

Após a instalação, execute o comando abaixo no terminal para iniciar a automação:

```bash
    npm run autoTicketsABEC.js
```

---
## Autores

- *[@euisaquevenancio](https://euisaquevenancio.github.io/portfolio/) - 19/04/2026*
