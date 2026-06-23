/*
    @euisaquevenancio - 23/06/2026
    Automação para captura de tickets no Citsmart da mantenedora ABEC, desenvolvido com Node.js.
    
    Instalando todas bibliotecas de uma vez via terminal:
    npm install

    Instalando as bibliotecas manualmente via terminal:
    npm install dotenv
    npm install axios@1.4.0 cheerio@1.0.0-rc.12
    npm install puppeteer
    npm install exceljs

    Executando o código via terminal:
    node autoTicketsABEC.js
*/

// Bibliotecas
require("dotenv").config();
const fs = require("fs"); // Manipulação de arquivos
const puppeteer = require("puppeteer"); // Acesso ao navegador (Chrome)
const ExcelJS = require("exceljs"); // Manipulação do Excel com JavaScript
const path = require("path"); // Uso de caminho de arquivos 
const { exec } = require("child_process"); // Permite executar comandos do Sistema Operacional

// Login Citsmart
const usuarioCitsmart = process.env.USUARIO;
const senhaCitsmart = process.env.SENHA;

async function main() {
    const horarioInicio = new Date().toLocaleTimeString('pt-BR'); 
    // Executando o navegador
    // const navegador = await puppeteer.launch({ headless: false });
    const navegador = await puppeteer.launch({
        executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        headless: !true
    });
    const pagina = await navegador.newPage();
    pagina.setDefaultTimeout(30000);

    let listaTicketsIgnorados = fs.readFileSync("dados/ticketsIgnorados.txt", "utf-8");
    listaTicketsIgnorados.split("\n").map((tI) => tI.trim()).filter((tI) => tI.length > 0);

    // Acessa a página de login do Citsmart
    await pagina.goto(
        "https://servicos.maristabrasil.org/citsmart/webmvc/login#/ec?idExperienceCenter=3q",
        { waitUntil: "networkidle2" }
    );
    
    // Realiza login no Citsmart
    await pagina.waitForSelector("#user_login");
    await pagina.waitForSelector("#password");
    await pagina.type("#user_login", usuarioCitsmart);
    await pagina.type("#password", senhaCitsmart);
    await pagina.keyboard.press("Enter");
    await pagina.waitForNavigation({ waitUntil: "networkidle2" });
    await new Promise((r) => setTimeout(r, 2000));

    const listaTicketsExcel = [];
    const listaTicketsNaoABEC = [];
    const listaTicketsIgnoradosExecucao = [];
    const listaRegularizacoes = [];
    let listaContasFinanceiras = fs.readFileSync("dados/contasFinanceiras.txt", "utf-8").split("\n").map((cF) => cF.trim()).filter((cF) => cF.length > 0);

    // Captura tickets das duas pesquisas
    const listaTicketsPesquisaNotaFiscalEletronica = await capturarTickets(pagina, "NOTA FISCAL ELETRÔNICA", listaTicketsIgnorados);
    const listaTicketsPesquisaNotaDeTerceiros = await capturarTickets(pagina, "NOTA DE TERCEIROS", listaTicketsIgnorados);

    listaTicketsIgnoradosExecucao.push(...listaTicketsPesquisaNotaFiscalEletronica.ticketsIgnoradosExecucao);
    listaTicketsIgnoradosExecucao.push(...listaTicketsPesquisaNotaDeTerceiros.ticketsIgnoradosExecucao);
    
    const pesquisas = [
        { termo: "Eletrônica", ...listaTicketsPesquisaNotaFiscalEletronica },
        { termo: "Terceiros", ...listaTicketsPesquisaNotaDeTerceiros },
    ];

    let contadorTicket = 1;

    for (const pesquisa of pesquisas) {
        const { todosTickets, datasEntradas, termo } = pesquisa;

        // Verificando cada um dos tickets
        for (let i = 0; i < todosTickets.length; i++) {
            const ticket = todosTickets[i];
            const entrada = datasEntradas[i];
            // Caso o ticket não exista ou ele esteja na lista de tickets que devem ser ignorados, pula esse ticket e segue o fluxo
            if (!ticket || listaTicketsIgnorados.includes(ticket) || ticket == "485540") {
                if (ticket == "485540") {
                    if (contadorTicket < 10) {
                        console.log(`❌ #0${contadorTicket} | ${ticket} | Entrada: ${entrada} | Ignorado`);
                    } else {
                        console.log(`❌ #${contadorTicket} | ${ticket} | Entrada: ${entrada} | Ignorado`);
                    }
                    contadorTicket++;
                }
                continue;
            }

            // Verifica se a barra de pesquisa esta disponível
            await pagina.waitForFunction(() => {
                const elementoBarraDePesquisa = document.querySelector("#pesquisaSolicitacao");
                return elementoBarraDePesquisa && !elementoBarraDePesquisa.disabled;
            }, { timeout: 20000 });

            // Seleciona a barra e realiza a pesquisa do ticket atual
            await pagina.focus("#pesquisaSolicitacao");
            await pagina.click("#pesquisaSolicitacao", { clickCount: 3 });
            await pagina.keyboard.press("Backspace");
            await pagina.type("#pesquisaSolicitacao", ticket);
            await pagina.keyboard.press("Enter");
            await pagina.keyboard.press("Enter");
            await new Promise((r) => setTimeout(r, 4000));

            // Captura a fila que o ticket pertence - NOTA FISCAL ELETRÔNICA ou TERCEIROS
            const fila = await pagina.evaluate(() => {
                const elementoFila = document.querySelector("div.tableless-td.ellipsis.solicitacao.ng-binding");
                if (elementoFila) {
                    return elementoFila.textContent.trim().toUpperCase();
                }
                
                return null;
            });

            // Tentativa de acessar o ticket
            try {
                await pagina.waitForSelector(".request-id", { visible: true, timeout: 20000 });
                await pagina.click(".request-id", { clickCount: 2 });
            } catch (err) {
                console.log("Não foi possível clicar no ticket: ", err.message);
            }

            // Aguardando a tela do ticket abrir
            await new Promise((r) => setTimeout(r, 4000));
            // Recarrega a página para tentar evitar falhas
            await pagina.reload({ waitUntil: "networkidle2" });
            await new Promise((r) => setTimeout(r, 8000));

            const erro = await pagina.waitForSelector("#div-form-builder > div > span.error", { visible: true, timeout: 2000 }).catch(() => null);
            if (erro) {
                await pagina.reload({ waitUntil: "domcontentloaded" });
                await new Promise((r) => setTimeout(r, 4000));
            }

            // Captura a mantenedora do ticket
            const mantenedora = await pagina.evaluate(() => {
                const seletoresMantenedora = [
                    "#testelancarnotasmbPage > div:nth-child(2) > div.col-md-2 select",
                    "#testelancarnotasmb\\.mantenedora > div > select",
                    "#formulariosGerais\\.Mantenedora > div > select"
                ];

                let elementoMantenedora = null;
                for (const seletor of seletoresMantenedora) {
                    elementoMantenedora = document.querySelector(seletor);
                    if (elementoMantenedora) {
                        break;
                    }
                }

                const mantenedoraSelecionada = Array.from(elementoMantenedora.querySelectorAll("option[selected]"));
                return mantenedoraSelecionada[1]?.textContent.trim() || null;
            });

            // Se a mantenedora não for ABEC
            if (mantenedora !== "ABEC") {
                // Adiciona o número do ticket e a mantenedora na lista
                listaTicketsNaoABEC.push({ ticket, mantenedora, entrada });

                // Retorna para a página de pesquisa e segue o fluxo
                await pagina.goto(
                    "https://servicos.maristabrasil.org/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load#/",
                    { waitUntil: "networkidle2" }
                );
                await new Promise((r) => setTimeout(r, 3500));
                continue;
            }

            // Captura a data de vencimento do ticket
            const vencimento = await pagina.evaluate(() => {
                const seletoresVencimento = [
                    "#testelancarnotasmb\\.data input",
                    "input#notasDeTerceiros_MB\\.data_vencimento"
                ];

                for (const seletor of seletoresVencimento) {
                    const elementoVencimento = document.querySelector(seletor);
                    if (!elementoVencimento) {
                        continue;
                    }

                    if (elementoVencimento.value?.trim()) {
                        return elementoVencimento.value?.trim();
                    } else {
                        return elementoVencimento.textContent?.trim();
                    }
                }

                return null;
            });

            // Captura a forma de pagamento do ticket
            let formaPagamento = await pagina.evaluate(() => {
                const seletoresFormaDePagamento = [
                    "#testelancarnotasmb\\.vencimentopagamento select",
                    "#notasDeTerceiros_MB\\.tipo_pagamento > div > select"
                ];

                for (const seletor of seletoresFormaDePagamento) {
                    const elementoFormaDePagamento = document.querySelector(seletor);
                    if (elementoFormaDePagamento) {
                        const formaDePagamentoSelecionada = elementoFormaDePagamento.querySelector("option[selected]");
                        
                        if (formaDePagamentoSelecionada.textContent.trim() == "--- SELECIONE ---") {
                            return "VERIFICAR";
                        }

                        if (formaDePagamentoSelecionada.textContent.trim().toUpperCase() == "DEPÓSITO" || formaDePagamentoSelecionada.textContent.trim().toUpperCase() == "DEPOSITO") {
                            return "CRÉDITO";
                        }
                        
                        return formaDePagamentoSelecionada.textContent.trim().toUpperCase();
                    }
                }

                return null; // se nenhum select existir
            });

            // Captura número do pedido do ticket
            const numeroDoPedido = await pagina.evaluate(() => {
                const elementoNumeroDoPedido = document.querySelector("#testelancarnotasmb\\.numeropedido input, #testelancarnotasmb\\.numeropedido2 input, #testelancarnotasmb\\.npedido input");

                if (elementoNumeroDoPedido) {
                    return elementoNumeroDoPedido.value?.trim();
                }
                return null;
            });

            // Captura o tipo da nota
            let tipoNota = await pagina.evaluate(() => {
                const elementoTipoNota = document.querySelector("#testelancarnotasmb\\.tiponata > div > select");
                if (elementoTipoNota) {
                    const tipoNotaSelecionada = elementoTipoNota.querySelector("option[selected]");

                    return tipoNotaSelecionada.textContent.trim().toUpperCase();       
                }
                
                return null;
            });

            // Captura o valor do ticket
            const valorNota = await pagina.evaluate(() => {
                const elementoValor = document.querySelector(
                    "#notasDeTerceiros_MB\\.valor_nota input, #testelancarnotasmb\\.valornotafiscalinicial input"
                );

                if (elementoValor) {
                    return elementoValor.value?.trim();
                }
                return null;
            });

            // Captura o tipo do ticket
            const tipo = await pagina.evaluate((termo, numeroDoPedido) => {
                if (termo == "Terceiros") {
                    return "REGULARIZAÇÃO";
                } else {
                    // REGULARIZAÇÃO
                    if (!numeroDoPedido || numeroDoPedido == "-" || numeroDoPedido.toUpperCase() == "X" || numeroDoPedido == "o" || numeroDoPedido == "0" || numeroDoPedido.toUpperCase() == "REGULARIZAÇÃO" || numeroDoPedido.toUpperCase() == "REGULARIZACAO" || numeroDoPedido.toUpperCase() == "CRIAR" || numeroDoPedido.toUpperCase() == "NÃO TEM" || numeroDoPedido == "000000") {
                        return "REGULARIZAÇÃO";
                    }
                    // OC
                    if (/^\d{8}$/.test(numeroDoPedido) || numeroDoPedido.length == 11 || numeroDoPedido.length == 10 || numeroDoPedido.toUpperCase().includes("OC")) {
                        return "OC";
                    }
                    // CONTRATO (letras, números, "/" e pode conter "-")
                    if ((numeroDoPedido.toUpperCase() == "CONTRATO" || numeroDoPedido.toUpperCase().includes("CONTRATO") || (numeroDoPedido.length >= 13 && numeroDoPedido.length <= 18)) && /[A-Za-z]/.test(numeroDoPedido) && /\d/.test(numeroDoPedido)) {
                        return "CONTRATO";
                    }
                    //
                    if (document.querySelector("#testelancarnotasmb\\.centro_custo input") && document.querySelector("#testelancarnotasmb\\.contamb input")) {
                        return "REGULARIZAÇÃO";
                    }

                    return "VERIFICAR";
                }
            }, termo, numeroDoPedido);

            let observacao = "";
            // Captura a situação do ticket
            const situacao = await pagina.evaluate(() => {
                const elementoSituacao = document.querySelector(".situacao > span");
                
                if (elementoSituacao) {
                    // Retorna "Fechada" se tiver a classe badge-default
                    if (elementoSituacao.classList.contains("badge-default")) {
                        observacao = "Fechada";
                        return "Fechada";
                    }
                }
                return null;                
            });

            // Captura o e-mail do solicitante do ticket
            let emailSolicitante = await pagina.evaluate(() => {
                const elementoEmailSolicitante = document.querySelector(".requester-information-value");
                if (elementoEmailSolicitante) {
                    return elementoEmailSolicitante.textContent.trim();
                }
                return null;
            });

            // Captura o nome do solicitante do ticket
            let nomeSolicitante = await pagina.evaluate(() => {
                const elementoNomeSolicitante = document.querySelector(".requester-information-title.ng-binding");
                if (elementoNomeSolicitante) {
                    return elementoNomeSolicitante.textContent.trim().toUpperCase();
                }
                return null;
            });

            // Emails dos solicitantes do RH
            const emailsRH = [
                "jaqueline.macedo@maristabrasil.org",
                "larissa.reis@maristabrasil.org",
                "beatriz.ricardo@maristabrasil.org",
                "karine.zorek@maristabrasil.org",
                "jennifer.soares@maristabrasil.org",
                "paloma.engels@maristabrasil.org"
            ];

            if (emailsRH.includes(emailSolicitante)) {
                observacao = "RH";
            }

            // Emails atrelados a aprovadora Leticia Castilhos - que não possui acesso ao portal e, por isso, não pode aprovar regularizações
            const emailsLeticia = [
                "caroline.borba@maristabrasil.org",
                "leticia.castilhos@maristabrasil.org"
            ];

            if (emailsLeticia.includes(emailSolicitante)) {
                observacao = "APROVADORA DO CR 35114 NÃO CONSEGUE APROVAR - PERGUNTAR SE PRECISA DE OUTRO CR";
            }

            // Emails atrelados a aprovadora Luana Alvarenga - que não possui acesso ao portal e, por isso, não pode aprovar regularizações
            const emailsLuana = [
                "tatiane.michalovicz@maristabrasil.org",
                "luana.braga@maristabrasil.org"
            ];

            if (emailsLuana.includes(emailSolicitante)) {
                observacao = "CRIAR REGULARIZAÇÃO COM O CR 35113 E NO LANÇAMENTO PASSAR PARA 35344";
            }

            // Captura a descrição  do ticket
            let descricaoTicket = await pagina.evaluate(() => {
                const elementoDescricao = document.querySelector("#service-request-view > div > div > div > div.service-request-wrapper > div > div > div.service-request-content.clearfix.s12 > div:nth-child(2) > div.panel.panel-default.service-request-panel-details > div > fieldset > div:nth-child(2) > div > div > div");
                if (elementoDescricao) {
                    return elementoDescricao.textContent
                            .replace(/\s*\n+\s*/g, ' ')
                            .replace(/\s+/g, ' ')
                            .trim();
                }
                return null;
            });

            // Captura a unidade do ticket
            unidade = await pagina.evaluate((termo) => {
                if (termo === "Terceiros") {
                    return "53";
                }

                const elementoUnidade = document.querySelector("#testelancarnotasmb\\.unidade > div > select");

                if (elementoUnidade) {
                    const unidadeSelecionada = elementoUnidade.selectedOptions[0];

                    if (unidadeSelecionada && unidadeSelecionada.textContent) {
                        if (unidadeSelecionada.textContent.trim().substring(0, 2) == "--") {
                            return "";
                        }

                        if (parseInt(unidadeSelecionada.textContent.trim().substring(0, 2)) > 9) {
                            return unidadeSelecionada.textContent.trim().substring(0, 2);
                        } else {
                            return unidadeSelecionada.textContent.trim().substring(0, 1);
                        }
                    }
                }

                return "";
            }, termo);

            // Capturando as informações das regularizações
            if (tipo == "REGULARIZAÇÃO") {
                // Captura o centro de custos do ticket
                const centroDeCustos = await pagina.evaluate((termo) => {
                    if (termo == "Terceiros") {
                        return 35119;
                    }

                    const elementoCentroDeCusto = document.querySelector(
                        "#testelancarnotasmb\\.centro_custo input"
                    );

                    if (elementoCentroDeCusto) {
                        if (elementoCentroDeCusto.value?.trim() == "35114" || elementoCentroDeCusto.value?.trim() == 35114) {
                            return "35114 (PERGUNTAR POR OUTRO CR - APROVADORA NÃO TEM ACESSO)";
                        }

                        return elementoCentroDeCusto.value?.trim() || null;
                    }

                    return null;
                }, termo);

                // Captura a conta financeira do ticket
                const contaFinanceira = await pagina.evaluate((termo, listaContasFinanceiras) => {
                    if (termo === "Terceiros") {
                        return "2141 - PUBLICIDADE";
                    }

                    const elementoContaFinanceira = document.querySelector("#testelancarnotasmb\\.contamb input");
                    let conta = elementoContaFinanceira.value?.trim().toUpperCase();

                    if (conta === "FORMAÇÃO E DESENVOLVIMENTO") {
                        return "2041 - CURSOS E TREINAMENTOS";
                    }

                    if (conta === "EVENTO INTERNO" || conta === "EVENTOS INTERNOS" || conta === "EVENTOS INTERNO") {
                        return "2045 - EVENTOS INTERNOS";
                    }

                    if (conta === "01-35111") {
                        return "2030 - ASSISTÊNCIA MÉDICA E ODONTOLÓGICAS";
                    }

                    conta = listaContasFinanceiras.find(c => {
                        const [numero, descricao] = c.split(" - ");
                        return (
                            conta.includes(numero) ||
                            conta.includes(descricao) ||
                            numero.includes(conta) ||
                            descricao.includes(conta)
                        );
                    });

                    if (conta == undefined || conta == "undefined" || conta == "" || conta == false) {
                        return elementoContaFinanceira.value?.trim().toUpperCase();
                    }

                    return conta;
                }, termo, listaContasFinanceiras);

                // Captura o CNPJ do fornecedor do ticket
                const cnpj = await pagina.evaluate((termo) => {
                    if (termo == "Terceiros") {
                        const elementoCNPJ = document.querySelector("#notasDeTerceiros_MB\\.informe_cnpj");

                        if (elementoCNPJ) {
                            return elementoCNPJ.value?.trim() || "";
                        }
                    }

                    return "";
                }, termo);

                // Captura o número da nota do ticket
                const numeroNota = await pagina.evaluate((termo) => {
                    if (termo == "Terceiros") {
                        const elementoNumeroNota = document.querySelector(
                            "#notasDeTerceiros_MB\\.numero_nota input"
                        );

                        if (elementoNumeroNota) {
                            return elementoNumeroNota.value?.trim() || "";
                        }
                    }

                    return "";
                }, termo);

                // Reunindo os dados da regularização do ticket
                const regularizacao = {
                    nome: nomeSolicitante,
                    email: emailSolicitante,
                    ticket: ticket,
                    numero: numeroNota,
                    valor: valorNota,
                    vencimento: vencimento,
                    descricaoTicket: descricaoTicket,
                    centroDeCustos: centroDeCustos,
                    contaFinanceira: contaFinanceira,
                    cnpj: cnpj,
                    unidade: unidade,
                    formaDePagamento: formaPagamento,
                };

                // Adicioando a regularização na lista de regularizações
                listaRegularizacoes.push(regularizacao);
            }

            // Reunindo as informações do ticket e adicionando-as na lista de tickets que vão para o Excel
            listaTicketsExcel.push({
                TICKET: ticket,
                ENTRADA: entrada,
                VENCIMENTO: vencimento,
                "FORMA PAGAMENTO": formaPagamento,
                VALOR: valorNota,
                "TIPO": tipo,
                "OBSERVAÇÃO": observacao,
                "Nº PEDIDO": numeroDoPedido,
                "DESCRIÇÃO": descricaoTicket,
                UNIDADE: unidade,
                "SOLICITANTE": emailSolicitante,
                "TIPO DE NOTA": tipoNota,
                FILA: fila,
            });

            if (contadorTicket < 10) {
                if (fila == "NOTA FISCAL ELETRÔNICA") {
                    console.log(`✅ #0${contadorTicket} | ${ticket} | Entrada: ${entrada} | Vencimento: ${vencimento} | ${formaPagamento} | Valor: ${valorNota} | Tipo: ${tipo} | UN: ${unidade} | FILA: ELETRÔNICA`);
                } else {
                    console.log(`✅ #0${contadorTicket} | ${ticket} | Entrada: ${entrada} | Vencimento: ${vencimento} | ${formaPagamento} | Valor: ${valorNota} | Tipo: ${tipo} | UN: ${unidade} | FILA: TERCEIROS`);
                }
            } else {
                if (fila == "NOTA FISCAL ELETRÔNICA") {
                    console.log(`✅ #${contadorTicket} | ${ticket} | Entrada: ${entrada} | Vencimento: ${vencimento} | ${formaPagamento} | Valor: ${valorNota} | Tipo: ${tipo} | UN: ${unidade} | FILA: ELETRÔNICA`);
                } else {
                    console.log(`✅ #${contadorTicket} | ${ticket} | Entrada: ${entrada} | Vencimento: ${vencimento} | ${formaPagamento} | Valor: ${valorNota} | Tipo: ${tipo} | UN: ${unidade} | FILA: TERCEIROS`);
                }
            }

            contadorTicket++;

            await pagina.goto(
                "https://servicos.maristabrasil.org/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load#/",
                { waitUntil: "networkidle2" }
            );

            await new Promise((r) => setTimeout(r, 3500));
        }
    }

    // Exibe tickets que não são ABEC
    if (listaTicketsNaoABEC.length > 0) {
        listaTicketsNaoABEC.forEach((t) => {
            if (contadorTicket < 10) {
                console.log(`❌ #0${contadorTicket} | ${t.ticket} | Entrada: ${t.entrada} | Mantenedora: ${t.mantenedora}`);
            } else {
                console.log(`❌ #${contadorTicket} | ${t.ticket} | Entrada: ${t.entrada} | Mantenedora: ${t.mantenedora}`);
            }
            contadorTicket++;
        });
    }

    const horarioFim = new Date().toLocaleTimeString('pt-BR');

    // Acessa a pasta onde os arquivos serão salvos
    const pastaDestino = path.join(__dirname, "dados");

    // Garante que a pasta exista
    if (!fs.existsSync(pastaDestino)) {
        fs.mkdirSync(pastaDestino, { recursive: true });
    }

    // Captura o arquivo excel
    const arquivoExcel = path.join(pastaDestino, "tickets.xlsx");
    
    const workbook = new ExcelJS.Workbook();
    // Se o arquivo existe, lê
    if (fs.existsSync(arquivoExcel)) {
        await workbook.xlsx.readFile(arquivoExcel);
    }

    let sheet = workbook.getWorksheet("Tickets");
    // Se a planilha não estiver estruturada, adiciona o cabeçalho 
    if (!sheet) {
        sheet = workbook.addWorksheet("Tickets");

        sheet.addRow([
            "TICKET",
            "USUÁRIO LANÇ.",
            "ENTRADA",
            "VENCIMENTO",
            "FORMA PAGAMENTO",
            "VALOR",
            "TIPO",
            "OBSERVAÇÃO",
            "Nº PEDIDO",
            "DESCRIÇÃO",
            "UNIDADE",
            "SOLICITANTE",
            "TIPO DE NOTA",
            "FILA"
        ]);
    }

    // Aplica a largura nas colunas
    sheet.columns = [
        { key: "TICKET", width: 8 },
        { key: "USUÁRIO LANÇ.", width: 8 },
        { key: "ENTRADA", width: 8 },
        { key: "VENCIMENTO", width: 8 },
        { key: "FORMA_PAGAMENTO", width: 8 },
        { key: "VALOR", width: 8 },
        { key: "TIPO", width: 8 },
        { key: "OBSERVACAO", width: 8 },
        { key: "PEDIDO", width: 8 },
        { key: "DESCRICAO", width: 8 },
        { key: "UNIDADE", width: 8 },
        { key: "SOLICITANTE", width: 8 },
        { key: "TIPO DE NOTA", width: 8 },
        { key: "FILA", width: 8 }
    ];

    // Adicionando os dados (tickets)
    if (contadorTicket > 1 && listaTicketsExcel.length > 0) {
        listaTicketsExcel.forEach(t => {
            sheet.addRow([
                t.TICKET,
                t["USUÁRIO LANÇ."],
                t.ENTRADA,
                t.VENCIMENTO,
                t["FORMA PAGAMENTO"],
                t.VALOR,
                t.TIPO,
                t.OBSERVAÇÃO,
                t["Nº PEDIDO"],
                t.DESCRIÇÃO,
                t.UNIDADE,
                t.SOLICITANTE,
                t["TIPO DE NOTA"],
                t.FILA
            ]);
        });
    }

    if (sheet.rowCount > 1) {
        // Remove a tabela antiga
        if (sheet.model.tables) {
            sheet.model.tables = [];
        }

        const linhasValidas = sheet.getSheetValues()
            .slice(2) // remove header
            .filter(row => Array.isArray(row)) // Remove undefined
            .map(row => row.slice(1)); // Remove índice fantasma

        sheet.addTable({
            name: "TabelaTickets",
            ref: "A1",
            headerRow: true,
            style: {
                theme: "TableStyleLight1"
            },
            columns: [
                { name: "TICKET" },
                { name: "USUÁRIO LANÇ." },
                { name: "ENTRADA" },
                { name: "VENCIMENTO" },
                { name: "FORMA PAGAMENTO" },
                { name: "VALOR" },
                { name: "TIPO" },
                { name: "OBSERVAÇÃO" },
                { name: "Nº PEDIDO" },
                { name: "DESCRIÇÃO" },
                { name: "UNIDADE" },
                { name: "SOLICITANTE" },
                { name: "TIPO DE NOTA" },
                { name: "FILA" }
            ],
            rows: linhasValidas
        });
    }

    if (contadorTicket > 1 && listaTicketsExcel.length > 0) {
        // Salva o arquivo excel
        await workbook.xlsx.writeFile(arquivoExcel);
    }

    const tempoExecucao = diferencaHoras(horarioInicio, horarioFim);

    if (contadorTicket > 1 && listaTicketsExcel.length > 0) {
        // Regularizações TXT
        const arquivoTxtRegularizacoes = path.join(pastaDestino, "regularizacoes.txt");

        let conteudoExistenteRegularizacoes = "";
        // Se o arquivo já existir, lê o conteúdo
        if (fs.existsSync(arquivoTxtRegularizacoes)) {
            conteudoExistenteRegularizacoes = fs.readFileSync(arquivoTxtRegularizacoes, "utf-8");
        }

        let novoConteudoRegularizacoes = "";
        for (const regularizacaoAtual of listaRegularizacoes) {
            const ticketStr = `TICKET ${regularizacaoAtual.ticket}`;

            // Verifica se o ticket já existe no arquivo
            if (!conteudoExistenteRegularizacoes.includes(ticketStr)) {
                novoConteudoRegularizacoes += `PAGAMENTO SOLICITADOR POR: ${regularizacaoAtual.nome} - ${regularizacaoAtual.email}\n`;
                novoConteudoRegularizacoes += `TICKET ${regularizacaoAtual.ticket}\n`;
                novoConteudoRegularizacoes += `NOTA FISCAL NÚMERO ${regularizacaoAtual.numero} | ${regularizacaoAtual.valor} | ${regularizacaoAtual.vencimento}\n`;
                novoConteudoRegularizacoes += `DESCRIÇÃO ${regularizacaoAtual.descricaoTicket}\n`;
                novoConteudoRegularizacoes += `CR ${regularizacaoAtual.centroDeCustos}\n`;
                novoConteudoRegularizacoes += `CONTA ${regularizacaoAtual.contaFinanceira}\n`;
                novoConteudoRegularizacoes += `CNPJ ${regularizacaoAtual.cnpj || ""}\n`;
                novoConteudoRegularizacoes += `${regularizacaoAtual.unidade} | ${regularizacaoAtual.formaDePagamento}\n\n`;
            }
        }

        // Só escreve se tiver algo novo
        if (novoConteudoRegularizacoes) {
            fs.appendFileSync(arquivoTxtRegularizacoes, novoConteudoRegularizacoes, "utf-8");
        }

        // ticketsIgnorados TXT
        const arquivoTxtTicketsIgnorados = path.join(pastaDestino, "ticketsIgnorados.txt");

        let conteudoExistenteTicketsIgnorados = "";
        // Se o arquivo já existir, lê o conteúdo
        if (fs.existsSync(arquivoTxtTicketsIgnorados)) {
            conteudoExistenteTicketsIgnorados = fs.readFileSync(arquivoTxtTicketsIgnorados, "utf-8");
        }

        let novoConteudoTicketsIgnorados = "";
        for (const ticketAtual of listaTicketsExcel) {
            novoConteudoTicketsIgnorados += `\n${ticketAtual.TICKET}`;
        }

        // Só escreve se tiver algo novo
        if (novoConteudoTicketsIgnorados) {
            fs.appendFileSync(arquivoTxtTicketsIgnorados, novoConteudoTicketsIgnorados, "utf-8");
        }

        // Exibe mensagem com o horário
        const horario = new Date().toLocaleTimeString('pt-BR');
        console.log(`\n✅ Todos os dados foram salvos às ${horario}.`);
    } else {
        const horario = new Date().toLocaleTimeString('pt-BR');
        console.log(`\n🎉 Não há nenhum ticket novo. Finalizado às ${horario}.`);
    }

    console.log(`🤖 Tempo de execução: ${tempoExecucao}.`);
    
    if (contadorTicket > 1 && listaTicketsExcel.length > 0) {
        exec(`start "" "${arquivoExcel}"`);
    }

    await navegador.close();
}

// Captura os números dos tickets com base no termo de pesquisa/fila
async function capturarTickets(pagina, termoPesquisa, ticketsIgnorados) {
    // Acessa a página para pesquisar os tickets
    await pagina.goto(
        "https://servicos.maristabrasil.org/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load#/",
        { waitUntil: "networkidle2" }
    );

    await pagina.reload({ waitUntil: "networkidle2" });
    
    // Verifica se a barra de pesquisa esta disponível
    await pagina.waitForSelector("#pesquisaSolicitacao", { timeout: 30000 });
    await pagina.focus("#pesquisaSolicitacao");
    await new Promise((r) => setTimeout(r, 1000));
    // Seleciona a barra e realiza a pesquisa do ticket atual
    await pagina.click("#pesquisaSolicitacao", { clickCount: 3 });
    await pagina.keyboard.press("Backspace");
    await pagina.type("#pesquisaSolicitacao", termoPesquisa);
    await pagina.keyboard.press("Enter");
    await pagina.keyboard.press("Enter");
    await new Promise((r) => setTimeout(r, 7000));

    let todosTickets = [];
    let datasEntradas = [];
    let ticketsIgnoradosExecucao = []; // apenas novos tickets ignorados
    let contadorPagina = 1;

    while (true) {
        const novosTicketsDaPagina = await pagina.$$("[name=list-item]");

        for (let i = 0; i < novosTicketsDaPagina.length; i++) {
            const ticketData = await pagina.evaluate((elemento) => {
                const elementoNumeroTicket = elemento.querySelector(".request-id");
                const elementoDataDeEntradaTicket = elemento.querySelector(".dataCriacao");
                const elementoFilaTicket = elemento.querySelector(".solicitacao");

                return {
                    ticket: elementoNumeroTicket ? elementoNumeroTicket.textContent.trim() : null,
                    entrada: elementoDataDeEntradaTicket ? elementoDataDeEntradaTicket.textContent.trim().split(" ")[0] : null,
                    solicitacao: elementoFilaTicket ? elementoFilaTicket.textContent.trim() : null,
                };
            }, novosTicketsDaPagina[i]);

            if (!ticketData.ticket) continue;

            // Ignora o ticket se ele não for das filas ABEC
            if (!["Nota Fiscal Eletrônica", "Nota de Terceiros"].includes(ticketData.solicitacao)) {
                ticketsIgnorados.push(ticketData.ticket);
                ticketsIgnoradosExecucao.push(ticketData.ticket);

                console.log(
                `❌ ${ticketData.ticket} | Solicitação: ${ticketData.solicitacao}`
                );

                continue;
            }

            if (ticketsIgnorados.includes(ticketData.ticket)) {
                ticketsIgnoradosExecucao.push(ticketData.ticket);
                continue;
            }

            todosTickets.push(ticketData.ticket);
            datasEntradas.push(ticketData.entrada);
        }

        const avancarDisabled = await pagina.$eval(
            "#button-avancar-pesquisa",
            (btn) => btn.hasAttribute("disabled")
        );

        if (avancarDisabled) break;

        await pagina.click("#button-avancar-pesquisa");

        try {
            await pagina.waitForSelector(".request-id", {
                visible: true,
                timeout: 20000
            });

            await new Promise((r) => setTimeout(r, 5000));
            contadorPagina++;
        } catch (err) {
            console.log("Erro ao avançar página:", err.message);
            break;
        }
    }

    const plural = todosTickets.length > 1 ? "tickets" : "ticket";
    const pluralPag = contadorPagina > 1 ? "páginas" : "página";
    console.log(`✅ ${termoPesquisa}: ${todosTickets.length} ${plural} em ${contadorPagina} ${pluralPag}`);

    return { todosTickets, datasEntradas, ticketsIgnoradosExecucao };
}

// Função para calcular a diferença entre dois horários no formato HH:mm:ss
function diferencaHoras(hora1, hora2) {
    try {
        // Quebra as strings em partes
        const [h1, m1, s1] = hora1.split(':').map(Number);
        const [h2, m2, s2] = hora2.split(':').map(Number);

        // Cria objetos Date no mesmo dia
        const dataBase = new Date();
        const date1 = new Date(dataBase.getFullYear(), dataBase.getMonth(), dataBase.getDate(), h1, m1, s1 || 0);
        const date2 = new Date(dataBase.getFullYear(), dataBase.getMonth(), dataBase.getDate(), h2, m2, s2 || 0);

        // Calcula a diferença em milissegundos
        let diffMs = date2 - date1;

        // Se negativo, inverte
        const negativo = diffMs < 0;
        diffMs = Math.abs(diffMs);

        // Converte para horas, minutos e segundos
        const horas = Math.floor(diffMs / (1000 * 60 * 60));
        const minutos = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));
        const segundos = Math.floor((diffMs % (1000 * 60)) / 1000);

        return `${negativo ? '-' : ''}${String(horas).padStart(2, '0')}:${String(minutos).padStart(2, '0')}:${String(segundos).padStart(2, '0')}`;
    } catch (err) {
        console.error("Erro ao calcular diferença: ", err);
        return null;
    }
}

// Executando o código
main().catch((err) => console.error(err));
