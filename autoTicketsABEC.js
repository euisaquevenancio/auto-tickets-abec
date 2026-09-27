/*
    @euisaquevenancio - 17/07/2026
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

// Bibliotecas utilizadas
require("dotenv").config(); // Manipulação de variáveis de ambiente .env
const fs = require("fs"); // Manipulação de arquivos - File System
const puppeteer = require("puppeteer"); // Manipulação do navegador - Chrome ou Firefox
const ExcelJS = require("exceljs"); // Manipulação de arquivos Excel
const path = require("path"); // Manipulação de caminhos de arquivos
const { exec } = require("child_process"); // Execução de comandos do sistema operacional - necessário para abrir o arquivo Excel no final do processo

// Capturando os dados de login no Citsmart
const usuario = process.env.USUARIO;
const senha = process.env.SENHA;

let contadorTickets = 0;
let listaTickets = [];
let listaRegularizacoes = [];
const listaContasFinanceiras = fs.readFileSync("dados/contasFinanceiras.txt", "utf-8")
                                 .split("\n")
                                 .map((contaFinanceira) => contaFinanceira?.trim())
                                 .filter((contaFinanceira) => contaFinanceira.length > 0);
// Capturando os tickets que serão ignorados
const listaTicketsIgnorados = fs.readFileSync("dados/ticketsIgnorados.txt", "utf-8")
                                .split("\n")
                                .map((ticketIgnorado) => ticketIgnorado?.trim())
                                .filter((ticketIgnorado) => ticketIgnorado.length > 0);

async function main() {
    // Declarando o navegador
    const navegador = await puppeteer.launch({
        executablePath: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        headless: false // Visualizar ou não a automação rodando
    });

    const horarioInicio = new Date().toLocaleTimeString("pt-BR");
    
    const abaNavegador = await navegador.newPage();
    abaNavegador.setDefaultTimeout(30000);

    // Realiza o login no Citsmart
    const loginSucesso = await loginCitsmart(abaNavegador);

    if (loginSucesso) {
        console.log();
        await capturarTickets(abaNavegador, "NOTA FISCAL ELETRÔNICA");
        await capturarTickets(abaNavegador, "NOTA DE TERCEIROS");
        await organizarLista(listaTickets);

        console.log();
        for (let i = 0; i < listaTickets.length; i++) {
            contadorTickets++;
            await acessarTicket(abaNavegador, listaTickets[i])
        }
    }
    
    navegador.close();
    const horarioFim = new Date().toLocaleTimeString("pt-BR");
    console.log(`\n🤖 Fim da execução do script às ${horarioFim}.`);
    console.log(`🕓 Tempo de execução: ${calcularDiferencaHoras(horarioInicio, horarioFim)}.\n`);

    await salvarTickets();
}

// Realiza o login no Citsmart
async function loginCitsmart(abaNavegador) {
    await abaNavegador.goto(
        "https://servicos.maristabrasil.org/citsmart/webmvc/login#/ec?idExperienceCenter=3q",
        { waitUntil: "networkidle2" }
    );

    // Realiza login no Citsmart
    if (!(await abaNavegador.$("#user_login") !== null)) {
        console.log("📢 O campo username não foi encontrado!");
        return false;
    }

    if (!(await abaNavegador.$("#password") !== null)) {
        console.log("📢 O campo password não foi encontrado!");
        return false;
    }

    await abaNavegador.type("#user_login", usuario);
    await new Promise((r) => setTimeout(r, 2000));
    await abaNavegador.type("#password", senha);

    await abaNavegador.keyboard.press("Enter");
    await new Promise((r) => setTimeout(r, 2000));

    await abaNavegador.type("#password", senha);

    await abaNavegador.keyboard.press("Enter");

    await abaNavegador.waitForNavigation({ waitUntil: "networkidle2" });
    await new Promise((r) => setTimeout(r, 2000));

    // Verificando se o elemento de mensagem de erro existe, já que a página não retorna nenhum status quando o login da certo ou não
    if (await abaNavegador.$("div.notification-container div.notification span.message") !== null) {
        console.log("📢  " + await abaNavegador.$eval(
            "div.notification-container div.notification span.message",
            elemento => elemento.textContent
        ));
        return false;
    }

    // Login realizado com sucesso
    return true;
}

// Captura as informações básicas do ticket, sem acessar
async function capturarTickets(abaNavegador, pesquisa) {
    let contadorTicketsPesquisa = 0;
    let contadorPagina = 1;
    let temPaginaSeguinte = true;

    await abaNavegador.goto(
        "https://servicos.maristabrasil.org/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load#/",
        { waitUntil: "networkidle2" }
    );

    await abaNavegador.reload({ waitUntil: "networkidle2" });
    await new Promise((r) => setTimeout(r, 4000));

    if (!(await abaNavegador.$("#pesquisaSolicitacao") !== null)) {
        console.log("📢 A barra de pesquisa não foi encontrada!");
        return false;
    }

    await abaNavegador.focus("#pesquisaSolicitacao");
    await abaNavegador.click("#pesquisaSolicitacao", { clickCount: 3 });
    await abaNavegador.keyboard.press("Backspace");
    await abaNavegador.type("#pesquisaSolicitacao", pesquisa);
    await abaNavegador.keyboard.press("Enter");
    await abaNavegador.keyboard.press("Enter");
    await new Promise((r) => setTimeout(r, 4000));

    do {
        // Captura todos os tickets da página
        const novosTicketsDaPagina = await abaNavegador.$$("[name=list-item]");

        for (let i = 0; i < novosTicketsDaPagina.length; i++) {
            // Captura as informações da capa do ticket
            const informacoesDoTicket = await abaNavegador.evaluate((elemento, pesquisa) => {
                const elementoNumeroTicket = elemento.querySelector(".request-id");
                const elementoDataCriacaoTicket = elemento.querySelector(".dataCriacao");
                
                if (elementoNumeroTicket && elementoDataCriacaoTicket) {
                    return {
                        numero: elementoNumeroTicket.textContent?.trim(),
                        dataCriacao: elementoDataCriacaoTicket.textContent?.trim().split(" ")[0],
                        fila: pesquisa
                    };
                } else {
                    console.log("📢 As informações da capa do ticket não foram encontrada!");
                    return null;
                }
            }, novosTicketsDaPagina[i], pesquisa);

            if (informacoesDoTicket == null) continue;
            if (listaTicketsIgnorados.includes(informacoesDoTicket.numero)) continue;
            if (informacoesDoTicket.numero == 485540 || informacoesDoTicket.numero == 581721 || informacoesDoTicket.numero == 627350 || informacoesDoTicket.numero == 645281 || informacoesDoTicket.numero == 649397) continue;

            // Adiciona o ticket na lista
            listaTickets.push(informacoesDoTicket);
            contadorTicketsPesquisa++;
        }

        const buttonAvancarPaginaEstaDesabilitado = await abaNavegador.$eval(
            "#button-avancar-pesquisa",
            (elementoButtonAvancarPagina) => elementoButtonAvancarPagina.hasAttribute("disabled")
        );
        // Finaliza a busca de tickets se não houver mais tickets em páginas seguintes
        if (buttonAvancarPaginaEstaDesabilitado) break;

        // Avança para a página seguinte, caso haja mais tickets
        await abaNavegador.click("#button-avancar-pesquisa");

        try {
            await abaNavegador.waitForSelector(".request-id", {
                visible: true,
                timeout: 20000
            });

            await new Promise((r) => setTimeout(r, 5000));
            contadorPagina++;
        } catch (erro) {
            console.log("📢 Erro ao tentar avançar para a próxima página de tickets: ", erro.message);
            break;
        }
    } while (temPaginaSeguinte);

    const ticketOuTickets = (contadorTicketsPesquisa > 1) ? "tickets" : "ticket";
    const paginaOuPaginas = (contadorPagina > 1) ? "páginas" : "página";

    console.log(`✅ Resultado pesquisa "${pesquisa}": ${contadorTicketsPesquisa} ${ticketOuTickets} em ${contadorPagina} ${paginaOuPaginas}.`);
    return true;
}

// Organiza os itens da lista em ordem crescente
async function organizarLista(lista) {
    lista.sort(function (ticketA, ticketB) {
        if (ticketA.numero > ticketB.numero) {
            return -1;
        }
        if (ticketA.numero < ticketB.numero) {
            return 1;
        }
        return 0;
    });
}

// Acessa o ticket, para capturar as demais informações
async function acessarTicket(abaNavegador, ticket) {
    await abaNavegador.goto(
        "https://servicos.maristabrasil.org/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load#/",
        { waitUntil: "networkidle2" }
    );

    await abaNavegador.reload({ waitUntil: "networkidle2" });
    await new Promise((r) => setTimeout(r, 3000));

    if (!(await abaNavegador.$("#pesquisaSolicitacao") !== null)) {
        console.log("📢 A barra de pesquisa não foi encontrada!");
        return;
    }

    // Pesquisa o ticket
    await abaNavegador.focus("#pesquisaSolicitacao");
    await abaNavegador.click("#pesquisaSolicitacao", { clickCount: 3 });
    await abaNavegador.type("#pesquisaSolicitacao", ticket.numero);
    await abaNavegador.keyboard.press("Enter");
    await abaNavegador.keyboard.press("Enter");
    await new Promise((r) => setTimeout(r, 4000));

    // Acessando o ticket
    await abaNavegador.waitForSelector(".request-id", { visible: true, timeout: 20000 });
    await abaNavegador.click(".request-id", { clickCount: 2 });
    // Aguardando a tela do ticket abrir
    await new Promise((r) => setTimeout(r, 4000));
    // Recarrega a página para tentar evitar falhas
    await abaNavegador.reload({ waitUntil: "networkidle2" });
    await new Promise((r) => setTimeout(r, 6000));

    const erro = await abaNavegador.waitForSelector("#div-form-builder > div > span.error", { visible: true, timeout: 2000 }).catch(() => null);
    if (erro) {
        await abaNavegador.reload({ waitUntil: "domcontentloaded" });
        await new Promise((r) => setTimeout(r, 4000));
    }

    // Iniciando a captura de informações do ticket
    const mantenedora = await abaNavegador.evaluate(() => {
        const elementoSelectMantenedora = document.querySelector("#testelancarnotasmbPage div:nth-child(2) div.col-md-2 select, #testelancarnotasmb\\.mantenedora div select, #formulariosGerais\\.Mantenedora div select");

        if (elementoSelectMantenedora) {
            const mantenedoraSelecionada = elementoSelectMantenedora.selectedOptions[0];

            if (mantenedoraSelecionada) {
                return mantenedoraSelecionada.textContent?.trim();
            } else {
                console.log("📢 Mantenedora não declarada!");
                return "";
            }
        }
    });
    ticket.mantenedora = mantenedora;

    // Ignorando o ticket, caso não pertença a ABEC ou seja o ticket 485540 (um chamado antigo pendente que não foi fechado)
    if (mantenedora == "SOME" || ticket.numero == "485540") {
        console.log(`❌ #${contadorTickets} | ${ticket.numero} | Criado em: ${ticket.dataCriacao} | Mantenedora: ${ticket.mantenedora}`);
        return;
    }

    // Nome do solicitante (quem abriu o ticket)
    const nomeSolicitante = await abaNavegador.evaluate(() => {
        const elementoNomeSolicitante = document.querySelector(".requester-information-title.ng-binding");
        if (elementoNomeSolicitante) {
            return elementoNomeSolicitante.textContent?.trim();
        }
        return "";
    });
    ticket.nomeSolicitante = nomeSolicitante;

    // E-mail do solicitante
    const emailSolicitante = await abaNavegador.evaluate(() => {
        const elementoEmailSolicitante = document.querySelector(".requester-information-value");
        if (elementoEmailSolicitante) {
            return elementoEmailSolicitante.textContent?.trim();
        }
        return "";
    });
    ticket.emailSolicitante = emailSolicitante;

    let observacao = "";
    // E-mails de solicitantes do RH (prioridade)
    const emailsRH = ["jaqueline.macedo@maristabrasil.org", "larissa.reis@maristabrasil.org", "beatriz.ricardo@maristabrasil.org", "karine.zorek@maristabrasil.org", "jennifer.soares@maristabrasil.org", "paloma.engels@maristabrasil.org"];
    if (emailSolicitante != null && emailsRH.includes(emailSolicitante)) {
        observacao = "RH - PRIORIDADE";
    }

    // Descrição do ticket
    const descricao = await abaNavegador.evaluate(() => {
        const elementoDescricao = document.querySelector("#service-request-view > div > div > div > div.service-request-wrapper > div > div > div.service-request-content.clearfix.s12 > div:nth-child(2) > div.panel.panel-default.service-request-panel-details > div > fieldset > div:nth-child(2) > div > div > div");
        if (elementoDescricao) {
            return elementoDescricao.textContent
                                    .replace(/\s*\n+\s*/g, ' ')
                                    .replace(/\s+/g, ' ')
                                    ?.trim();
        }
        return "";
    });
    ticket.descricao = descricao;

    // Unidade do ticket
    const unidade = await abaNavegador.evaluate((ticket) => {
        if (ticket.fila == "NOTA DE TERCEIROS") {
            return "-";
        }

        const elementoUnidade = document.querySelector("#testelancarnotasmb\\.unidade > div > select");
        
        if (elementoUnidade) {
            const unidadeSelecionada = elementoUnidade.selectedOptions[0];
            if (unidadeSelecionada) {
                if (unidadeSelecionada.textContent?.trim().substring(0, 2) == "--") {
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
    }, ticket);
    ticket.unidade = unidade;

    // Número do pedido
    const numeroPedido = await abaNavegador.evaluate(() => {
        const elementoNumeroPedido = document.querySelector("#testelancarnotasmb\\.numeropedido input, #testelancarnotasmb\\.numeropedido2 input, #testelancarnotasmb\\.npedido input");
        if (elementoNumeroPedido) {
            return elementoNumeroPedido.value?.trim();
        }
        return "";
    });
    ticket.numeroPedido = numeroPedido;

    // Tipo da nota
    const tipoNota = await abaNavegador.evaluate((ticket) => {
        if (ticket.fila == "NOTA DE TERCEIROS") {
            return "-";
        }

        const elementoTipoNota = document.querySelector("#testelancarnotasmb\\.tiponata > div > select");
        if (elementoTipoNota) {
            const tipoNotaSelecionado = elementoTipoNota.selectedOptions[0];

            if (tipoNotaSelecionado) {
                return tipoNotaSelecionado.textContent?.trim().toUpperCase();
            }
            return "";
        }
        return "";
    }, ticket);
    ticket.tipoNota = tipoNota;

    // Valor da nota
    const valorNota = await abaNavegador.evaluate(() => {
        const elementoValorNota = document.querySelector("#notasDeTerceiros_MB\\.valor_nota input, #testelancarnotasmb\\.valornotafiscalinicial input");
        if (elementoValorNota) {
            return elementoValorNota.value?.trim();
        }
        return "";
    });
    ticket.valorNota = valorNota;

    // Tipo de lançamento (contrato, OC ou regularização)
    const tipoLancamento = await abaNavegador.evaluate((ticket) => {
        if (ticket.fila == "NOTA DE TERCEIROS") {
            return "REGULARIZAÇÃO";
        } else {
            // Regularização
            if (!ticket.numeroPedido || ticket.numeroPedido == "-" || ticket.numeroPedido.toUpperCase() == "X" || ticket.numeroPedido.toUpperCase() == "O" || ticket.numeroPedido == "0" || ticket.numeroPedido.toUpperCase() == "REGULARIZAÇÃO" || ticket.numeroPedido.toUpperCase() == "REGULARIZACAO" || ticket.numeroPedido.toUpperCase() == "CRIAR" || ticket.numeroPedido.toUpperCase() == "NÃO TEM" || ticket.numeroPedido == "000000") {
                return "REGULARIZAÇÃO";
            }
            // OC
            if (/^\d{8}$/.test(ticket.numeroPedido) || ticket.numeroPedido.length == 11 || ticket.numeroPedido.length == 10 || ticket.numeroPedido.toUpperCase().includes("OC")) {
                return "OC";
            }
            // Contrato
            if ((ticket.numeroPedido.toUpperCase() == "CONTRATO" || ticket.numeroPedido.toUpperCase().includes("CONTRATO") || (ticket.numeroPedido.length >= 13 && ticket.numeroPedido.length <= 18)) && /[A-Za-z]/.test(ticket.numeroPedido) && /\d/.test(ticket.numeroPedido)) {
                return "CONTRATO";
            } else if (document.querySelector("#testelancarnotasmb\\.centro_custo input") && document.querySelector("#testelancarnotasmb\\.contamb input")) {
                return "REGULARIZAÇÃO";
            }
            return "VERIFICAR";
        }
    }, ticket);
    ticket.tipoLancamento = tipoLancamento;

    // Forma de pagamento
    const formaPagamento = await abaNavegador.evaluate(() => {
        const elementoFormaPagamento = document.querySelector("#testelancarnotasmb\\.vencimentopagamento select, #notasDeTerceiros_MB\\.tipo_pagamento > div > select");
        if (elementoFormaPagamento) {
            const formaPagamentoSelecionada = elementoFormaPagamento.selectedOptions[0];

            if (formaPagamentoSelecionada) {
                if (formaPagamentoSelecionada.textContent?.trim() == "--- SELECIONE ---") {
                    return "VERIFICAR";
                }
                if (formaPagamentoSelecionada.textContent?.trim().toUpperCase() == "DEPÓSITO" || formaPagamentoSelecionada.textContent?.trim().toUpperCase() == "DEPOSITO") {
                    return "CRÉDITO";
                }
                return formaPagamentoSelecionada.textContent?.trim().toUpperCase();
            }
            return "";
        }
        return "";
    });
    ticket.formaPagamento = formaPagamento;

    // Data de vencimento
    const dataVencimento = await abaNavegador.evaluate(() => {
        const elementoDataVencimento = document.querySelector("#testelancarnotasmb\\.data input, input#notasDeTerceiros_MB\\.data_vencimento");
        if (elementoDataVencimento) {
            return elementoDataVencimento.value?.trim();
        }
        return "";
    });
    ticket.dataVencimento = dataVencimento;

    // Capturando informações específicas do lançamento de regularizações
    let cnpjFornecedor = "";
    let numeroNota = "";
    if (tipoLancamento == "REGULARIZAÇÃO") {
        if (ticket.fila == "NOTA DE TERCEIROS") {
            ticket.centroCustos = "35146 OU 35119";
            ticket.contaFinanceira = "2141 - PUBLICIDADE";

            // CNPJ do fornecedor
            cnpjFornecedor = await abaNavegador.evaluate(() => {
                const elementoCnpj = document.querySelector("#notasDeTerceiros_MB\\.informe_cnpj");

                if (elementoCnpj) {
                    if (elementoCnpj.value?.trim() == "undefined") return "";
                    return elementoCnpj.value?.trim();
                }
                return "";
            });

            // Número da nota fiscal
            numeroNota = await abaNavegador.evaluate(() => {
                const elementoNumeroNota = document.querySelector("#notasDeTerceiros_MB\\.numero_nota input");

                if (elementoNumeroNota) {
                    if (elementoNumeroNota.value?.trim() == "undefined") return "";
                    return elementoNumeroNota.value?.trim();
                }
                return "";
            });
        } else {
            // Centro de custos - CR
            const centroCustos = await abaNavegador.evaluate(() => {
                const elementoCentroCustos = document.querySelector("#testelancarnotasmb\\.centro_custo input");
                if (elementoCentroCustos) {
                    if (elementoCentroCustos.value?.trim() == "35113") {
                        // CR 35113 - Desiree Silva
                        observacao = "CR 35113 NÃO CONSEGUE APROVAR - PERGUNTAR SE PRECISA DE OUTRO CR PARA ABRIR A REGULARIZAÇÃO";
                        return "CR 35113 NÃO CONSEGUE APROVAR - PERGUNTAR SE PRECISA DE OUTRO CR PARA ABRIR A REGULARIZAÇÃO";
                    } else if (elementoCentroCustos.value?.trim() == "35114") {
                        // CR 35114 - Letícia Castilhos
                        observacao = "CR 35114 NÃO CONSEGUE APROVAR - PERGUNTAR SE PRECISA DE OUTRO CR PARA ABRIR A REGULARIZAÇÃO";
                        return "CR 35114 NÃO CONSEGUE APROVAR - PERGUNTAR SE PRECISA DE OUTRO CR PARA ABRIR A REGULARIZAÇÃO";
                    } else if (elementoCentroCustos.value?.trim() == "35113") {
                        // CR 35344 - Luana Alvarenga
                        observacao = "CR 35344 NÃO CONSEGUE APROVAR - PERGUNTAR SE PRECISA DE OUTRO CR PARA ABRIR A REGULARIZAÇÃO";
                        return "CR 35344 NÃO CONSEGUE APROVAR - PERGUNTAR SE PRECISA DE OUTRO CR PARA ABRIR A REGULARIZAÇÃO";
                    }

                    return elementoCentroCustos.value?.trim();
                }
                return "";
            });
            ticket.centroCustos = centroCustos;

            // Conta financeira
            const contaFinanceira = await abaNavegador.evaluate((ticket, listaContasFinanceiras) => {
                const elementoContaFinanceira = document.querySelector("#testelancarnotasmb\\.contamb input");
                if (elementoContaFinanceira) {
                    let conta = elementoContaFinanceira.value?.trim();

                    if (conta == "FORMAÇÃO E DESENVOLVIMENTO") {
                        return "2041 - CURSOS E TREINAMENTOS";
                    } else if (conta == "EVENTO INTERNO" || elementoContaFinanceira.value?.trim() ==  "EVENTOS INTERNOS" || elementoContaFinanceira.value?.trim() == "EVENTOS INTERNO") {
                        return "2045 - EVENTOS INTERNOS";
                    } else if (conta == "01-35111") {
                        return "2030 - ASSISTÊNCIA MÉDICA E ODONTOLÓGICAS";
                    }

                    conta = listaContasFinanceiras.find(contaFinanceira => {
                        const [numero, descricao] = contaFinanceira.split(" - ");
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
                }
                return "";
            }, ticket, listaContasFinanceiras);
            ticket.contaFinanceira = contaFinanceira;
        }
        ticket.observacao = observacao;

        ticket.numeroNota = numeroNota;
        ticket.cnpjFornecedor = cnpjFornecedor;
        
        // Reunindo os dados do ticket para criar a regularização
        let regularizacao = {
            nomeSolicitante: ticket.nomeSolicitante,
            emailSolicitante: ticket.emailSolicitante,
            ticket: ticket.numero,
            numeroNota: ticket.numeroNota,
            cnpjFornecedor: ticket.cnpjFornecedor,
            valorNota: ticket.valorNota,
            dataVencimento: ticket.dataVencimento,
            descricao: ticket.descricao,
            centroCustos: ticket.centroCustos,
            contaFinanceira: ticket.contaFinanceira,
            unidade: ticket.unidade,
            formaPagamento: ticket.formaPagamento
        };
        
        listaRegularizacoes.push(regularizacao);
    }
    
    console.log(`✅ #${contadorTickets} | ${ticket.numero} | Criado em: ${ticket.dataCriacao} | Vencimento: ${ticket.dataVencimento} | Valor: R$ ${ticket.valorNota} | Mantenedora: ${ticket.mantenedora} | UNIDADE: ${ticket.unidade}`);
}

async function salvarTickets() {
    const pastaDestino = path.join(__dirname, "dados");

    // Garante que a pasta exista
    if (!fs.existsSync(pastaDestino)) {
        fs.mkdirSync(pastaDestino, { recursive: true });
    }

    // Captura o arquivo Excel
    const arquivoExcel = path.join(pastaDestino, "tickets.xlsx");

    const workbook = new ExcelJS.Workbook();
    // Se o arquivo existe, lê
    if (fs.existsSync(arquivoExcel)) {
        await workbook.xlsx.readFile(arquivoExcel);
    }

    let planilha = workbook.getWorksheet("Tickets");
    // Se a planilha não estiver estruturada, adiciona o cabeçalho
    if (!planilha) {
        planilha = workbook.addWorksheet("Tickets");

        planilha.addRow([
            "TICKET",
            "RESPONSÁVEL",
            "CRIADO EM",
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
    planilha.columns = [
        { key: "TICKET", width: 8 },
        { key: "RESPONSÁVEL", width: 8 },
        { key: "CRIADO EM", width: 8 },
        { key: "VENCIMENTO", width: 8 },
        { key: "FORMA PAGAMENTO", width: 8 },
        { key: "VALOR", width: 8 },
        { key: "TIPO", width: 8 },
        { key: "OBSERVAÇÃO", width: 8 },
        { key: "Nº PEDIDO", width: 8 },
        { key: "DESCRIÇÃO", width: 8 },
        { key: "UNIDADE", width: 8 },
        { key: "SOLICITANTE", width: 8 },
        { key: "TIPO DE NOTA", width: 8 },
        { key: "FILA", width: 8 }
    ];

    // Adicionando os tickets na planilha
    if (contadorTickets > 0 && listaTickets.length > 0) {
        for (let i = 0; i < listaTickets.length; i++) {
            if (listaTickets[i].mantenedora == "SOME" || listaTickets[i].numero == "485540") {
                contadorTickets--;
                continue;
            }

            const dataCriacao = converterParaData(listaTickets[i].dataCriacao);
            const dataVencimento = converterParaData(listaTickets[i].dataVencimento);

            planilha.addRow([
                listaTickets[i].numero,
                "",
                dataCriacao,
                dataVencimento,
                listaTickets[i].formaPagamento,
                listaTickets[i].valorNota,
                listaTickets[i].tipoLancamento,
                listaTickets[i].observacao,
                listaTickets[i].numeroPedido,
                listaTickets[i].descricao,
                listaTickets[i].unidade,
                listaTickets[i].emailSolicitante,
                listaTickets[i].tipoNota,
                listaTickets[i].fila
            ]);
        }
    }

    if (planilha.rowCount > 1) {
        // Remove os tickets da tabela antiga
        if (planilha.model.tables) {
            planilha.model.tables = [];
        }

        // Formata a coluna VENCIMENTO como dd/mm/aaaa
        for (let i = 2; i <= planilha.rowCount; i++) {
            planilha.getCell(i, 3).numFmt = "dd/mm/yyyy";
            planilha.getCell(i, 4).numFmt = "dd/mm/yyyy";
        }

        const linhasValidas = planilha.getSheetValues().slice(2) // Remove o cabeçalho
                                                       .filter(linha => Array.isArray(linha)) // Remove undefined
                                                       .map(linha => linha.slice(1)); // Remove índice fantasma
        
        planilha.addTable({
            name: "TabelaTickets",
            ref: "A1",
            headerRow: true,
            style: {
                theme: "TableStyleLight1"
            },
            columns: [
                { name: "TICKET" },
                { name: "RESPONSÁVEL" },
                { name: "CRIADO EM" },
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
                { name: "FILA" },
            ],
            rows: linhasValidas
        });
    }

    if (contadorTickets > 0 && listaTickets.length > 0) {
        // Salva o arquivo excel
        await workbook.xlsx.writeFile(arquivoExcel);

        // Arquivo que contém as regularizações
        const arquivoRegularizacoes = path.join(pastaDestino, "regularizacoes.txt");

        let conteudoExistenteRegularizacoes = "";
        // Se o arquivo existir, lê o conteúdo
        if (fs.existsSync(arquivoRegularizacoes)) {
            conteudoExistenteRegularizacoes = fs.readFileSync(arquivoRegularizacoes);
        }

        let novoConteudoRegularizacoes = "";
        for (const regularizacaoAtual of listaRegularizacoes) {
            // Verifica se o ticket já existe no arquivo
            if (!conteudoExistenteRegularizacoes.includes(`TICKET ${regularizacaoAtual.ticket}`)) {
                novoConteudoRegularizacoes += `PAGAMENTO SOLICITADO POR: ${regularizacaoAtual.nomeSolicitante} - ${regularizacaoAtual.emailSolicitante}\n` +
                                              `TICKET: ${regularizacaoAtual.ticket}\n` +
                                              `NÚMERO DA NOTA FISCAL: ${regularizacaoAtual.numeroNota}\n` +
                                              `VALOR DA NOTA FISCAL: ${regularizacaoAtual.valorNota}\n` +
                                              `DATA DE VENCIMENTO: ${regularizacaoAtual.dataVencimento}\n` +
                                              `DESCRIÇÃO: ${regularizacaoAtual.descricao}\n` +
                                              `CENTRO DE CUSTOS (CR): ${regularizacaoAtual.centroCustos}\n` +
                                              `CONTA FINANCEIRA: ${regularizacaoAtual.contaFinanceira}\n` +
                                              `FORNECEDOR: ${regularizacaoAtual.cnpjFornecedor}\n` +
                                              `UNIDADE: ${regularizacaoAtual.unidade}\n` +
                                              `FORMA DE PAGAMENTO: ${regularizacaoAtual.formaPagamento}\n\n`;
            }
        }

        // Só escreve no TXT se tiver conteúdo novo
        if (novoConteudoRegularizacoes != "") {
            fs.appendFileSync(arquivoRegularizacoes, novoConteudoRegularizacoes, "utf-8");
        }

        // Adicionando os novos tickets que devem ser ignorados (já estão na planilha)
        const arquivoTicketsIgnorados = path.join(pastaDestino, "ticketsIgnorados.txt");

        let conteudoExistenteTicketsIgnorados = "";
        // Se o arquivo existir, lê o conteúdo
        if (fs.existsSync(arquivoTicketsIgnorados)) {
            conteudoExistenteTicketsIgnorados = fs.readFileSync(arquivoTicketsIgnorados, "utf-8");
        }

        let novoConteudoTicketsIgnorados = "";
        for (const ticketAtual of listaTickets) {
            novoConteudoTicketsIgnorados += `\n${ticketAtual.numero}`;
        }

        // Só escreve no TXT se tiver conteúdo novo
        if (novoConteudoTicketsIgnorados != "") {
            fs.appendFileSync(arquivoTicketsIgnorados, novoConteudoTicketsIgnorados, "utf-8");
        }

        exec(`start "" "${arquivoExcel}"`);
    }
}

// Função para calcular a diferença entre dois horários no formato HH:mm:ss
function calcularDiferencaHoras(horarioInicio, horarioFim) {
    try {
        // Quebra as strings em partes
        const [h1, m1, s1] = horarioInicio.split(":").map(Number);
        const [h2, m2, s2] = horarioFim.split(":").map(Number);

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
        console.error("Erro ao calcular diferença de horas: ", err);
        return null;
    }
}

function converterParaData(data) {
    if (!data) return null;

    // Se já for um objeto Date
    if (data instanceof Date) {
        return data;
    }

    // Se vier como string no formato dd/mm/aaaa
    if (typeof data === "string") {
        const partes = data.trim().split("/");

        if (partes.length === 3) {
            const dia = Number(partes[0]);
            const mes = Number(partes[1]);
            const ano = Number(partes[2]);

            return new Date(ano, mes - 1, dia);
        }
    }

    return null;
}

// Executando o código
main().catch((err) => {
    console.error("Erro na execução do script: ", err);
});
