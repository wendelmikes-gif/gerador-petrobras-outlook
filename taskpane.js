/*
   GERADOR PETROBRAS
   Suplemento de integração com Outlook Web.
   Recuperação de pacote temporário do Gerador 2H3.
   Mantém os testes de preenchimento e clipboard.
*/

const WORKER_BASE_URL =
    "https://gerador-petrobras-bridge.wendelmikes.workers.dev";

const MARCADOR_PACOTE_REGEX =
    /\[\[GERADOR_PETROBRAS_PACKAGE:([a-f0-9]{64})\]\]/i;


/* ATUALIZAÇÃO DE STATUS */

function atualizarStatus(mensagem) {
    const status = document.getElementById("status");

    if (status) {
        status.textContent = mensagem;
    }
}


/* EXECUÇÃO ASSÍNCRONA DO OFFICE.JS */

function executarAsync(operacao) {
    return new Promise((resolve, reject) => {
        operacao((resultado) => {
            if (
                resultado.status ===
                Office.AsyncResultStatus.Succeeded
            ) {
                resolve(resultado);
            } else {
                reject(
                    new Error(
                        resultado.error?.message ||
                        "Não foi possível preencher o campo."
                    )
                );
            }
        });
    });
}


/* LEITURA DE CAMPO DA MENSAGEM */

function obterCampoAsync(operacao) {
    return new Promise((resolve, reject) => {
        operacao((resultado) => {
            if (
                resultado.status ===
                Office.AsyncResultStatus.Succeeded
            ) {
                resolve(resultado.value);
            } else {
                reject(
                    new Error(
                        resultado.error?.message ||
                        "Não foi possível ler o conteúdo da mensagem."
                    )
                );
            }
        });
    });
}


/* CONVERSÃO DE DESTINATÁRIOS */

function converterListaDestinatarios(valor) {
    if (!valor) {
        return [];
    }

    let lista = valor;

    if (typeof lista === "string") {
        const texto = lista.trim();

        if (!texto) {
            return [];
        }

        try {
            const json = JSON.parse(texto);

            if (Array.isArray(json)) {
                lista = json;
            } else {
                lista = texto.split(/[;,]+/);
            }
        } catch {
            lista = texto.split(/[;,]+/);
        }
    }

    if (!Array.isArray(lista)) {
        lista = [lista];
    }

    return lista
        .map((item) => {
            if (typeof item === "string") {
                const email = item.trim();

                if (!email) {
                    return null;
                }

                return {
                    emailAddress: email,
                    displayName: email
                };
            }

            if (item && typeof item === "object") {
                const email =
                    item.emailAddress ||
                    item.email ||
                    item.address;

                if (!email) {
                    return null;
                }

                return {
                    emailAddress: String(email).trim(),
                    displayName: String(
                        item.displayName ||
                        item.name ||
                        email
                    )
                };
            }

            return null;
        })
        .filter((item) => item !== null);
}


/* RECUPERAÇÃO DO PACOTE TEMPORÁRIO */

async function recuperarEmailGerador() {
    const item = Office.context.mailbox.item;

    if (!item) {
        atualizarStatus(
            "Não foi possível acessar a mensagem atual."
        );
        return;
    }

    const botao = document.getElementById("btnRecuperar");

    try {
        if (botao) {
            botao.disabled = true;
        }

        atualizarStatus(
            "Procurando o identificador temporário na mensagem..."
        );

        const corpoAtual = await obterCampoAsync(
            (callback) =>
                item.body.getAsync(
                    Office.CoercionType.Text,
                    callback
                )
        );

        const correspondencia =
            String(corpoAtual || "").match(MARCADOR_PACOTE_REGEX);

        if (!correspondencia) {
            throw new Error(
                "Identificador temporário não encontrado. " +
                "Prepare o e-mail no Gerador antes de abrir este suplemento."
            );
        }

        const token = correspondencia[1];

        atualizarStatus(
            "Recuperando o pacote temporário..."
        );

        const resposta = await fetch(
            `${WORKER_BASE_URL}/api/email-packages/consume/${token}`,
            {
                method: "GET",
                cache: "no-store",
                headers: {
                    "Accept": "application/json"
                }
            }
        );

        const resultado = await resposta.json();

        if (
            !resposta.ok ||
            resultado.success !== true ||
            !resultado.data
        ) {
            if (resposta.status === 410) {
                throw new Error(
                    "Este pacote expirou ou já foi utilizado. " +
                    "Prepare novamente o e-mail no Gerador."
                );
            }

            throw new Error(
                resultado.error ||
                resultado.message ||
                `Falha ao recuperar o pacote (HTTP ${resposta.status}).`
            );
        }

        const pacote = resultado.data;

        if (
            typeof pacote.assunto !== "string" ||
            typeof pacote.html !== "string" ||
            !pacote.html.trim()
        ) {
            throw new Error(
                "O pacote recuperado está incompleto."
            );
        }

        const destinatariosPara =
            converterListaDestinatarios(pacote.para);

        const destinatariosCc =
            converterListaDestinatarios(pacote.cc);

        if (destinatariosPara.length === 0) {
            throw new Error(
                "O pacote não contém destinatário em Para."
            );
        }

        atualizarStatus(
            "Preenchendo destinatários..."
        );

        await executarAsync((callback) =>
            item.to.setAsync(
                destinatariosPara,
                callback
            )
        );

        await executarAsync((callback) =>
            item.cc.setAsync(
                destinatariosCc,
                callback
            )
        );

        atualizarStatus(
            "Preenchendo assunto e template oficial..."
        );

        await executarAsync((callback) =>
            item.subject.setAsync(
                pacote.assunto,
                callback
            )
        );

        await executarAsync((callback) =>
            item.body.setAsync(
                pacote.html,
                {
                    coercionType: Office.CoercionType.Html
                },
                callback
            )
        );

        atualizarStatus(
            "E-mail recuperado e preenchido. " +
            "Confira destinatários, assunto e template antes de enviar."
        );

    } catch (erro) {
        atualizarStatus(
            "Não foi possível recuperar o e-mail: " +
            erro.message
        );

    } finally {
        if (botao) {
            botao.disabled = false;
        }
    }
}


/* TESTE ORIGINAL DE PREENCHIMENTO */

async function preencherEmailTeste() {
    const item = Office.context.mailbox.item;

    if (!item) {
        atualizarStatus(
            "Não foi possível acessar a mensagem atual."
        );
        return;
    }

    const assunto = "TESTE - GERADOR PETROBRAS";

    const corpoHTML = `
        <div style="font-family:Arial,sans-serif;color:#222;">
            <table style="width:100%;border-collapse:collapse;">
                <tr>
                    <td colspan="2"
                        style="background:#4D611A;color:white;
                        padding:12px;text-align:center;
                        font-weight:bold;">
                        OCORRÊNCIA - TESTE
                    </td>
                </tr>

                <tr>
                    <td style="width:25%;padding:10px;
                    border:1px solid #777;font-weight:bold;">
                        IDENTIFICAÇÃO
                    </td>

                    <td style="padding:10px;border:1px solid #777;">
                        TESTE 000001
                    </td>
                </tr>

                <tr>
                    <td style="padding:10px;border:1px solid #777;
                    font-weight:bold;">
                        OBSERVAÇÃO
                    </td>

                    <td style="padding:10px;border:1px solid #777;">
                        Este é um teste de integração do
                        Gerador Petrobras com o Outlook Web.
                    </td>
                </tr>
            </table>

            <p style="color:#c00000;font-style:italic;
            text-align:center;">
                MENSAGEM DE TESTE — NÃO ENVIAR
            </p>
        </div>
    `;

    try {
        atualizarStatus("Preenchendo destinatários...");

        await executarAsync((callback) =>
            item.to.setAsync(
                [{
                    displayName: "Destinatário de teste",
                    emailAddress: "destinatario@exemplo.com"
                }],
                callback
            )
        );

        await executarAsync((callback) =>
            item.cc.setAsync(
                [{
                    displayName: "Cópia de teste",
                    emailAddress: "copia@exemplo.com"
                }],
                callback
            )
        );

        atualizarStatus("Preenchendo assunto e corpo...");

        await executarAsync((callback) =>
            item.subject.setAsync(
                assunto,
                callback
            )
        );

        await executarAsync((callback) =>
            item.body.setAsync(
                corpoHTML,
                {
                    coercionType: Office.CoercionType.Html
                },
                callback
            )
        );

        atualizarStatus(
            "Teste concluído! Confira os campos no Outlook. Não envie."
        );

    } catch (erro) {
        atualizarStatus("Erro: " + erro.message);
    }
}


/* TESTE DE LEITURA DO CLIPBOARD */

async function testarLeituraClipboard() {
    const marcadorEsperado =
        "GERADOR_PETROBRAS_TESTE_001";

    try {
        atualizarStatus(
            "Solicitando leitura do clipboard..."
        );

        if (
            !navigator.clipboard ||
            typeof navigator.clipboard.readText !== "function"
        ) {
            throw new Error(
                "A leitura do clipboard não está disponível neste navegador."
            );
        }

        const conteudo =
            await navigator.clipboard.readText();

        if (conteudo.trim() === marcadorEsperado) {
            atualizarStatus(
                "SUCESSO! O suplemento conseguiu ler o clipboard."
            );
        } else {
            atualizarStatus(
                "Clipboard acessado, mas o conteúdo não corresponde " +
                "ao marcador esperado. Confira o texto copiado."
            );
        }

    } catch (erro) {
        atualizarStatus(
            "Falha na leitura do clipboard: " + erro.message
        );
    }
}


/* INICIALIZAÇÃO DO OUTLOOK */

Office.onReady((informacoes) => {

    const botaoRecuperar =
        document.getElementById("btnRecuperar");

    const botaoPreencher =
        document.getElementById("btnPreencher");

    const botaoClipboard =
        document.getElementById("btnTestarClipboard");

    if (botaoRecuperar) {
        botaoRecuperar.addEventListener(
            "click",
            recuperarEmailGerador
        );
    }

    if (botaoPreencher) {
        botaoPreencher.addEventListener(
            "click",
            preencherEmailTeste
        );
    }

    if (botaoClipboard) {
        botaoClipboard.addEventListener(
            "click",
            testarLeituraClipboard
        );
    }

    if (informacoes.host === Office.HostType.Outlook) {

        if (botaoRecuperar) {
            botaoRecuperar.disabled = false;
        }

        if (botaoPreencher) {
            botaoPreencher.disabled = false;
        }

        if (botaoClipboard) {
            botaoClipboard.disabled = false;
        }

        atualizarStatus(
            "Complemento conectado ao Outlook. " +
            "Botões habilitados."
        );

    } else {
        atualizarStatus(
            "O suplemento carregou, mas o Outlook não foi identificado."
        );
    }
});
