/* 
   GERADOR PETROBRAS
   Suplemento de integração com Outlook Web.
   Mantém o teste de preenchimento e adiciona
   a verificação de leitura do clipboard.
*/

function atualizarStatus(mensagem) {
    const status = document.getElementById("status");

    if (status) {
        status.textContent = mensagem;
    }
}

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
        atualizarStatus("Solicitando leitura do clipboard...");

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

    const botaoPreencher =
        document.getElementById("btnPreencher");

    const botaoClipboard =
        document.getElementById("btnTestarClipboard");

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

        if (botaoPreencher) {
            botaoPreencher.disabled = false;
        }

        if (botaoClipboard) {
            botaoClipboard.disabled = false;
        }

        atualizarStatus(
            "Complemento conectado ao Outlook. Botões habilitados."
        );

    } else {
        atualizarStatus(
            "O suplemento carregou, mas o Outlook não foi identificado."
        );
    }
});
