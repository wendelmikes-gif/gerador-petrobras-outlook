
/* Gerador Petrobras - Outlook Add-in
   Teste de preenchimento e leitura do clipboard.
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
            if (resultado.status === Office.AsyncResultStatus.Succeeded) {
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

async function preencherEmailTeste() {
    const item = Office.context.mailbox.item;

    if (!item) {
        atualizarStatus("Não foi possível acessar a mensagem atual.");
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
            item.subject.setAsync(assunto, callback)
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

async function testarLeituraClipboard() {
    const marcadorEsperado = "GERADOR_PETROBRAS_TESTE_001";

    try {
        atualizarStatus("Verificando acesso ao clipboard...");

        if (!navigator.clipboard ||
            typeof navigator.clipboard.readText !== "function") {
            throw new Error(
                "A API de leitura do clipboard não está disponível."
            );
        }

        const conteudo = await navigator.clipboard.readText();

        if (conteudo.trim() === marcadorEsperado) {
            atualizarStatus(
                "SUCESSO! O suplemento conseguiu ler o clipboard."
            );
        } else {
            atualizarStatus(
                "Leitura realizada, mas o conteúdo não corresponde " +
                "ao marcador de teste. Copie o texto indicado e tente novamente."
            );
        }

    } catch (erro) {
        atualizarStatus(
            "Não foi possível ler o clipboard: " + erro.message
        );
    }
}

Office.onReady(() => {
    const botaoPreencher = document.getElementById("btnPreencher");
    const botaoClipboard = document.getElementById("btnTestarClipboard");

    if (botaoPreencher) {
        botaoPreencher.disabled = false;
        botaoPreencher.addEventListener("click", preencherEmailTeste);
    }

    if (botaoClipboard) {
        botaoClipboard.disabled = false;
        botaoClipboard.addEventListener("click", testarLeituraClipboard);
    }

    atualizarStatus("Complemento conectado ao Outlook.");
});
