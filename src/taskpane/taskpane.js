// ====================================================================
// taskpane.js - CÓDIGO COMPLETO E CORRIGIDO
// ====================================================================

// --- 1. FUNÇÃO DE INICIALIZAÇÃO DO OFFICE ---
// Este é o ponto de entrada. Ele só executa o código APÓS a API do Office estar pronta.
Office.onReady(async (info) => {
    if (info.isSuccess && info.host === Office.HostType.Excel) {
        // O Office está pronto e o add-in está sendo executado no Excel.
        
        // Associa as funções assíncronas aos eventos de clique dos botões
        document.getElementById("connect-button").onclick = connectToFlow;
        document.getElementById("test-connection-button").onclick = testConnection;
        
        displayFeedback("Add-in pronto. Conecte-se ao CI&T Flow.");
    }
});

// --- 2. FUNÇÕES DO PAINEL DE TAREFAS (AGORA ASYNC) ---

/**
 * Função para lidar com o clique do botão "Conectar ao CI&T Flow".
 * Deve ser assíncrona porque pode envolver a leitura do Excel (Excel.run)
 * ou uma chamada de autenticação (fetch).
 */
// Variável global para armazenar a referência do diálogo
let authDialog = null;

async function connectToFlow() {
    displayFeedback("Aguardando credenciais...");

    // URL do arquivo que será exibido no modal
    // No ambiente de desenvolvimento (npm start), use a URL base do seu add-in
    const dialogUrl = `${window.location.protocol}//${window.location.host}/dialog.html`;

    try {
        // Opções para o tamanho do modal (em porcentagem)
        const dialogOptions = { height: 40, width: 30 }; 

        Office.context.ui.displayDialogAsync(dialogUrl, dialogOptions, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                displayFeedback(`Erro ao abrir modal: ${asyncResult.error.message}`);
                console.error(asyncResult.error);
                return;
            }

            // Abertura bem-sucedida, armazena a referência e configura os eventos
            authDialog = asyncResult.value;

            // 1. Manipulador para receber a mensagem do modal (as credenciais)
            authDialog.addEventHandler(Office.EventType.DialogMessageReceived, processCredentials);

            // 2. Manipulador para erros ou fechamento inesperado
            authDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
        });

    } catch (error) {
        displayFeedback(`Erro inesperado: ${error.message}`);
    }
}

/**
 * Funçao de callback que é executada quando o modal envia uma mensagem
 * (neste caso, as credenciais).
 */
function processCredentials(arg) {
    authDialog.close(); // Fecha o modal após receber a mensagem
    
    try {
        const data = JSON.parse(arg.message); // O dado é enviado como string JSON
        
        // **AQUI ESTÃO AS SUAS CREDENCIAIS!**
        const { username, password } = data;

        displayFeedback(`Credenciais recebidas para o usuário: ${username}. Iniciando autenticação na API...`);
        // Agora você pode usar 'username' e 'password' para fazer a chamada real de autenticação.

    } catch (e) {
        displayFeedback("Erro ao processar a mensagem do modal.");
    }
}

/**
 * Funçao de callback para lidar com o fechamento/erros do modal
 */
function dialogClosed(arg) {
    switch (arg.error) {
        case 12002: // O usuário clicou no X
            displayFeedback("Conexão cancelada pelo usuário.");
            break;
        case 12006: // Erro de autenticação (pode ser usado em fluxos OAuth)
            displayFeedback("Erro no fluxo de autenticação.");
            break;
        default:
            displayFeedback("O modal foi fechado inesperadamente.");
            break;
    }
    authDialog = null;
}

/**
 * Função para lidar com o clique do botão "Testar Conexão".
 * Assíncrona para permitir a chamada de rede (fetch).
 */
async function testConnection() {
    displayFeedback("Testando a conexão com a API...");

    try {
        // --- Chamada HTTP real usando fetch e await ---
        const response = await fetch("https://api.ci-tflow.com/v1/healthcheck", {
            method: 'GET',
            // Adicionar cabeçalhos de autorização aqui, se necessário.
            // headers: { 'Authorization': 'Bearer SEU_TOKEN_AQUI' } 
        });

        if (response.ok) {
            displayFeedback("Teste de conexão OK! O serviço CI&T Flow está ativo.");
        } else {
            // Se o status for 4xx ou 5xx
            displayFeedback(`Erro HTTP ${response.status}: Falha ao conectar. Verifique o token.`);
        }

    } catch (error) {
        // Captura erros de rede (DNS, timeout, etc.)
        displayFeedback("Erro de rede: Não foi possível alcançar o servidor CI&T Flow.");
        console.error("Erro em testConnection:", error);
    }
}

/**
 * Função auxiliar para exibir mensagens de feedback no painel (síncrona).
 */
function displayFeedback(message) {
    const feedbackElement = document.getElementById("feedback-message");
    if (feedbackElement) {
        feedbackElement.innerText = message;
    }
}