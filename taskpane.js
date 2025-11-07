Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("testConnection").onclick = testConnection;
        document.getElementById("createConnection").onclick = showTokenForm;
        document.getElementById("executeData").onclick = executeData;
        document.getElementById("saveToken").onclick = saveToken;
        document.getElementById("cancelToken").onclick = hideTokenForm;
    }
});

// Cookie utilities
function setCookie(name, value, days) {
    const expires = new Date();
    expires.setTime(expires.getTime() + (days * 24 * 60 * 60 * 1000));
    document.cookie = `${name}=${value};expires=${expires.toUTCString()};path=/`;
}

function getCookie(name) {
    const nameEQ = name + "=";
    const ca = document.cookie.split(';');
    for (let i = 0; i < ca.length; i++) {
        let c = ca[i];
        while (c.charAt(0) === ' ') c = c.substring(1, c.length);
        if (c.indexOf(nameEQ) === 0) return c.substring(nameEQ.length, c.length);
    }
    return null;
}

// UI utilities
function showStatus(message, isError = false) {
    const statusDiv = document.getElementById("statusMessage");
    const alertClass = isError ? 'alert-danger' : 'alert-success';
    const icon = isError ? 'bi-exclamation-triangle' : 'bi-check-circle';
    
    statusDiv.innerHTML = `
        <div class="alert ${alertClass} alert-dismissible fade show" role="alert">
            <i class="${icon} me-2"></i>
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        </div>
    `;
    
    // Auto-hide after 5 seconds
    setTimeout(() => {
        const alert = statusDiv.querySelector('.alert');
        if (alert) {
            const bsAlert = new bootstrap.Alert(alert);
            bsAlert.close();
        }
    }, 5000);
}

function setLoading(elementId, isLoading) {
    const element = document.getElementById(elementId);
    const originalText = element.dataset.originalText || element.innerHTML;
    
    if (isLoading) {
        element.dataset.originalText = originalText;
        element.innerHTML = `
            <span class="spinner-border spinner-border-sm me-2" role="status"></span>
            Carregando...
        `;
        element.disabled = true;
    } else {
        element.innerHTML = originalText;
        element.disabled = false;
    }
}

// Token form functions
function showTokenForm() {
    const tokenForm = document.getElementById("tokenForm");
    tokenForm.classList.remove('d-none');
    document.getElementById("tokenInput").focus();
    
    // Scroll to form
    tokenForm.scrollIntoView({ behavior: 'smooth' });
}

function hideTokenForm() {
    const tokenForm = document.getElementById("tokenForm");
    tokenForm.classList.add('d-none');
    document.getElementById("tokenInput").value = '';
}

function saveToken() {
    const token = document.getElementById("tokenInput").value.trim();
    
    if (!token) {
        showStatus("Por favor, insira um token válido.", true);
        return;
    }
    
    // Save token for 6 months (180 days)
    setCookie('flowToken', token, 180);
    hideTokenForm();
    showStatus("Token salvo com sucesso!");
}

// API functions
async function testConnection() {
    setLoading('testConnection', true);
    
    try {
        const token = getCookie('flowToken');
        
        if (!token) {
            showStatus("Não conectado. Verifique suas credenciais.", true);
            return;
        }
        
        // Test API call - replace with your actual API endpoint
        const response = await fetch('https://jsonplaceholder.typicode.com/posts/1', {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (response.ok) {
            showStatus("Conexão estabelecida com sucesso!");
        } else {
            showStatus("Não conectado. Verifique suas credenciais.", true);
        }
        
    } catch (error) {
        console.error('Connection test error:', error);
        showStatus("Erro ao testar conexão. Verifique sua internet.", true);
    } finally {
        setLoading('testConnection', false);
    }
}

async function executeData() {
    setLoading('executeData', true);
    
    try {
        const token = getCookie('flowToken');
        
        if (!token) {
            showStatus("Token não encontrado. Crie uma conexão primeiro.", true);
            return;
        }
        
        // Get data from active worksheet
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            const range = worksheet.getUsedRange();
            
            if (!range) {
                showStatus("Nenhum dado encontrado na planilha ativa.", true);
                return;
            }
            
            range.load("values");
            await context.sync();
            
            // Convert data to binary format (JSON string to base64)
            const jsonData = JSON.stringify(range.values);
            const binaryData = btoa(unescape(encodeURIComponent(jsonData)));
            
            // Send data to API - replace with your actual API endpoint
            const response = await fetch('https://jsonplaceholder.typicode.com/posts', {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    data: binaryData,
                    format: 'base64',
                    timestamp: new Date().toISOString()
                })
            });
            
            if (response.ok) {
                const result = await response.json();
                showStatus(`Dados enviados com sucesso! ID: ${result.id}`);
            } else {
                showStatus("Erro ao enviar dados. Verifique sua conexão.", true);
            }
        });
        
    } catch (error) {
        console.error('Execute data error:', error);
        showStatus("Erro ao processar dados da planilha.", true);
    } finally {
        setLoading('executeData', false);
    }
}

// Convert worksheet data to binary format
function convertToBinary(data) {
    try {
        // Convert array data to JSON string
        const jsonString = JSON.stringify(data);
        
        // Convert to base64 (binary representation)
        const binaryData = btoa(unescape(encodeURIComponent(jsonString)));
        
        return binaryData;
    } catch (error) {
        console.error('Binary conversion error:', error);
        throw new Error('Falha na conversão para formato binário');
    }
}
