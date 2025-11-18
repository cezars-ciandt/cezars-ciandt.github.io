/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Initialize Office Add-in - everything runs under Excel context
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      console.log("Running in Excel context");

      document.getElementById("submitConnection").onclick = saveToken;
      document.getElementById("executeData").onclick = executeData;
    };

});
    
async function saveToken() {
  const tokenInput = document.getElementById("connectionInput");

  const token = tokenInput.value.trim();

  await Excel.run(async (context) => {
    const settings = context.workbook.settings;
    settings.add("flowToken", token);
    
  });
  
};

async function getToken() {

  await Excel.run(async (context) => {
    const settings = context.workbook.settings;
    const tokenRetr = settings.getItem("flowToken");
    tokenRetr.load("value");
  })
  await context.sync();
  
  return tokenRetr.value;
};

async function createFormData() {
    const boundary = '----FormBoundary' + Date.now();

    await Excel.run(async (context) => {
        const currentSheet = context.workbook.worksheets.getActiveWorksheet();
        const sheetName = currentSheet.load('name');
        const usedRange = currentSheet.getUsedRange();
        usedRange.load("values");

        await context.sync();

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(usedRange.values);
        XLSX.utils.book_append_sheet(wb, ws, currentSheet.name);

        // Converte a pasta de trabalho para um formato binário
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

        // Função para converter string em array buffer
        function s2ab(s) {
          const buf = new ArrayBuffer(s.length);
          const view = new Uint8Array(buf);
          for (let i = 0; i < s.length; i++) {
            view[i] = s.charCodeAt(i) & 0xFF;
              }
          return buf;
            };

        // Converte a string binária em um ArrayBuffer
        const fileData = s2ab(wbout);

        const authHeaders = { 'FlowToken': await getToken() };
        const BASE_URL = "https://flow.ciandt.com/advanced-flows";
        const FLOW_ID = "79f9184d-9823-4f73-b284-dd24bdb852dc";
        const protocol = new URL(BASE_URL).protocol;
        const httpModule = protocol === 'https:' ? require('https') : require('http');

        let data = '';
        data += `--${boundary}\r\n`;
        data += `Content-Disposition: form-data; name="file"; filename="${filename}"\r\n`;
        data += `Content-Type: application/octet-stream\r\n\r\n`;

        const payload = Buffer.concat([
            Buffer.from(data, 'utf8'),
            fileData,
            Buffer.from(`\r\n--${boundary}--\r\n`, 'utf8')
        ]);
        
        console.log({ payload, boundary });    
        return { payload, boundary };
      }
    )};

    // Helper function to make HTTP requests
function makeRequest(options, data) {
    return new Promise((resolve, reject) => {
        const req = httpModule.request(options, (res) => {
            let responseData = '';
            res.on('data', (chunk) => { responseData += chunk; });
            res.on('end', () => {
                if (res.statusCode >= 200 && res.statusCode < 300) {
                    try {
                        resolve(JSON.parse(responseData));
                    } catch (e) {
                        resolve(responseData);
                    }
                } else {
                    reject(new Error(`Request failed with status ${res.statusCode}: ${responseData}`));
                }
            });
        });
        req.on('error', reject);
        if (data) req.write(data);
        req.end();
    });
};

async function uploadAndExecuteFlow() {
    try {
        const authHeaders = { 'FlowToken': await getToken() };

        // Step 1: Upload file for File/VideoFile File-UOq2c
        const { payload: filePayload1, boundary: fileBoundary1 } = await createFormData();

        const fileUploadOptions1 = {
            hostname: 'flow.ciandt.com',
            port: 443,
            path: '/api/v2/files',
            method: 'POST',
            headers: {
                'Content-Type': `multipart/form-data; boundary=${fileBoundary1}`,
                'Content-Length': filePayload1.length,
                ...authHeaders
            }
        };

        const fileUploadResult1 = await makeRequest(fileUploadOptions1, filePayload1);
        const filePath1 = fileUploadResult1.path;
        console.log('File upload 1 successful! File path:', filePath1);

        // Step 2: Execute flow with all file paths
        const executePayload = JSON.stringify({
            "output_type": "chat",
            "input_type": "chat",
            "input_value": "hello world!",
            "tweaks": {
            "File-UOq2c": {
                      "path": [
                                "filePath1"
                      ]
            }
            }
        });

        const executeOptions = {
            hostname: 'flow.ciandt.com',
            port: 443,
            path: `/api/v1/run/lithia-ddl-gen-test`,
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(executePayload),
                ...authHeaders
            }
        };

        const result = await makeRequest(executeOptions, executePayload);
        console.log('Flow execution successful!');
        console.log(result);

    } catch (error) {
        console.error('Error:', error.message);
    }
};
        
    //     // Cria uma nova pasta de trabalho e adiciona os dados
    //     const wb = XLSX.utils.book_new();
    //     const ws = XLSX.utils.aoa_to_sheet(usedRange.values);
    //     XLSX.utils.book_append_sheet(wb, ws, activeSheet.name);

    //     // Converte a pasta de trabalho para um formato binário
    //     const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

    //     // Função para converter string em array buffer
    //     function s2ab(s) {
    //       const buf = new ArrayBuffer(s.length);
    //       const view = new Uint8Array(buf);
    //       for (let i = 0; i < s.length; i++) {
    //         view[i] = s.charCodeAt(i) & 0xFF;
    //           }
    //       return buf;
    //     };

    //     // Converte a string binária em um ArrayBuffer
    //     const binaryData = s2ab(wbout);
    //     const apiURL = "https://flow.ciandt.com/advanced-flows/" + "/api/v1/run/" + "79f9184d-9823-4f73-b284-dd24bdb852dc";

    //     let headers = 
    // });

    // return { payload, boundary };


async function executeData() {
  await uploadAndExecuteFlow();

};