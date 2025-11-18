/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Initialize the page
function initializePage() {
  // Initialize submit button handler
  const submitButton = document.getElementById("submitConnection");
  if (submitButton) {
    submitButton.onclick = saveToken;
  }
  
  // Load saved token if exists
  loadSavedToken();
  
  // Check health status
  checkHealthStatus();
}

// Check health status of external API
async function checkHealthStatus() {
  const healthBox = document.getElementById("healthStatus");
  const healthText = document.getElementById("healthText");
  
  if (!healthBox || !healthText) return;
  
  // Set checking state
  healthBox.className = "health-status checking";
  healthText.textContent = "Checking...";
  
  try {
    const response = await fetch("https://flow.ciandt.com/ai-orchestration-api/v1/health", {
      method: "GET",
      mode: "cors",
      headers: {
        "Accept": "application/json"
      }
    });
    
    if (response.ok) {
      // Success
      healthBox.className = "health-status success";
      healthText.textContent = `${response.status}`;
    } else {
      // Error response
      healthBox.className = "health-status error";
      healthText.textContent = `${response.status}`;
    }
  } catch (error) {
    // Network error or CORS issue
    healthBox.className = "health-status error";
    healthText.textContent = "CORS/Network Error";
    console.error("Health check failed:", error);
    console.log("This may be a CORS issue. The API server needs to allow requests from your origin.");
  }
}

// Check if Office is available, otherwise init directly
if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady((info) => {
    initializePage();
  });
} else {
  // Standalone mode (not in Office environment)
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializePage);
  } else {
    initializePage();
  }
}

// Save token as a cookie
function saveToken() {
  const tokenInput = document.getElementById("connectionInput");
  const token = tokenInput.value.trim();
  
  if (!token) {
    alert("Please enter a token");
    return;
  }
  
  try {
    // Set cookie with 365 days expiration
    const expirationDays = 365;
    const date = new Date();
    date.setTime(date.getTime() + (expirationDays * 24 * 60 * 60 * 1000));
    const expires = "expires=" + date.toUTCString();
    document.cookie = `flowToken=${token}; ${expires}; path=/; SameSite=Strict`;
    
    alert("Token saved successfully!");
    console.log("Token saved as cookie");
  } catch (error) {
    console.error("Error saving token:", error);
    alert("Error saving token. Please try again.");
  }
}

// Load saved token from cookie
function loadSavedToken() {
  try {
    const savedToken = getCookie("flowToken");
    const tokenInput = document.getElementById("connectionInput");
    
    if (savedToken && tokenInput) {
      tokenInput.value = savedToken;
      console.log("Token loaded from cookie");
    }
  } catch (error) {
    console.error("Error loading token:", error);
  }
}

// Helper function to get cookie by name
function getCookie(name) {
  const nameEQ = name + "=";
  const cookies = document.cookie.split(';');
  
  for (let i = 0; i < cookies.length; i++) {
    let cookie = cookies[i].trim();
    if (cookie.indexOf(nameEQ) === 0) {
      return cookie.substring(nameEQ.length, cookie.length);
    }
  }
  return null;
}

// Get the stored token (can be used by other functions)
export function getToken() {
  return getCookie("flowToken");
}

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
