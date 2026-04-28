/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("Office is ready in Outlook with reason: " + info.reason);
    
    // Hide sideload message and show app body
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    document.getElementById("getMessageBody").onclick = getMessageBody;
    document.getElementById("testRegex").onclick = testRegex;
    document.getElementById("regexInput").oninput = updateTestRegexButtonState;
    
    // Set initial button state
    updateTestRegexButtonState();
    
    // Add event handler for item changed (for pinnable task pane)
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      onItemChanged,
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to add item changed handler: " + asyncResult.error.message);
        } else {
          console.log("Item changed handler added successfully");
        }
      }
    );
  }
});

function updateTestRegexButtonState() {
  const regexInput = document.getElementById("regexInput");
  const messageBodyContainer = document.getElementById("messageBodyContainer");
  const testRegexButton = document.getElementById("testRegex");
  
  // Check if message body has content
  const hasMessageBody = messageBodyContainer && messageBodyContainer.textContent.trim() !== "";
  
  // Check if regex input has content and is valid
  let hasValidRegex = false;
  if (regexInput && regexInput.value.trim() !== "") {
    try {
      // Try to create the regex to validate it
      new RegExp(regexInput.value);
      hasValidRegex = true;
    } catch (error) {
      // Invalid regex
      hasValidRegex = false;
    }
  }
  
  // Enable button only if both conditions are met
  testRegexButton.disabled = !(hasMessageBody && hasValidRegex);
  
  console.log(`Test RegEx button state - hasMessageBody: ${hasMessageBody}, hasValidRegex: ${hasValidRegex}, disabled: ${testRegexButton.disabled}`);
}

function onItemChanged() {
  console.log("Item changed event triggered");
  
  // Only retrieve message body if it was previously retrieved
  const messageBodyContainer = document.getElementById("messageBodyContainer");
  if (messageBodyContainer && messageBodyContainer.textContent.trim() !== "") {
    console.log("Previous message had body content, retrieving new message body");
    getMessageBody();
  } else {
    console.log("Previous message had no body content, skipping auto-retrieval");
    // Clear message body and update button state
    messageBodyContainer.textContent = "";
    updateTestRegexButtonState();
  }
}

function getMessageBody() {
  // Get the selected format from radio buttons
  console.log("Getting message body...");
  const selectedFormat = document.querySelector('input[name="bodyFormat"]:checked').value;
  
  // Determine the coercion type based on selection
  const coercionType = selectedFormat === 'html' 
    ? Office.CoercionType.Html 
    : Office.CoercionType.Text;
  
  // Get the message body
  Office.context.mailbox.item.body.getAsync(
    coercionType,
    function (asyncResult) {
      const messageBodyContainer = document.getElementById("messageBodyContainer");
      console.log("Async result received for getMessageBody with status: " + asyncResult.status);
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        // Display the body content as text (not rendered HTML)
        messageBodyContainer.textContent = asyncResult.value;
        console.log("Message body retrieved successfully");
        
        // Update the Test RegEx button state
        updateTestRegexButtonState();
      } else {
        // Display error message
        messageBodyContainer.textContent = "Error retrieving message body: " + asyncResult.error.message;
        console.error("Error retrieving message body: " + asyncResult.error.message);
        
        // Update the Test RegEx button state
        updateTestRegexButtonState();
      }
    }
  );
}

function testRegex() {
  console.log("Testing RegEx...");
  
  const regexInput = document.getElementById("regexInput").value;
  const messageBodyContainer = document.getElementById("messageBodyContainer");
  const regexResultsSection = document.getElementById("regexResultsSection");
  const regexResultsContainer = document.getElementById("regexResultsContainer");
  
  // Validate inputs
  if (!regexInput || regexInput.trim() === "") {
    regexResultsContainer.textContent = "Error: Please enter a regular expression pattern.";
    regexResultsSection.style.display = "block";
    return;
  }
  
  if (!messageBodyContainer.textContent || messageBodyContainer.textContent.trim() === "") {
    regexResultsContainer.textContent = "Error: Please retrieve the message body first.";
    regexResultsSection.style.display = "block";
    return;
  }
  
  try {
    // Create the regular expression with global flag to find all matches
    const regex = new RegExp(regexInput, "g");
    const messageBody = messageBodyContainer.textContent;
    
    // Find all matches
    const matches = [];
    let match;
    while ((match = regex.exec(messageBody)) !== null) {
      matches.push(match[0]);
    }
    
    // Display results
    if (matches.length > 0) {
      let resultText = `Found ${matches.length} match${matches.length === 1 ? '' : 'es'}:\n\n`;
      matches.forEach((m, index) => {
        resultText += `${index + 1}. ${m}\n`;
      });
      regexResultsContainer.textContent = resultText;
      console.log(`RegEx test found ${matches.length} matches`);
    } else {
      regexResultsContainer.textContent = "No matches found.";
      console.log("RegEx test found no matches");
    }
    
    // Show the results section
    regexResultsSection.style.display = "block";
    
  } catch (error) {
    regexResultsContainer.textContent = "Error: Invalid regular expression - " + error.message;
    regexResultsSection.style.display = "block";
    console.error("RegEx test error: " + error.message);
  }
}