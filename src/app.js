Office.onReady(function () {
    console.log("Office.js is ready");
    document.getElementById("send-button").addEventListener("click", sendMessage);
    document.getElementById("message-input").addEventListener("keypress", function (e) {
        if (e.key === "Enter") sendMessage();
    });
});

async function sendMessage() {
    const inputField = document.getElementById("message-input");
    const message = inputField.value.trim();
    if (!message) return;

    console.log("User message:", message);
    appendMessage(message, "user-message");
    inputField.value = "";

    let selectedText = "";
    let range = null; // Store range for potential override
    await Word.run(async (context) => {
        range = context.document.getSelection();
        range.load("text");
        await context.sync();
        selectedText = range.text;
        console.log("Selected text from Word:", selectedText);
    }).catch((error) => console.log("Error getting selection:", error));

    const prompt = selectedText 
        ? `Context from document: "${selectedText}". User message: ${message}. Generate only one crisp response until asked to elaborate more.` 
        : message;
    console.log("Prompt sent to Gemini:", prompt);

    streamGeminiResponse(prompt, selectedText, range);
}

async function streamGeminiResponse(prompt, selectedText, range) {
    try {
        console.log("Starting Gemini stream request");
        const apiKey = "test_key"; // Replace with your API key
        const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + apiKey;

        const response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{ parts: [{ text: prompt }] }],
                generationConfig: { temperature: 0.7, maxOutputTokens: 2048 }
            })
        });

        if (!response.ok) {
            console.log("API response status:", response.status);
            throw new Error(`API request failed with status ${response.status}`);
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let llmMessageDiv = null;
        let fullResponse = ""; // Store full response
        let buffer = "";

        while (true) {
            const { done, value } = await reader.read();
            if (done) {
                console.log("Stream completed");
                break;
            }

            const chunk = decoder.decode(value, { stream: true });
            buffer += chunk;
            console.log("Raw chunk:", chunk);

            let startIndex = 0;
            while (true) {
                try {
                    const jsonEnd = buffer.indexOf("}", startIndex) + 1;
                    if (jsonEnd <= 0) break;

                    const potentialJson = buffer.substring(startIndex, jsonEnd);
                    const data = JSON.parse(potentialJson);
                    const text = data.candidates?.[0]?.content?.parts?.[0]?.text || "";
                    
                    if (text) {
                        if (!llmMessageDiv) {
                            llmMessageDiv = appendMessage("", "llm-message", true);
                        }
                        llmMessageDiv.textContent += text;
                        fullResponse += text;
                        console.log("Streamed chunk:", text);
                        scrollToBottom();
                    }

                    startIndex = jsonEnd;
                } catch (e) {
                    break;
                }
            }

            buffer = buffer.substring(startIndex);
        }

        // Handle any remaining buffer
        if (buffer.trim()) {
            try {
                const data = JSON.parse(buffer);
                const text = data.candidates?.[0]?.content?.parts?.[0]?.text || "";
                if (text) {
                    if (!llmMessageDiv) {
                        llmMessageDiv = appendMessage("", "llm-message", true);
                    }
                    llmMessageDiv.textContent += text;
                    fullResponse += text;
                    console.log("Final streamed chunk:", text);
                    scrollToBottom();
                }
            } catch (e) {
                console.log("Error parsing final buffer:", e, "Buffer:", buffer);
            }
        }

        // Replace the selection part in streamGeminiResponse with this code:
        if (selectedText && fullResponse) {
            let position = { top: 0, left: 0 };
            
            try {
                await Word.run(async (context) => {
                    const range = context.document.getSelection();
                    range.load("rectangles");
                    await context.sync();
                    
                    if (range.rectangles && range.rectangles.length > 0) {
                        position = {
                            top: range.rectangles[0].top,
                            left: range.rectangles[0].left
                        };
                    } else {
                        // Fallback positioning relative to the Office taskpane
                        const taskpane = document.querySelector('.ms-Dialog');
                        if (taskpane) {
                            position = {
                                top: taskpane.offsetTop + 50,
                                left: taskpane.offsetLeft + 50
                            };
                        }
                    }
                });
            } catch (error) {
                console.log("Error getting selection position:", error);
            }

            showSuggestionHover(fullResponse, range, position);
        }

    } catch (error) {
        console.log("Gemini Streaming Error:", error.message);
        appendMessage(`Error: ${error.message}`, "llm-message");
    }
}

function appendMessage(text, className, isStreaming = false) {
    const chatContainer = document.getElementById("chat-container");
    if (!chatContainer) {
        console.error("Chat container not found!");
        return null;
    }
    
    const messageDiv = document.createElement("div");
    messageDiv.className = `message ${className}`;
    messageDiv.textContent = text;
    chatContainer.appendChild(messageDiv);
    if (!isStreaming) scrollToBottom();
    return messageDiv;
}

function scrollToBottom() {
    const chatContainer = document.getElementById("chat-container");
    if (chatContainer) {
        chatContainer.scrollTop = chatContainer.scrollHeight;
    }
}

function showConfirmDialog(response, range) {
    const dialog = document.getElementById("confirm-dialog");
    const preview = document.getElementById("response-preview");
    preview.textContent = response;
    dialog.style.display = "block";

    document.getElementById("confirm-yes").onclick = async () => {
        await Word.run(async (context) => {
            range.insertText(response, "Replace");
            await context.sync();
            console.log("Text replaced in Word:", response);
        });
        dialog.style.display = "none";
    };

    document.getElementById("confirm-no").onclick = () => {
        console.log("User declined to replace text");
        dialog.style.display = "none";
    };
}

function showSuggestionHover(suggestion, range, position) {
    const hoverElement = document.createElement('div');
    hoverElement.className = 'suggestion-hover';
    
    const preview = document.createElement('div');
    preview.textContent = suggestion.length > 200 
        ? suggestion.substring(0, 200) + '...' 
        : suggestion;
    hoverElement.appendChild(preview);

    const actionsDiv = document.createElement('div');
    actionsDiv.className = 'suggestion-actions';
    
    const acceptButton = document.createElement('button');
    acceptButton.className = 'suggestion-action-button';
    acceptButton.textContent = 'Accept (Tab)';
    acceptButton.onclick = async () => {
        try {
            await Word.run(async (context) => {
                // Get the current selection
                const currentRange = context.document.getSelection();
                // Insert the suggestion text at the selection
                currentRange.insertText(suggestion, "Replace");
                await context.sync();
                console.log("Text inserted successfully at selection");
            });
        } catch (error) {
            console.error("Error inserting text:", error);
        }
        hoverElement.remove();
    };

    const rejectButton = document.createElement('button');
    rejectButton.className = 'suggestion-action-button';
    rejectButton.textContent = 'Reject (Esc)';
    rejectButton.onclick = () => hoverElement.remove();

    actionsDiv.appendChild(acceptButton);
    actionsDiv.appendChild(rejectButton);
    hoverElement.appendChild(actionsDiv);

    hoverElement.style.position = 'absolute';
    hoverElement.style.top = `${position.top + 20}px`;
    hoverElement.style.left = `${position.left}px`;

    document.body.appendChild(hoverElement);

    document.addEventListener('keydown', function handleKeys(e) {
        if (e.key === 'Tab') {
            e.preventDefault();
            acceptButton.click();
            document.removeEventListener('keydown', handleKeys);
        } else if (e.key === 'Escape') {
            rejectButton.click();
            document.removeEventListener('keydown', handleKeys);
        }
    });
}