import * as webllm from "https://esm.run/@mlc-ai/web-llm";

let engine = null;
let chatHistory = [];
let currentSystemPrompt = "";

// DOM elements
const statusDiv = document.getElementById('status');
const modelSelect = document.getElementById('modelSelect');
const systemPromptTextarea = document.getElementById('systemPrompt');
const writingInput = document.getElementById('writingInput');
const analyzeBtn = document.getElementById('analyzeBtn');
const updatePromptBtn = document.getElementById('updatePrompt');
const clearChatBtn = document.getElementById('clearChat');
const exportBtn = document.getElementById('exportBtn');
const chatContainer = document.getElementById('chatContainer');

// Initialize the model
async function initializeModel() {
    try {
        updateStatus('loading', 'Loading AI model... This may take a minute on first load.');
        
        const selectedModel = modelSelect.value;
        
        engine = await webllm.CreateMLCEngine(selectedModel, {
            initProgressCallback: (progress) => {
                updateStatus('loading', `Loading model: ${Math.round(progress.progress * 100)}%`);
            }
        });
        
        currentSystemPrompt = systemPromptTextarea.value;
        
        updateStatus('ready', '✓ AI model ready');
        analyzeBtn.disabled = false;
        
    } catch (error) {
        updateStatus('error', `Error loading model: ${error.message}`);
        console.error('Model initialization error:', error);
    }
}

// Update status display
function updateStatus(type, message) {
    statusDiv.className = `status ${type}`;
    
    if (type === 'loading') {
        statusDiv.innerHTML = `<span class="loading-spinner"></span>${message}`;
    } else {
        statusDiv.textContent = message;
    }
}

// Add message to chat
function addMessage(role, content) {
    const messageDiv = document.createElement('div');
    messageDiv.className = 'message';
    
    const roleSpan = document.createElement('div');
    roleSpan.className = `message-role ${role}`;
    roleSpan.textContent = role === 'user' ? 'Student' : 'Writing Coach';
    
    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content';
    contentDiv.textContent = content;
    
    messageDiv.appendChild(roleSpan);
    messageDiv.appendChild(contentDiv);
    
    // Remove empty state if present
    const emptyState = chatContainer.querySelector('.empty-state');
    if (emptyState) {
        emptyState.remove();
    }
    
    chatContainer.appendChild(messageDiv);
    chatContainer.scrollTop = chatContainer.scrollHeight;
    
    // Add to history
    chatHistory.push({ role, content });
}

// Get feedback on writing
async function analyzesWriting() {
    const writing = writingInput.value.trim();
    
    if (!writing) {
        alert('Please enter some writing to analyze.');
        return;
    }
    
    analyzeBtn.disabled = true;
    updateStatus('loading', 'Analyzing writing...');
    
    try {
        // Add user message
        addMessage('user', writing);
        
        // Build messages array with system prompt
        const messages = [
            { role: 'system', content: currentSystemPrompt },
            ...chatHistory.map(msg => ({ role: msg.role, content: msg.content }))
        ];
        
        // Get AI response
        const response = await engine.chat.completions.create({
            messages: messages,
        });
        
        const feedback = response.choices[0].message.content;
        
        // Add assistant response
        addMessage('assistant', feedback);
        
        // Clear input
        writingInput.value = '';
        
        updateStatus('ready', '✓ Feedback generated');
        
    } catch (error) {
        updateStatus('error', `Error: ${error.message}`);
        console.error('Analysis error:', error);
    } finally {
        analyzeBtn.disabled = false;
    }
}

// Update system prompt
async function updateSystemPrompt() {
    const newPrompt = systemPromptTextarea.value.trim();
    
    if (!newPrompt) {
        alert('Please enter coach instructions.');
        return;
    }
    
    currentSystemPrompt = newPrompt;
    
    // Reset the conversation with new system prompt
    chatHistory = [];
    chatContainer.innerHTML = '<div class="empty-state">Conversation reset with new coach instructions</div>';
    
    updateStatus('ready', '✓ Coach instructions updated');
}

// Clear chat history
function clearChat() {
    if (chatHistory.length === 0) return;
    
    if (confirm('Clear the conversation history?')) {
        chatHistory = [];
        chatContainer.innerHTML = '<div class="empty-state">Conversation cleared</div>';
        updateStatus('ready', '✓ Conversation cleared');
    }
}

// Export chat as DOCX
async function exportChat() {
    if (chatHistory.length === 0) {
        alert('No conversation to export.');
        return;
    }
    
    try {
        updateStatus('loading', 'Generating document...');
        
        // Create DOCX content using docx library
        const { Document, Packer, Paragraph, TextRun, HeadingLevel } = await import('https://esm.run/docx@8.5.0');
        
        const children = [
            new Paragraph({
                text: "Writing Coach Conversation",
                heading: HeadingLevel.HEADING_1,
            }),
            new Paragraph({
                text: `Exported: ${new Date().toLocaleString()}`,
                spacing: { after: 200 },
            }),
            new Paragraph({
                text: "Coach Instructions:",
                heading: HeadingLevel.HEADING_2,
            }),
            new Paragraph({
                text: currentSystemPrompt,
                spacing: { after: 400 },
            }),
            new Paragraph({
                text: "Conversation:",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 200 },
            }),
        ];
        
        // Add each message
        chatHistory.forEach((msg, index) => {
            const roleName = msg.role === 'user' ? 'Student' : 'Writing Coach';
            
            children.push(
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `${roleName}:`,
                            bold: true,
                        }),
                    ],
                    spacing: { before: 200 },
                })
            );
            
            children.push(
                new Paragraph({
                    text: msg.content,
                    spacing: { after: 200 },
                })
            );
        });
        
        const doc = new Document({
            sections: [{
                properties: {},
                children: children,
            }],
        });
        
        const blob = await Packer.toBlob(doc);
        
        // Download file
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `writing-coach-${new Date().toISOString().split('T')[0]}.docx`;
        a.click();
        URL.revokeObjectURL(url);
        
        updateStatus('ready', '✓ Document exported');
        
    } catch (error) {
        updateStatus('error', `Export error: ${error.message}`);
        console.error('Export error:', error);
    }
}

// Event listeners
analyzeBtn.addEventListener('click', analyzesWriting);
updatePromptBtn.addEventListener('click', updateSystemPrompt);
clearChatBtn.addEventListener('click', clearChat);
exportBtn.addEventListener('click', exportChat);

modelSelect.addEventListener('change', async () => {
    if (confirm('Changing models will reload the AI. Continue?')) {
        await initializeModel();
    }
});

// Allow Enter to submit (with Shift+Enter for new lines)
writingInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        if (!analyzeBtn.disabled) {
            analyzesWriting();
        }
    }
});

// Initialize on load
initializeModel();