<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Curriculum Diagnostic & Fixer</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.4.21/mammoth.browser.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
body { font-family: 'Inter', sans-serif; }
.spinner {
    border: 3px solid rgba(255,255,255,0.3);
    border-radius: 50%;
    border-top: 3px solid #3b82f6;
    width: 24px;
    height: 24px;
    -webkit-animation: spin 1s linear infinite;
    animation: spin 1s linear infinite;
}
@-webkit-keyframes spin {
    0% { -webkit-transform: rotate(0deg); }
    100% { -webkit-transform: rotate(360deg); }
}
@keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
}
    </style>
</head>
<body class="bg-slate-50 min-h-screen text-slate-800 p-4 md:p-8">

    <div class="max-w-6xl mx-auto">
        <header class="mb-8 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
            <div>
                <h1 class="text-3xl font-bold text-slate-900">Curriculum Diagnostic & Fixer</h1>
                <p class="text-slate-500 mt-2">Evaluate and automatically correct teacher documents based on your official guidelines.</p>
            </div>
            <div class="bg-white p-3 rounded-lg shadow-sm border border-slate-200 w-full md:w-auto">
                <label for="apiKeyInput" class="block text-xs font-semibold text-slate-600 mb-1">Gemini API Key</label>
                <div class="flex gap-2">
                    <input type="password" id="apiKeyInput" placeholder="Enter API Key" class="text-sm border border-slate-300 rounded-md px-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-blue-500 w-full md:w-64">
                    <button onclick="saveApiKey()" class="bg-slate-800 hover:bg-slate-700 text-white text-xs font-medium py-1.5 px-3 rounded-md transition-colors">Save</button>
                </div>
                <p id="apiSaveMessage" class="text-xs text-green-600 mt-1 hidden">Key saved to browser!</p>
            </div>
        </header>

        <div class="grid grid-cols-1 lg:grid-cols-2 gap-8">
            <!-- Left Column: Inputs -->
            <div class="flex flex-col gap-6">
                
                <!-- Framework Input -->
                <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-200 flex-grow flex flex-col">
                    <label class="block text-sm font-semibold text-slate-700 mb-2">1. Diagnostic Framework / Guidelines (Upload)</label>
                    <p class="text-xs text-slate-500 mb-3">Upload your official guidelines rulebook. Supported formats: <b>.docx, .pdf, .txt, .md</b></p>
                    
                    <div id="frameworkDropZone" class="border-2 border-dashed border-slate-300 rounded-xl p-8 flex-grow flex flex-col items-center justify-center text-center cursor-pointer hover:bg-slate-50 transition-all" onclick="document.getElementById('frameworkFileInput').click()" ondragover="handleDragOver(event, 'frameworkDropZone')" ondragleave="handleDragLeave(event, 'frameworkDropZone')" ondrop="handleDrop(event, 'framework')">
                        <svg xmlns="http://www.w3.org/2000/svg" width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="text-blue-500 mb-3"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/></svg>
                        <p class="text-sm font-medium text-slate-700">Click to browse or drag & drop</p>
                        <p id="frameworkFileNameDisplay" class="text-xs text-slate-500 mt-2 bg-slate-100 py-1 px-3 rounded-full inline-block">No file selected</p>
                        <input type="file" id="frameworkFileInput" class="hidden" accept=".txt,.md,.docx,.pdf" onchange="handleFileUpload(event, 'framework')">
                    </div>
                </div>

                <!-- Teacher's Document Input -->
                <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-200 flex-grow flex flex-col">
                    <label class="block text-sm font-semibold text-slate-700 mb-2">2. Teacher's Document (Upload)</label>
                    <p class="text-xs text-slate-500 mb-3">Upload the lesson plan or reading text. Supported formats: <b>.docx, .pdf, .txt, .md</b></p>
                    
                    <div id="dropZone" class="border-2 border-dashed border-slate-300 rounded-xl p-8 flex-grow flex flex-col items-center justify-center text-center cursor-pointer hover:bg-slate-50 transition-all" onclick="document.getElementById('fileInput').click()" ondragover="handleDragOver(event, 'dropZone')" ondragleave="handleDragLeave(event, 'dropZone')" ondrop="handleDrop(event, 'document')">
                        <svg xmlns="http://www.w3.org/2000/svg" width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="text-blue-500 mb-3"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/></svg>
                        <p class="text-sm font-medium text-slate-700">Click to browse or drag & drop</p>
                        <p id="fileNameDisplay" class="text-xs text-slate-500 mt-2 bg-slate-100 py-1 px-3 rounded-full inline-block">No file selected</p>
                        <input type="file" id="fileInput" class="hidden" accept=".txt,.md,.docx,.pdf" onchange="handleFileUpload(event, 'document')">
                    </div>
                </div>

                <button id="analyzeBtn" onclick="analyzeDocument()" class="bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3 px-6 rounded-xl shadow-sm transition-colors flex justify-center items-center gap-3">
                    <span>Evaluate & Fix Document</span>
                    <div id="loadingSpinner" class="spinner hidden"></div>
                </button>
            </div>

            <!-- Right Column: Output -->
            <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-200 flex flex-col h-full min-h-[600px]">
                <h2 class="text-lg font-semibold text-slate-800 mb-4 flex items-center gap-2">
                    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="text-blue-600"><path d="M12 2v20"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>
                    AI Evaluation & Result
                </h2>
                
                <div id="resultsArea" class="flex-grow bg-slate-50 border border-slate-200 rounded-lg p-4 overflow-y-auto text-sm text-slate-600 whitespace-pre-wrap font-mono">
                    Results will appear here...
                </div>

                <button id="copyBtn" onclick="copyResults()" class="mt-4 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 px-4 rounded-lg border border-slate-300 transition-colors hidden">
                    Copy Revised Version
                </button>

                <button id="downloadBtn" onclick="downloadFixedDocument()" class="mt-2 bg-green-600 hover:bg-green-700 text-white font-semibold py-2 px-4 rounded-lg transition-colors hidden flex items-center justify-center gap-2">
                    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                    Download Fixed Document (.docx)
                </button>
            </div>
        </div>
    </div>

    <script>
let extractedDocumentText = "";
let extractedFrameworkText = "";
let extractedDocumentBase64 = null;
let extractedDocumentMimeType = null;
let extractedFrameworkBase64 = null;
let extractedFrameworkMimeType = null;
let originalTeacherFileName = "Document.docx";

// Initialize API Key from LocalStorage
document.addEventListener('DOMContentLoaded', () => {
    const savedKey = localStorage.getItem('geminiApiKey');
    if (savedKey) {
        document.getElementById('apiKeyInput').value = savedKey;
    }
});

function saveApiKey() {
    const key = document.getElementById('apiKeyInput').value;
    localStorage.setItem('geminiApiKey', key);
    const msg = document.getElementById('apiSaveMessage');
    msg.classList.remove('hidden');
    setTimeout(() => {
        msg.classList.add('hidden');
    }, 3000);
}

// Drag and Drop Event Handlers
function handleDragOver(event, zoneId) {
    event.preventDefault();
    document.getElementById(zoneId).classList.add('bg-blue-50', 'border-blue-400');
}

function handleDragLeave(event, zoneId) {
    event.preventDefault();
    document.getElementById(zoneId).classList.remove('bg-blue-50', 'border-blue-400');
}

function handleDrop(event, type) {
    event.preventDefault();
    const zoneId = type === 'framework' ? 'frameworkDropZone' : 'dropZone';
    const inputId = type === 'framework' ? 'frameworkFileInput' : 'fileInput';
    
    document.getElementById(zoneId).classList.remove('bg-blue-50', 'border-blue-400');
    
    if (event.dataTransfer.files && event.dataTransfer.files.length > 0) {
        document.getElementById(inputId).files = event.dataTransfer.files;
        handleFileUpload({ target: { files: event.dataTransfer.files } }, type);
    }
}

// File Parsing Logic
async function handleFileUpload(event, type = 'document') {
    const file = event.target.files[0];
    if (!file) return;

    const nameDisplayId = type === 'framework' ? 'frameworkFileNameDisplay' : 'fileNameDisplay';
    const nameDisplay = document.getElementById(nameDisplayId);
    nameDisplay.textContent = "Extracting text...";
    nameDisplay.className = "text-xs text-blue-600 mt-2 bg-blue-50 py-1 px-3 rounded-full inline-block font-medium";
    
    if (type === 'framework') {
        extractedFrameworkText = "";
        extractedFrameworkBase64 = null;
        extractedFrameworkMimeType = null;
    } else {
        extractedDocumentText = "";
        extractedDocumentBase64 = null;
        extractedDocumentMimeType = null;
    }

    try {
        let text = "";
        let base64Data = null;
        let mimeType = null;
        const fileExt = file.name.toLowerCase();
        
        if (fileExt.endsWith('.docx')) {
            try {
                const arrayBuffer = await file.arrayBuffer();
                const result = await mammoth.extractRawText({arrayBuffer: arrayBuffer});
                text = result.value || "";
                if (!text.trim()) throw new Error("Document is empty or could not be parsed.");
            } catch (err) {
                console.error("Docx parsing error:", err);
                throw new Error("Failed to parse the .docx file. It may be corrupted or encrypted.");
            }
        } else if (fileExt.endsWith('.pdf')) {
            try {
                // Read as base64 for native Gemini multimodal support
                base64Data = await new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve(reader.result.split(',')[1]);
                    reader.onerror = reject;
                    reader.readAsDataURL(file);
                });
                mimeType = "application/pdf";
            } catch (err) {
                console.error("PDF loading error:", err);
                throw new Error("Failed to load the PDF file for analysis.");
            }
        } else if (fileExt.endsWith('.txt') || fileExt.endsWith('.md')) {
            try {
                text = await file.text();
            } catch (err) {
                console.error("Text file parsing error:", err);
                throw new Error("Failed to read the text file.");
            }
        } else {
            showMessageBox("Unsupported File", "Please upload a .docx, .pdf, .txt, or .md file.");
            nameDisplay.textContent = "Invalid file type";
            nameDisplay.className = "text-xs text-red-600 mt-2 bg-red-50 py-1 px-3 rounded-full inline-block";
            return;
        }
        
        if (type === 'framework') {
            extractedFrameworkText = text;
            extractedFrameworkBase64 = base64Data;
            extractedFrameworkMimeType = mimeType;
        } else {
            extractedDocumentText = text;
            extractedDocumentBase64 = base64Data;
            extractedDocumentMimeType = mimeType;
            originalTeacherFileName = file.name;
        }
        
        nameDisplay.textContent = `✓ ${file.name}`;
        nameDisplay.className = "text-xs text-green-700 mt-2 bg-green-100 py-1 px-3 rounded-full inline-block font-semibold border border-green-200";
    } catch (error) {
        console.error("File reading error:", error);
        showMessageBox("Error", error.message || "Failed to read the document. Ensure it's a valid text, pdf, or docx file.");
        nameDisplay.textContent = "Error reading file";
        nameDisplay.className = "text-xs text-red-600 mt-2 bg-red-50 py-1 px-3 rounded-full inline-block";
    }
}

async function analyzeDocument() {
    const frameworkText = extractedFrameworkText;
    const documentText = extractedDocumentText;
    const analyzeBtn = document.getElementById('analyzeBtn');
    const loadingSpinner = document.getElementById('loadingSpinner');
    const resultsArea = document.getElementById('resultsArea');
    const copyBtn = document.getElementById('copyBtn');
    const downloadBtn = document.getElementById('downloadBtn');

    // Basic validation
    if (!frameworkText.trim() && !extractedFrameworkBase64) {
        showMessageBox("Error", "Please upload the Diagnostic Framework document first.");
        return;
    }
    if (!documentText.trim() && !extractedDocumentBase64) {
        showMessageBox("Error", "Please upload a valid Teacher Document first.");
        return;
    }

    // Update UI for loading state
    analyzeBtn.disabled = true;
    analyzeBtn.classList.add('opacity-75', 'cursor-not-allowed');
    loadingSpinner.classList.remove('hidden');
    resultsArea.innerHTML = "Analyzing document for accuracy, authenticity, and purposefulness...\n\nPlease wait.";
    copyBtn.classList.add('hidden');
    downloadBtn.classList.add('hidden');

    // Constructing the Prompt and System Instruction
    const systemPrompt = `You are a strict, expert Curriculum Evaluator and Instructional Designer. 
You are provided with a 'Diagnostic Framework/Guidelines' and a 'Teacher Document'.

Your task is to:
1. EVALUATE the teacher document strictly against the framework. Specifically, check how accurate, authentic, and purposeful the texts and questioning are.
   - "Accurate" means factual correctness and alignment with learning objectives.
   - "Authentic" means the text feels like real-world material (e.g., articles, primary sources) rather than contrived "school text," preserving the intended tone and format.
   - "Purposeful" means the questioning drives deeper thinking and explicitly addresses the cognitive demands of the target grade level.
2. CRITIQUE the document. Identify the target grade level based on the context. Point out specific violations of the guidelines regarding accuracy, authenticity, and purposefulness.
3. FIX the document. Completely rewrite the texts and questions so they perfectly align with the framework.
   - MANDATORY: You must maintain the authenticity of the text while adjusting the vocabulary, sentence complexity, and cognitive demand to be perfectly appropriate for the identified grade level.
   - Ensure the content is neither too simplistic nor too advanced for the student age group, but remains engaging and rigorous.
   - The rewritten document MUST be highly organized and completely match the structure, format, and layout expectations of a professional diagnostic test.
   - Use semantic structure (e.g., Markdown headers, bullet points, numbered lists for questions) to clearly separate sections, passages, and questions.
   - Do NOT just provide a summary of what to change. You must provide the FULL, ready-to-use rewritten document.

Format your response exactly like this:

=== CRITIQUE ===
Target Grade Level: [Identified Grade]
Accuracy: [Evaluation]
Authenticity: [Evaluation]
Purposefulness: [Evaluation]
Overall Critique: [Summary of violations and necessary fixes]

=== FIXED DOCUMENT ===
[The entirely rewritten, highly organized, and ready-to-use diagnostic document for the teacher, maintaining authenticity while tailored to the specific grade level. Ensure proper structure using headings, lists, and spacing.]`;

    const apiKey = localStorage.getItem('geminiApiKey');
    if (!apiKey) {
        showMessageBox("Error", "Please enter and save your Gemini API Key in the top right corner.");
        
        analyzeBtn.disabled = false;
        analyzeBtn.classList.remove('opacity-75', 'cursor-not-allowed');
        loadingSpinner.classList.add('hidden');
        resultsArea.innerHTML = "Results will appear here...";
        return;
    }

    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

    const parts = [];

    // Add framework part (text or PDF inline data)
    parts.push({ text: "DIAGNOSTIC FRAMEWORK:\n" });
    if (extractedFrameworkBase64) {
        parts.push({
            inlineData: {
                mimeType: extractedFrameworkMimeType,
                data: extractedFrameworkBase64
            }
        });
    } else {
        parts.push({ text: frameworkText });
    }

    parts.push({ text: "\n\n---\n\nTEACHER DOCUMENT:\n" });

    // Add document part (text or PDF inline data)
    if (extractedDocumentBase64) {
        parts.push({
            inlineData: {
                mimeType: extractedDocumentMimeType,
                data: extractedDocumentBase64
            }
        });
    } else {
        parts.push({ text: documentText });
    }

    const payload = {
        contents: [{ parts: parts }],
        systemInstruction: { parts: [{ text: systemPrompt }] }
    };

    try {
        const response = await fetchWithRetry(apiUrl, payload);
        const result = await response.json();
        
        if (result.candidates && result.candidates.length > 0 && result.candidates[0].content) {
            const generatedText = result.candidates[0].content.parts[0].text;
            resultsArea.innerHTML = escapeHTML(generatedText);
            copyBtn.classList.remove('hidden');
            downloadBtn.classList.remove('hidden');
        } else {
            resultsArea.innerHTML = "Error: Unexpected response format from the API.";
        }
    } catch (error) {
        console.error("API Error:", error);
        resultsArea.innerHTML = `Error connecting to the AI: ${error.message}`;
    } finally {
        // Restore UI state
        analyzeBtn.disabled = false;
        analyzeBtn.classList.remove('opacity-75', 'cursor-not-allowed');
        loadingSpinner.classList.add('hidden');
    }
}

// Helper to implement exponential backoff as per instructions
async function fetchWithRetry(url, payload, retries = 3, delay = 1000) {
    try {
        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        
        if (!response.ok) {
            if (response.status === 429 && retries > 0) {
                await new Promise(resolve => setTimeout(resolve, delay));
                return fetchWithRetry(url, payload, retries - 1, delay * 2);
            }
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response;
    } catch (error) {
        if (retries > 0) {
            await new Promise(resolve => setTimeout(resolve, delay));
            return fetchWithRetry(url, payload, retries - 1, delay * 2);
        }
        throw error;
    }
}

function copyResults() {
    const resultsText = document.getElementById('resultsArea').innerText;
    // Extract just the fixed document if possible, otherwise copy all
    let textToCopy = resultsText;
    if (resultsText.includes("=== FIXED DOCUMENT ===")) {
        textToCopy = resultsText.split("=== FIXED DOCUMENT ===")[1].trim();
    }
    
    // Fallback clipboard copying method to bypass iFrame restrictions
    const textArea = document.createElement("textarea");
    textArea.value = textToCopy;
    document.body.appendChild(textArea);
    textArea.select();
    try {
        document.execCommand('copy');
        showMessageBox("Success", "Fixed document copied to clipboard!");
    } catch (err) {
        showMessageBox("Error", "Failed to copy text.");
    }
    document.body.removeChild(textArea);
}

function downloadFixedDocument() {
    const resultsText = document.getElementById('resultsArea').innerText;
    // Extract just the fixed document portion
    let textToSave = resultsText;
    if (resultsText.includes("=== FIXED DOCUMENT ===")) {
        textToSave = resultsText.split("=== FIXED DOCUMENT ===")[1].trim();
    }

    // Use the docx library to create a clean Word Document
    const { Document, Packer, Paragraph, TextRun } = docx;

    const paragraphs = textToSave.split('\n').map(line => {
        return new Paragraph({
            children: [new TextRun(line)],
        });
    });

    const doc = new Document({
        sections: [{
            properties: {},
            children: paragraphs,
        }],
    });

    Packer.toBlob(doc).then(blob => {
        let baseName = originalTeacherFileName;
        const lastDotIndex = baseName.lastIndexOf('.');
        if (lastDotIndex !== -1) {
            baseName = baseName.substring(0, lastDotIndex);
        }
        saveAs(blob, `Fixed_${baseName}.docx`);
    });
}

function escapeHTML(str) {
    return str.replace(/[&<>'"]/g, 
        tag => ({
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            "'": '&#39;',
            '"': '&quot;'
        }[tag] || tag)
    );
}

// Custom message box to replace alert()
function showMessageBox(title, message) {
    const overlay = document.createElement('div');
    overlay.className = 'fixed inset-0 bg-slate-900 bg-opacity-50 flex justify-center items-center z-50';
    
    const box = document.createElement('div');
    box.className = 'bg-white p-6 rounded-xl shadow-xl max-w-sm w-full mx-4';
    
    const titleEl = document.createElement('h3');
    titleEl.className = 'text-lg font-bold text-slate-900 mb-2';
    titleEl.innerText = title;
    
    const msgEl = document.createElement('p');
    msgEl.className = 'text-slate-600 mb-6';
    msgEl.innerText = message;
    
    const btn = document.createElement('button');
    btn.className = 'w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-2 rounded-lg transition-colors';
    btn.innerText = 'OK';
    btn.onclick = () => document.body.removeChild(overlay);
    
    box.appendChild(titleEl);
    box.appendChild(msgEl);
    box.appendChild(btn);
    overlay.appendChild(box);
    document.body.appendChild(overlay);
}
    </script>
</body>
</html>
