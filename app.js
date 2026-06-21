// Import docx as ES module from esm.sh CDN
import * as docx from 'https://esm.sh/docx@8.5.0';

// Ensure Lucide icons are initialized on load
window.addEventListener('DOMContentLoaded', () => {
    lucide.createIcons();
    loadSavedApiKey();
    setupEventListeners();
});

// App State
let uploadedFiles = [];
let generatedData = null;

// DOM Elements
const apiKeyInput = document.getElementById('apiKey');
const toggleApiKeyBtn = document.getElementById('toggleApiKey');
const saveKeyCheckbox = document.getElementById('saveKey');
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const uploadedFilesContainer = document.getElementById('uploadedFilesContainer');
const fileCountSpan = document.getElementById('fileCount');
const thumbnailsGrid = document.getElementById('thumbnailsGrid');
const modelSelect = document.getElementById('modelSelect');
const customPromptTextarea = document.getElementById('customPrompt');
const generateBtn = document.getElementById('generateBtn');
const statusCard = document.getElementById('statusCard');
const progressBarFill = document.getElementById('progressBarFill');
const logsConsole = document.getElementById('logsConsole');
const downloadsCard = document.getElementById('downloadsCard');
const emptyState = document.getElementById('emptyState');
const downloadEnBtn = document.getElementById('downloadEnBtn');
const downloadHiBtn = document.getElementById('downloadHiBtn');
const previewTrigger = document.getElementById('previewTrigger');
const previewContent = document.getElementById('previewContent');

// 1. Setup & Event Listeners
function setupEventListeners() {
    // API key show/hide
    toggleApiKeyBtn.addEventListener('click', () => {
        const type = apiKeyInput.type === 'password' ? 'text' : 'password';
        apiKeyInput.type = type;
        const iconName = type === 'password' ? 'eye' : 'eye-off';
        toggleApiKeyBtn.innerHTML = `<i data-lucide="${iconName}"></i>`;
        lucide.createIcons({ attrs: { class: 'lucide' } });
    });

    // Check generate button state on key input or file selection
    apiKeyInput.addEventListener('input', checkButtonState);

    // Dropzone events
    dropzone.addEventListener('click', () => fileInput.click());
    
    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('drag-active');
    });
    
    dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('drag-active');
    });
    
    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('drag-active');
        handleSelectedFiles(e.dataTransfer.files);
    });
    
    fileInput.addEventListener('change', () => {
        handleSelectedFiles(fileInput.files);
    });

    // Generation Action
    generateBtn.addEventListener('click', startGenerationFlow);

    // Preview accordion
    previewTrigger.addEventListener('click', () => {
        previewTrigger.classList.toggle('active');
        const isVisible = previewContent.style.display === 'block';
        previewContent.style.display = isVisible ? 'none' : 'block';
    });

    // Tab buttons
    const tabBtns = document.querySelectorAll('.tab-btn');
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            tabBtns.forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
            
            btn.classList.add('active');
            const tabId = btn.getAttribute('data-tab');
            document.getElementById(tabId).classList.add('active');
        });
    });

    // Downloads
    downloadEnBtn.addEventListener('click', () => generateAndDownloadDocx('english'));
    downloadHiBtn.addEventListener('click', () => generateAndDownloadDocx('hindi'));
}

// Check if we can enable Generate button
function checkButtonState() {
    const hasKey = apiKeyInput.value.trim().length > 0;
    const hasFiles = uploadedFiles.length > 0;
    generateBtn.disabled = !(hasKey && hasFiles);
}

// Load API key from local storage
function loadSavedApiKey() {
    const savedKey = localStorage.getItem('gemini_api_key');
    if (savedKey) {
        apiKeyInput.value = savedKey;
        saveKeyCheckbox.checked = true;
    } else {
        apiKeyInput.value = "";
        saveKeyCheckbox.checked = false;
    }
}

// 2. Handle Uploaded Files
function handleSelectedFiles(files) {
    const validFiles = Array.from(files).filter(file => {
        const isImage = file.type.startsWith('image/');
        const isPdf = file.type === 'application/pdf';
        const isDocx = file.name.endsWith('.docx') || file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
        return isImage || isPdf || isDocx;
    });
    
    if (validFiles.length === 0) {
        addLog('No valid files (images, PDFs, or .docx) selected.', 'error');
        return;
    }

    validFiles.forEach(file => {
        // Prevent duplicate loads
        if (uploadedFiles.some(f => f.name === file.name && f.size === file.size)) return;
        
        // Add to array
        uploadedFiles.push(file);
        
        // Render thumbnail
        if (file.type.startsWith('image/')) {
            const reader = new FileReader();
            reader.onload = (e) => {
                createThumbnail(file, e.target.result);
            };
            reader.readAsDataURL(file);
        } else {
            // For PDFs and Word documents, we don't need base64 data url for the thumbnail
            createThumbnail(file, null);
        }
    });

    updateFilesUI();
    checkButtonState();
    addLog(`Loaded ${validFiles.length} file(s). Total: ${uploadedFiles.length}`, 'info');
}

function createThumbnail(file, src) {
    const wrapper = document.createElement('div');
    wrapper.className = 'thumbnail-wrapper';
    
    if (file.type.startsWith('image/')) {
        const img = document.createElement('img');
        img.src = src;
        img.alt = file.name;
        wrapper.appendChild(img);
    } else {
        const isPdf = file.type === 'application/pdf' || file.name.endsWith('.pdf');
        const iconName = isPdf ? 'file-text' : 'file';
        const bgClass = isPdf ? 'pdf-bg' : 'docx-bg';
        
        const placeholder = document.createElement('div');
        placeholder.className = `file-thumbnail-placeholder ${bgClass}`;
        placeholder.innerHTML = `
            <i data-lucide="${iconName}"></i>
            <span class="file-ext">${isPdf ? 'PDF' : 'DOCX'}</span>
            <span class="file-name-tooltip">${file.name}</span>
        `;
        wrapper.appendChild(placeholder);
    }
    
    const removeBtn = document.createElement('button');
    removeBtn.className = 'btn-remove-file';
    removeBtn.innerHTML = '<i data-lucide="x"></i>';
    removeBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        removeFile(file, wrapper);
    });

    wrapper.appendChild(removeBtn);
    thumbnailsGrid.appendChild(wrapper);
    
    lucide.createIcons({ attrs: { class: 'lucide' } });
}

function removeFile(file, wrapper) {
    uploadedFiles = uploadedFiles.filter(f => f !== file);
    wrapper.remove();
    updateFilesUI();
    checkButtonState();
    addLog(`Removed file: ${file.name}`, 'info');
}

function updateFilesUI() {
    if (uploadedFiles.length > 0) {
        uploadedFilesContainer.style.display = 'block';
        fileCountSpan.textContent = uploadedFiles.length;
    } else {
        uploadedFilesContainer.style.display = 'none';
        fileCountSpan.textContent = '0';
    }
}

// 3. Logger/Console Utilities
function addLog(message, type = 'info') {
    const logLine = document.createElement('div');
    logLine.className = `log-line ${type}`;
    
    if (type === 'loading') {
        logLine.textContent = message;
    } else {
        const prefix = type === 'error' ? '❌ ' : type === 'success' ? '✅ ' : '• ';
        logLine.textContent = prefix + message;
    }
    
    logsConsole.appendChild(logLine);
    logsConsole.scrollTop = logsConsole.scrollHeight;
}

function clearLogs() {
    logsConsole.innerHTML = '';
}

function setProgress(percentage) {
    progressBarFill.style.width = `${percentage}%`;
}

// Helper to convert file to base64
function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = error => reject(error);
    });
}

// Helper to extract plain text from .docx file using mammoth.js
function extractTextFromDocx(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(event) {
            const arrayBuffer = event.target.result;
            if (typeof mammoth === 'undefined') {
                reject(new Error("Mammoth library is not loaded. Cannot parse Word files. Please check your internet connection."));
                return;
            }
            mammoth.extractRawText({ arrayBuffer: arrayBuffer })
                .then(function(result) {
                    resolve(result.value);
                })
                .catch(function(err) {
                    reject(err);
                });
        };
        reader.onerror = function(err) {
            reject(err);
        };
        reader.readAsArrayBuffer(file);
    });
}

// 4. Gemini API Call
async function startGenerationFlow() {
    const apiKey = apiKeyInput.value.trim();
    if (!apiKey) {
        alert("Please enter a valid Gemini API key!");
        return;
    }

    if (saveKeyCheckbox.checked) {
        localStorage.setItem('gemini_api_key', apiKey);
    } else {
        localStorage.removeItem('gemini_api_key');
    }

    // Update UI elements
    statusCard.style.display = 'block';
    emptyState.style.display = 'none';
    downloadsCard.style.display = 'none';
    generateBtn.disabled = true;
    setProgress(10);
    clearLogs();
    
    addLog('Starting compilation process...', 'info');

    try {
        // Step 1: Base64 encode images & PDFs, and extract text from DOCX
        setProgress(25);
        addLog('Parsing uploaded files...', 'loading');
        
        const mediaParts = [];
        const extractedTexts = [];
        
        for (let file of uploadedFiles) {
            const isImage = file.type.startsWith('image/');
            const isPdf = file.type === 'application/pdf';
            const isDocx = file.name.endsWith('.docx') || file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
            
            if (isImage) {
                const base64 = await fileToBase64(file);
                mediaParts.push({
                    inlineData: {
                        mimeType: file.type,
                        data: base64
                    }
                });
                addLog(`Encoded image: ${file.name} (${Math.round(file.size / 1024)} KB)`, 'info');
            } else if (isPdf) {
                const base64 = await fileToBase64(file);
                mediaParts.push({
                    inlineData: {
                        mimeType: 'application/pdf',
                        data: base64
                    }
                });
                addLog(`Encoded PDF document: ${file.name} (${Math.round(file.size / 1024)} KB)`, 'info');
            } else if (isDocx) {
                addLog(`Extracting text from Word document: ${file.name}...`, 'loading');
                const docText = await extractTextFromDocx(file);
                extractedTexts.push({
                    name: file.name,
                    text: docText
                });
                addLog(`Extracted text from ${file.name} (${docText.length} characters)`, 'success');
            }
        }
        
        // Step 2: Form prompt and call Gemini API
        setProgress(50);
        addLog('Calling Google Gemini API model...', 'loading');
        const model = modelSelect.value;
        
        const responseJson = await callGeminiAPI(apiKey, model, mediaParts, extractedTexts);
        
        // Step 3: Handle API success
        setProgress(85);
        addLog('Parsing bilingual response JSON...', 'success');
        
        generatedData = responseJson;
        
        // Render preview in UI
        renderPreview();
        
        setProgress(100);
        addLog('Question generation completed!', 'success');
        addLog('DOCX files ready for download.', 'success');
        
        // Show download card
        downloadsCard.style.display = 'block';
        statusCard.querySelector('.spinner').classList.remove('spinner');
        statusCard.querySelector('.card-header i').innerHTML = '<i data-lucide="check-circle" class="text-success"></i>';
        lucide.createIcons();

    } catch (error) {
        setProgress(0);
        addLog(error.message || 'An unexpected error occurred.', 'error');
        generateBtn.disabled = false;
        console.error(error);
    }
}

async function callGeminiAPI(apiKey, model, mediaParts, extractedTexts) {
    const customPrompt = customPromptTextarea.value.trim();
    
    // Schema definition in uppercase as expected by Gemini API
    const responseSchema = {
        type: "OBJECT",
        properties: {
            english: {
                type: "OBJECT",
                properties: {
                    mcqs: {
                        type: "ARRAY",
                        description: "List of exactly 10 multiple choice questions in English",
                        items: {
                            type: "OBJECT",
                            properties: {
                                question: { type: "STRING" },
                                options: {
                                    type: "ARRAY",
                                    items: { type: "STRING" }
                                },
                                correctAnswer: { type: "STRING", description: "The correct option text exactly matching one of the options" }
                            },
                            required: ["question", "options", "correctAnswer"]
                        }
                    },
                    shortAnswers: {
                        type: "ARRAY",
                        description: "List of exactly 10 short answer questions in English",
                        items: {
                            type: "OBJECT",
                            properties: {
                                question: { type: "STRING" },
                                answer: { type: "STRING" }
                            },
                            required: ["question", "answer"]
                        }
                    }
                },
                required: ["mcqs", "shortAnswers"]
            },
            hindi: {
                type: "OBJECT",
                properties: {
                    mcqs: {
                        type: "ARRAY",
                        description: "List of exactly 10 multiple choice questions in Hindi (exact translation of English set)",
                        items: {
                            type: "OBJECT",
                            properties: {
                                question: { type: "STRING" },
                                options: {
                                    type: "ARRAY",
                                    items: { type: "STRING" }
                                },
                                correctAnswer: { type: "STRING", description: "The correct option text exactly matching one of the Hindi options" }
                            },
                            required: ["question", "options", "correctAnswer"]
                        }
                    },
                    shortAnswers: {
                        type: "ARRAY",
                        description: "List of exactly 10 short answer questions in Hindi (exact translation of English set)",
                        items: {
                            type: "OBJECT",
                            properties: {
                                question: { type: "STRING" },
                                answer: { type: "STRING" }
                            },
                            required: ["question", "answer"]
                        }
                    }
                },
                required: ["mcqs", "shortAnswers"]
            }
        },
        required: ["english", "hindi"]
    };

    const parts = [...mediaParts];
    
    // Add extracted texts from DOCX files
    extractedTexts.forEach(doc => {
        parts.push({
            text: `\n\n[Content of Word document "${doc.name}":]\n${doc.text}\n`
        });
    });
    
    // Add custom instructions prompt
    parts.push({
        text: `${customPrompt}\n\nStrict requirement: Generate exactly 10 MCQs and 10 Short Answer Questions. Ensure the Hindi set matches the English set in context and order (100% translation parity). Return the output using the specified JSON schema.`
    });

    const requestBody = {
        contents: [
            {
                parts: parts
            }
        ],
        generationConfig: {
            responseMimeType: "application/json",
            responseSchema: responseSchema
        }
    };

    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const response = await fetch(apiUrl, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        const errorMessage = errorData.error?.message || `HTTP ${response.status} Error`;
        throw new Error(`Gemini API Error: ${errorMessage}`);
    }

    const data = await response.json();
    
    // Extract textual content containing JSON
    try {
        const textContent = data.candidates[0].content.parts[0].text;
        const parsedJson = JSON.parse(textContent);
        
        // Simple sanity check of contents
        if (!parsedJson.english || !parsedJson.hindi || !parsedJson.english.mcqs || !parsedJson.english.shortAnswers) {
            throw new Error("Invalid structure returned from Gemini API.");
        }
        
        return parsedJson;
    } catch (err) {
        console.error("Failed to parse Gemini response text as JSON:", err);
        throw new Error("Gemini returned successfully, but failed to format the response into proper JSON structure. Please try again.");
    }
}

// 5. Render Live Preview
function renderPreview() {
    if (!generatedData) return;

    const { english, hindi } = generatedData;
    
    // EN MCQs
    const enMcqContainer = document.getElementById('previewEnMcqs');
    enMcqContainer.innerHTML = '';
    english.mcqs.forEach((item, index) => {
        const card = document.createElement('div');
        card.className = 'preview-q-card';
        card.innerHTML = `
            <div class="q-title">${index + 1}. ${item.question}</div>
            <ul class="q-options">
                ${item.options.map(opt => `<li class="q-option">${opt}</li>`).join('')}
            </ul>
            <div class="q-answer">
                <i data-lucide="check" class="text-success" style="width:14px;height:14px;"></i>
                <span>Correct: ${item.correctAnswer}</span>
            </div>
        `;
        enMcqContainer.appendChild(card);
    });

    // EN Short Qs
    const enShortContainer = document.getElementById('previewEnShorts');
    enShortContainer.innerHTML = '';
    english.shortAnswers.forEach((item, index) => {
        const card = document.createElement('div');
        card.className = 'preview-q-card';
        card.innerHTML = `
            <div class="q-title">${index + 1}. ${item.question}</div>
            <div class="q-answer-short"><strong>Ans:</strong> ${item.answer}</div>
        `;
        enShortContainer.appendChild(card);
    });

    // HI MCQs
    const hiMcqContainer = document.getElementById('previewHiMcqs');
    hiMcqContainer.innerHTML = '';
    hindi.mcqs.forEach((item, index) => {
        const card = document.createElement('div');
        card.className = 'preview-q-card';
        card.innerHTML = `
            <div class="q-title">${index + 1}. ${item.question}</div>
            <ul class="q-options">
                ${item.options.map(opt => `<li class="q-option">${opt}</li>`).join('')}
            </ul>
            <div class="q-answer">
                <i data-lucide="check" class="text-success" style="width:14px;height:14px;"></i>
                <span>सही उत्तर: ${item.correctAnswer}</span>
            </div>
        `;
        hiMcqContainer.appendChild(card);
    });

    // HI Short Qs
    const hiShortContainer = document.getElementById('previewHiShorts');
    hiShortContainer.innerHTML = '';
    hindi.shortAnswers.forEach((item, index) => {
        const card = document.createElement('div');
        card.className = 'preview-q-card';
        card.innerHTML = `
            <div class="q-title">${index + 1}. ${item.question}</div>
            <div class="q-answer-short"><strong>उत्तर:</strong> ${item.answer}</div>
        `;
        hiShortContainer.appendChild(card);
    });

    lucide.createIcons();
}

// 6. Generate DOCX with docx.js (UMD)
async function generateAndDownloadDocx(language) {
    if (!generatedData) {
        alert("No generated data found!");
        return;
    }

    try {
        const dataset = generatedData[language];
        const isEnglish = language === 'english';
        
        const titleText = isEnglish ? "ASSESSMENT QUESTION BANK" : "मूल्यांकन प्रश्न बैंक";
        const mcqHeadingText = isEnglish ? "Section A: Multiple Choice Questions (MCQs)" : "भाग अ: बहुविकल्पीय प्रश्न (MCQs)";
        const shortHeadingText = isEnglish ? "Section B: Short Answer Questions" : "भाग ब: लघु उत्तरीय प्रश्न";
        const answerLabel = isEnglish ? "Correct Answer" : "सही उत्तर";
        const modelLabel = isEnglish ? "Generated via Gemini AI model" : "जेमिनी एआई मॉडल द्वारा उत्पन्न";

        const { Document, Paragraph, TextRun, PageBreak, HeadingLevel, AlignmentType } = docx;

        const children = [];

        // Main Document Header
        children.push(
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 120 },
                children: [
                    new TextRun({
                        text: titleText,
                        bold: true,
                        size: 36, // 18pt
                        font: "Calibri",
                        color: "1A365D" // Deep blue color
                    })
                ]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 360 },
                children: [
                    new TextRun({
                        text: `${modelLabel} • ${new Date().toLocaleDateString()}`,
                        size: 20, // 10pt
                        font: "Calibri",
                        color: "718096" // Slate gray
                    })
                ]
            })
        );

        // --- PAGE 1: MCQs ---
        children.push(
            new Paragraph({
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 240, after: 180 },
                children: [
                    new TextRun({
                        text: mcqHeadingText,
                        bold: true,
                        size: 28, // 14pt
                        font: "Calibri",
                        color: "2B6CB0"
                    })
                ]
            })
        );

        dataset.mcqs.forEach((mcq, idx) => {
            // Question text
            children.push(
                new Paragraph({
                    spacing: { before: 120, after: 80 },
                    children: [
                        new TextRun({
                            text: `${idx + 1}. `,
                            bold: true,
                            font: "Calibri",
                            size: 24 // 12pt
                        }),
                        new TextRun({
                            text: mcq.question,
                            font: "Calibri",
                            size: 24 // 12pt
                        })
                    ]
                })
            );

            // Options (Indented)
            const letters = ['A', 'B', 'C', 'D'];
            mcq.options.forEach((opt, optIdx) => {
                children.push(
                    new Paragraph({
                        indent: { left: 720 }, // 0.5 inches indentation
                        spacing: { after: 40 },
                        children: [
                            new TextRun({
                                text: `(${letters[optIdx]}) `,
                                bold: true,
                                font: "Calibri",
                                size: 22 // 11pt
                            }),
                            new TextRun({
                                text: opt,
                                font: "Calibri",
                                size: 22 // 11pt
                            })
                        ]
                    })
                );
            });

            // Answer
            children.push(
                new Paragraph({
                    indent: { left: 720 },
                    spacing: { after: 180 },
                    children: [
                        new TextRun({
                            text: `${answerLabel}: `,
                            bold: true,
                            font: "Calibri",
                            size: 22,
                            color: "2F855A" // Dark green
                        }),
                        new TextRun({
                            text: mcq.correctAnswer,
                            font: "Calibri",
                            size: 22,
                            color: "2F855A"
                        })
                    ]
                })
            );
        });

        // PAGE BREAK
        children.push(new PageBreak());

        // --- PAGE 2: SHORT ANSWER QUESTIONS ---
        children.push(
            new Paragraph({
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 240, after: 180 },
                children: [
                    new TextRun({
                        text: shortHeadingText,
                        bold: true,
                        size: 28, // 14pt
                        font: "Calibri",
                        color: "2B6CB0"
                    })
                ]
            })
        );

        dataset.shortAnswers.forEach((sa, idx) => {
            // Question
            children.push(
                new Paragraph({
                    spacing: { before: 120, after: 80 },
                    children: [
                        new TextRun({
                            text: `Q${idx + 1}. `,
                            bold: true,
                            font: "Calibri",
                            size: 24 // 12pt
                        }),
                        new TextRun({
                            text: sa.question,
                            font: "Calibri",
                            size: 24 // 12pt
                        })
                    ]
                })
            );

            // Answer
            children.push(
                new Paragraph({
                    indent: { left: 720 },
                    spacing: { after: 180 },
                    children: [
                        new TextRun({
                            text: isEnglish ? "Answer: " : "उत्तर: ",
                            bold: true,
                            font: "Calibri",
                            size: 22,
                            color: "4A5568"
                        }),
                        new TextRun({
                            text: sa.answer,
                            font: "Calibri",
                            size: 22,
                            color: "2D3748"
                        })
                    ]
                })
            );
        });

        // Initialize document with the children list inside a single section
        const doc = new Document({
            sections: [{
                properties: {},
                children: children
            }]
        });

        // Pack the document and download
        const blob = await docx.Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement("a");
        a.href = url;
        a.download = isEnglish ? "assessment_english.docx" : "assessment_hindi.docx";
        document.body.appendChild(a);
        a.click();
        
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        addLog(`Successfully generated and downloaded ${language} document.`, 'success');

    } catch (err) {
        console.error("DOCX generation error:", err);
        alert(`Failed to generate DOCX file: ${err.message}`);
    }
}
