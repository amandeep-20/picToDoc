// docx is loaded via UMD CDN script tag in index.html and is globally available as window.docx
const docx = window.docx;

// Resolve the API key from localStorage, falling back to window.GEMINI_API_KEY (from config.js)
let GEMINI_API_KEY = localStorage.getItem('gemini_api_key') || (window.GEMINI_API_KEY || "");



// App State
let uploadedFiles = [];
let generatedData = null;
let currentMode = 'single'; // 'single' or 'fa'

// DOM Elements
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const uploadedFilesContainer = document.getElementById('uploadedFilesContainer');
const fileCountSpan = document.getElementById('fileCount');
const thumbnailsGrid = document.getElementById('thumbnailsGrid');
const modelSelect = document.getElementById('modelSelect');
const customPromptTextarea = document.getElementById('customPrompt');
const promptInputGroup = document.getElementById('promptInputGroup');
const faSettingsContainer = document.getElementById('faSettingsContainer');
const generateBtn = document.getElementById('generateBtn');
const statusCard = document.getElementById('statusCard');
const progressBarFill = document.getElementById('progressBarFill');
const logsConsole = document.getElementById('logsConsole');
const resultsCard = document.getElementById('resultsCard');
const emptyState = document.getElementById('emptyState');
const downloadEnBtn = document.getElementById('downloadEnBtn');
const downloadHiBtn = document.getElementById('downloadHiBtn');

// Mode Switcher Elements
const modeSingleBtn = document.getElementById('modeSingleBtn');
const modeFaBtn = document.getElementById('modeFaBtn');

// FA Header Inputs
const faExamName = document.getElementById('faExamName');
const faSession = document.getElementById('faSession');
const faClass = document.getElementById('faClass');
const faSubject = document.getElementById('faSubject');
const faTime = document.getElementById('faTime');
const faMaxMarks = document.getElementById('faMaxMarks');
const faMcqCount = document.getElementById('faMcqCount');
const faMcqMarks = document.getElementById('faMcqMarks');
const faShortTotal = document.getElementById('faShortTotal');
const faShortLimit = document.getElementById('faShortLimit');
const faShortMarks = document.getElementById('faShortMarks');
const faLongHeader = document.getElementById('faLongHeader');
const faLongMarks = document.getElementById('faLongMarks');

// Settings Modal DOM Elements
const settingsBtn = document.getElementById('settingsBtn');
const settingsModal = document.getElementById('settingsModal');
const modalApiKey = document.getElementById('modalApiKey');
const toggleModalApiKey = document.getElementById('toggleModalApiKey');
const closeSettingsBtn = document.getElementById('closeSettingsBtn');
const saveSettingsBtn = document.getElementById('saveSettingsBtn');

// 1. Setup & Event Listeners
function setupEventListeners() {
    // Dropzone events
    
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

    // Reset Button (Start Over)
    const resetBtn = document.getElementById('resetBtn');
    if (resetBtn) {
        resetBtn.addEventListener('click', () => {
            if (confirm("Are you sure you want to start over? All current generated questions will be deleted.")) {
                generatedData = null;
                localStorage.removeItem('generated_paper_data');
                resultsCard.style.display = 'none';
                emptyState.style.display = 'flex';
                clearLogs();
                // Reset file lists
                uploadedFiles = [];
                thumbnailsGrid.innerHTML = '';
                updateFilesUI();
                checkButtonState();
            }
        });
    }

    // Settings Modal handlers
    if (settingsBtn) {
        settingsBtn.addEventListener('click', () => {
            modalApiKey.value = GEMINI_API_KEY;
            settingsModal.style.display = 'flex';
        });
    }

    if (closeSettingsBtn) {
        closeSettingsBtn.addEventListener('click', () => {
            settingsModal.style.display = 'none';
        });
    }

    if (saveSettingsBtn) {
        saveSettingsBtn.addEventListener('click', () => {
            const keyVal = modalApiKey.value.trim();
            if (!keyVal) {
                if (!confirm("Are you sure you want to clear your API key? The app will not work without one.")) {
                    return;
                }
            }
            GEMINI_API_KEY = keyVal;
            localStorage.setItem('gemini_api_key', keyVal);
            settingsModal.style.display = 'none';
            checkButtonState();
            addLog("Gemini API key updated and saved locally.", "success");
        });
    }

    if (toggleModalApiKey) {
        toggleModalApiKey.addEventListener('click', () => {
            const type = modalApiKey.type === 'password' ? 'text' : 'password';
            modalApiKey.type = type;
            const iconName = type === 'password' ? 'eye' : 'eye-off';
            toggleModalApiKey.innerHTML = `<i data-lucide="${iconName}"></i>`;
            lucide.createIcons({ attrs: { class: 'lucide' } });
        });
    }

    // Mode Switcher handlers
    if (modeSingleBtn && modeFaBtn) {
        modeSingleBtn.addEventListener('click', () => switchMode('single'));
        modeFaBtn.addEventListener('click', () => switchMode('fa'));
    }
}

function switchMode(mode) {
    if (currentMode === mode) return;
    currentMode = mode;
    
    // Toggle active classes on tab buttons
    if (mode === 'single') {
        modeSingleBtn.classList.add('active');
        modeFaBtn.classList.remove('active');
        promptInputGroup.style.display = 'flex';
        faSettingsContainer.style.display = 'none';
        
        // Update texts
        document.querySelector('.card-upload .card-header h2').textContent = "Source Materials (Images / PDFs / Word)";
        document.querySelector('.dropzone-subtext').textContent = "Upload study pages, images (JPG/PNG), PDFs, or Word docs (.docx) (Max 50MB per file)";
        fileInput.setAttribute('accept', 'image/*,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        generateBtn.querySelector('span').textContent = "Generate Bilingual Documents";
        
        // Update empty state text
        emptyState.querySelector('p').textContent = "Enter your Gemini API key, upload 3-4 images/PDFs/Word docs, and click \"Generate Bilingual Documents\" to start the process.";
    } else {
        modeFaBtn.classList.add('active');
        modeSingleBtn.classList.remove('active');
        promptInputGroup.style.display = 'none';
        faSettingsContainer.style.display = 'flex';
        
        // Update texts
        document.querySelector('.card-upload .card-header h2').textContent = "Source Generated Documents (.docx / .pdf / .xlsx / .xls)";
        document.querySelector('.dropzone-subtext').textContent = "Upload previously generated documents (.docx, .pdf, .xlsx, .xls) (Max 50MB per file)";
        fileInput.setAttribute('accept', '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/pdf,.xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel');
        generateBtn.querySelector('span').textContent = "Generate FA Exam Paper";
        
        // Update empty state text
        emptyState.querySelector('p').textContent = "Enter your Gemini API key, upload 4-5 generated files (.docx, .pdf, .xlsx, .xls), fill the exam metadata, and click \"Generate FA Exam Paper\" to start.";
    }

    // Reset current file list and results state
    uploadedFiles = [];
    thumbnailsGrid.innerHTML = '';
    updateFilesUI();
    checkButtonState();
    
    // Clear results/logs state for the new mode
    generatedData = null;
    resultsCard.style.display = 'none';
    emptyState.style.display = 'flex';
    clearLogs();
}

// Check if we can enable Generate button
function checkButtonState() {
    const hasFiles = uploadedFiles.length > 0;
    generateBtn.disabled = !hasFiles;
}

// 2. Handle Uploaded Files
function handleSelectedFiles(files) {
    const maxSizeBytes = 50 * 1024 * 1024; // 50MB
    const validFiles = Array.from(files).filter(file => {
        const isImage = file.type.startsWith('image/');
        const isPdf = file.type === 'application/pdf' || file.name.endsWith('.pdf');
        const isDocx = file.name.endsWith('.docx') || file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
        const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || 
                        file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                        file.type === 'application/vnd.ms-excel';
        
        let isValidType = false;
        if (currentMode === 'single') {
            isValidType = isImage || isPdf || isDocx;
        } else { // 'fa' mode
            isValidType = isDocx || isPdf || isExcel;
        }
        
        if (!isValidType) {
            const modeText = currentMode === 'single' ? 'Single mode (Images/PDFs/Word)' : 'FA mode (Word/PDFs/Excel)';
            addLog(`Skipped: '${file.name}' is not a supported file format in ${modeText}.`, 'error');
            return false;
        }
        
        if (file.size > maxSizeBytes) {
            addLog(`Skipped: '${file.name}' exceeds the 50MB file size limit (${Math.round(file.size / (1024 * 1024))}MB).`, 'error');
            return false;
        }
        
        return true;
    });
    
    if (validFiles.length === 0) {
        addLog('No valid files under 50MB selected.', 'error');
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
        const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || 
                        file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                        file.type === 'application/vnd.ms-excel';
        
        let iconName = 'file';
        let bgClass = 'docx-bg';
        let extLabel = 'DOCX';
        
        if (isPdf) {
            iconName = 'file-text';
            bgClass = 'pdf-bg';
            extLabel = 'PDF';
        } else if (isExcel) {
            iconName = 'file-spreadsheet';
            bgClass = 'excel-bg';
            extLabel = 'EXCEL';
        }
        
        const placeholder = document.createElement('div');
        placeholder.className = `file-thumbnail-placeholder ${bgClass}`;
        placeholder.innerHTML = `
            <i data-lucide="${iconName}"></i>
            <span class="file-ext">${extLabel}</span>
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

// Helper to extract plain text from .pdf file using pdf.js
function extractTextFromPdf(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async function(event) {
            const arrayBuffer = event.target.result;
            try {
                if (typeof pdfjsLib === 'undefined') {
                    reject(new Error("PDF.js library is not loaded. Cannot parse PDF files. Please check your internet connection."));
                    return;
                }
                pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';
                const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
                const pdf = await loadingTask.promise;
                let fullText = "";
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(" ");
                    fullText += pageText + "\n";
                }
                resolve(fullText);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = function(err) {
            reject(err);
        };
        reader.readAsArrayBuffer(file);
    });
}

// Helper to extract plain text from Excel files using SheetJS
function extractTextFromExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(event) {
            const arrayBuffer = event.target.result;
            try {
                if (typeof XLSX === 'undefined') {
                    reject(new Error("SheetJS (XLSX) library is not loaded. Cannot parse Excel files. Please check your internet connection."));
                    return;
                }
                const data = new Uint8Array(arrayBuffer);
                const workbook = XLSX.read(data, { type: 'array' });
                let fullText = "";
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const csvText = XLSX.utils.sheet_to_csv(worksheet);
                    fullText += `Sheet: ${sheetName}\n${csvText}\n\n`;
                });
                resolve(fullText);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = function(err) {
            reject(err);
        };
        reader.readAsArrayBuffer(file);
    });
}

// 4. Gemini API Call
// 4. Gemini API Call
async function startGenerationFlow() {
    const apiKey = GEMINI_API_KEY.trim();
    if (!apiKey) {
        alert("Gemini API key is not configured! Please click the settings gear icon to configure it.");
        settingsModal.style.display = 'flex';
        return;
    }

    // Update UI elements
    statusCard.style.display = 'block';
    emptyState.style.display = 'none';
    resultsCard.style.display = 'none';
    generateBtn.disabled = true;
    setProgress(10);
    clearLogs();
    
    addLog('Starting compilation process...', 'info');

    try {
        // Step 1: Parse uploaded files
        setProgress(25);
        addLog('Parsing uploaded files...', 'loading');
        
        const mediaParts = [];
        const extractedTexts = [];
        
        for (let file of uploadedFiles) {
            const isImage = file.type.startsWith('image/');
            const isPdf = file.type === 'application/pdf' || file.name.endsWith('.pdf');
            const isDocx = file.name.endsWith('.docx') || file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
            const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || 
                            file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                            file.type === 'application/vnd.ms-excel';
            
            if (isImage && currentMode === 'single') {
                const base64 = await fileToBase64(file);
                mediaParts.push({
                    name: file.name,
                    inlineData: {
                        mimeType: file.type,
                        data: base64
                    }
                });
                addLog(`Encoded image: ${file.name} (${Math.round(file.size / 1024)} KB)`, 'info');
            } else if (isPdf) {
                const base64 = await fileToBase64(file);
                mediaParts.push({
                    name: file.name,
                    inlineData: {
                        mimeType: 'application/pdf',
                        data: base64
                    }
                });
                addLog(`Encoded PDF document for native processing: ${file.name} (${Math.round(file.size / 1024)} KB)`, 'info');
            } else if (isExcel && currentMode === 'fa') {
                addLog(`Extracting text from Excel spreadsheet: ${file.name}...`, 'loading');
                const excelText = await extractTextFromExcel(file);
                extractedTexts.push({
                    name: file.name,
                    text: excelText
                });
                addLog(`Extracted text from Excel ${file.name} (${excelText.length} characters)`, 'success');
            } else if (isDocx) {
                addLog(`Extracting text from Word document: ${file.name}...`, 'loading');
                const docText = await extractTextFromDocx(file);
                extractedTexts.push({
                    name: file.name,
                    text: docText
                });
                addLog(`Extracted text from Word doc ${file.name} (${docText.length} characters)`, 'success');
            } else {
                addLog(`Skipped: '${file.name}' is not supported in the current mode.`, 'warn');
            }
        }
        
        if (currentMode === 'fa' && extractedTexts.length === 0 && mediaParts.length === 0) {
            throw new Error("No documents successfully parsed. Please upload at least one .docx, .pdf, or Excel question bank document.");
        }

        // Step 2: Form prompt and call Gemini API
        setProgress(50);
        addLog('Calling Google Gemini API model...', 'loading');
        const model = modelSelect.value;
        
        let responseJson;
        if (currentMode === 'single') {
            responseJson = await callGeminiAPI(apiKey, model, mediaParts, extractedTexts);
        } else {
            responseJson = await callGeminiAPIForFaPaper(apiKey, model, mediaParts, extractedTexts);
            
            // Attach examHeader info to save it in local state
            responseJson.examHeader = {
                examName: faExamName.value.trim(),
                session: faSession.value.trim(),
                class: faClass.value.trim(),
                subject: faSubject.value.trim(),
                time: faTime.value.trim(),
                maxMarks: faMaxMarks.value.trim(),
                mcqCount: parseInt(faMcqCount.value) || 6,
                mcqMarks: faMcqMarks.value.trim(),
                shortTotal: parseInt(faShortTotal.value) || 4,
                shortLimit: parseInt(faShortLimit.value) || 2,
                shortMarks: faShortMarks.value.trim(),
                longHeader: faLongHeader.value.trim(),
                longMarks: faLongMarks.value.trim()
            };
        }
        
        // Step 3: Handle API success
        setProgress(85);
        addLog('Parsing bilingual response JSON...', 'success');
        
        generatedData = responseJson;
        
        // Render preview in UI
        renderPreview();
        
        setProgress(100);
        addLog('Question generation completed!', 'success');
        addLog('DOCX files ready for download.', 'success');
        
        saveDataToLocalStorage();

        // Show results card
        resultsCard.style.display = 'block';
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

async function callGeminiAPIForFaPaper(apiKey, model, mediaParts, extractedTexts) {
    const mcqCount = parseInt(faMcqCount.value) || 6;
    const shortTotal = parseInt(faShortTotal.value) || 4;

    const parts = [];

    // Add media files (PDFs)
    mediaParts.forEach((media, idx) => {
        parts.push({
            text: `\n\n--- Source Document (PDF) ${idx + 1}: ${media.name} ---`
        });
        parts.push({
            inlineData: media.inlineData
        });
    });

    let sourceQuestionsText = "";
    extractedTexts.forEach((doc, idx) => {
        sourceQuestionsText += `\n\n--- Source Document (Word/Excel) ${idx + 1}: ${doc.name} ---\n${doc.text}\n`;
    });

    const promptText = `
You are an expert school examination paper compiler. Your task is to generate a consolidated exam paper based ONLY on the questions present in the provided source documents.

Source Documents Content:
${sourceQuestionsText}

Instructions:
1. Extract and select exactly ${mcqCount} Multiple Choice Questions (MCQs) and exactly ${shortTotal} Short Answer Questions from the source documents.
2. Under no circumstances should you invent new questions or introduce external concepts. You must only use/adapt the questions that exist in the source documents.
3. For each selected question, you must compile:
   - An English version.
   - A Hindi version with 100% exact translation parity.
4. For MCQs, ensure that:
   - The correct answer matches one of the options.
   - The translation is natural and accurate in both languages.
5. In the questions and answers, do not include references to chapter numbers or names (e.g., 'From Chapter 1', 'According to Chapter...', 'Acc to chapter...', etc.).
6. Return the result in the exact JSON schema provided below.
`;

    parts.push({ text: promptText });

    const responseSchema = {
        type: "OBJECT",
        properties: {
            english: {
                type: "OBJECT",
                properties: {
                    mcqs: {
                        type: "ARRAY",
                        description: `List of exactly ${mcqCount} multiple choice questions in English`,
                        items: {
                            type: "OBJECT",
                            properties: {
                                question: { type: "STRING" },
                                options: { type: "ARRAY", items: { type: "STRING" } },
                                correctAnswer: { type: "STRING" }
                            },
                            required: ["question", "options", "correctAnswer"]
                        }
                    },
                    shortAnswers: {
                        type: "ARRAY",
                        description: `List of exactly ${shortTotal} short answer questions in English`,
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
                        description: `List of exactly ${mcqCount} multiple choice questions in Hindi (exact translation of English set)`,
                        items: {
                            type: "OBJECT",
                            properties: {
                                question: { type: "STRING" },
                                options: { type: "ARRAY", items: { type: "STRING" } },
                                correctAnswer: { type: "STRING" }
                            },
                            required: ["question", "options", "correctAnswer"]
                        }
                    },
                    shortAnswers: {
                        type: "ARRAY",
                        description: `List of exactly ${shortTotal} short answer questions in Hindi (exact translation of English set)`,
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

    const requestBody = {
        contents: [{ parts: parts }],
        generationConfig: {
            responseMimeType: "application/json",
            responseSchema: responseSchema
        }
    };

    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const response = await fetch(apiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        const errorMessage = errorData.error?.message || `HTTP ${response.status} Error`;
        throw new Error(`Gemini API Error: ${errorMessage}`);
    }

    const data = await response.json();
    try {
        const textContent = data.candidates[0].content.parts[0].text;
        const parsedJson = JSON.parse(textContent);
        if (!parsedJson.english || !parsedJson.hindi || !parsedJson.english.mcqs || !parsedJson.english.shortAnswers) {
            throw new Error("Invalid structure returned from Gemini API.");
        }
        return parsedJson;
    } catch (err) {
        console.error("Failed to parse Gemini response text as JSON:", err);
        throw new Error("Gemini successfully returned, but failed to format the response into the proper JSON structure. Please check model and parameters and try again.");
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

    const parts = [];
    mediaParts.forEach(media => {
        parts.push({
            inlineData: media.inlineData
        });
    });
    
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
        const card = createInteractiveQuestionCard(item, index, 'mcqs', 'english');
        enMcqContainer.appendChild(card);
    });

    // EN Short Qs
    const enShortContainer = document.getElementById('previewEnShorts');
    enShortContainer.innerHTML = '';
    english.shortAnswers.forEach((item, index) => {
        const card = createInteractiveQuestionCard(item, index, 'shortAnswers', 'english');
        enShortContainer.appendChild(card);
    });

    // HI MCQs
    const hiMcqContainer = document.getElementById('previewHiMcqs');
    hiMcqContainer.innerHTML = '';
    hindi.mcqs.forEach((item, index) => {
        const card = createInteractiveQuestionCard(item, index, 'mcqs', 'hindi');
        hiMcqContainer.appendChild(card);
    });

    // HI Short Qs
    const hiShortContainer = document.getElementById('previewHiShorts');
    hiShortContainer.innerHTML = '';
    hindi.shortAnswers.forEach((item, index) => {
        const card = createInteractiveQuestionCard(item, index, 'shortAnswers', 'hindi');
        hiShortContainer.appendChild(card);
    });

    lucide.createIcons();
}

function createInteractiveQuestionCard(item, index, type, language) {
    const card = document.createElement('div');
    card.className = 'preview-q-card';
    card.setAttribute('data-index', index);
    card.setAttribute('data-type', type);
    card.setAttribute('data-lang', language);

    const isMcq = type === 'mcqs';
    const isEn = language === 'english';

    // 1. Render normal display
    let displayHtml = '';
    if (isMcq) {
        displayHtml = `
            <div class="q-display">
                <div class="q-title">${index + 1}. ${item.question}</div>
                <ul class="q-options">
                    ${item.options.map(opt => `<li class="q-option">${opt}</li>`).join('')}
                </ul>
                <div class="q-answer">
                    <i data-lucide="check" class="text-success" style="width:14px;height:14px;"></i>
                    <span>${isEn ? 'Correct Answer' : 'सही उत्तर'}: ${item.correctAnswer}</span>
                </div>
            </div>
        `;
    } else {
        displayHtml = `
            <div class="q-display">
                <div class="q-title">${index + 1}. ${item.question}</div>
                <div class="q-answer-short"><strong>${isEn ? 'Ans' : 'उत्तर'}:</strong> ${item.answer}</div>
            </div>
        `;
    }

    // 2. Render actions toolbar
    const toolbarHtml = `
        <div class="q-actions-toolbar">
            <button type="button" class="btn-action btn-action-refine">
                <i data-lucide="sparkles"></i> <span>${isEn ? 'Refine with AI' : 'एआई परिष्करण'}</span>
            </button>
            <button type="button" class="btn-action btn-action-edit">
                <i data-lucide="pencil"></i> <span>${isEn ? 'Edit Manually' : 'मैन्युअल एडिट'}</span>
            </button>
        </div>
    `;

    // 3. Render AI refine drawer
    const refineDrawerHtml = `
        <div class="edit-drawer drawer-refine" style="display: none;">
            <label>${isEn ? 'How should AI refine this question?' : 'एआई को इस प्रश्न को कैसे परिष्कृत करना चाहिए?'}</label>
            <textarea class="refine-prompt-input" rows="2" placeholder="${isEn ? 'e.g., make it harder, change options, ask about photosynthesis...' : 'उदा., इसे और कठिन बनाएं, विकल्प बदलें, प्रकाश संश्लेषण के बारे में पूछें...' }"></textarea>
            <div class="drawer-actions">
                <button type="button" class="btn-drawer-sub btn-drawer-cancel">${isEn ? 'Cancel' : 'रद्द करें'}</button>
                <button type="button" class="btn-drawer-sub btn-drawer-sub-refine btn-drawer-apply">${isEn ? 'Regenerate' : 'पुनः उत्पन्न करें'}</button>
            </div>
        </div>
    `;

    // 4. Render manual edit drawer
    let editDrawerHtml = '';
    if (isMcq) {
        editDrawerHtml = `
            <div class="edit-drawer drawer-edit" style="display: none;">
                <div class="input-group">
                    <label>${isEn ? 'Question Text' : 'प्रश्न पाठ'}</label>
                    <input type="text" class="edit-q-input" value="${(item.question || '').replace(/"/g, '&quot;')}">
                </div>
                <div class="input-group">
                    <label>${isEn ? 'Options' : 'विकल्प'}</label>
                    <div class="edit-options-grid">
                        <input type="text" class="edit-opt-0" value="${(item.options[0] || '').replace(/"/g, '&quot;')}">
                        <input type="text" class="edit-opt-1" value="${(item.options[1] || '').replace(/"/g, '&quot;')}">
                        <input type="text" class="edit-opt-2" value="${(item.options[2] || '').replace(/"/g, '&quot;')}">
                        <input type="text" class="edit-opt-3" value="${(item.options[3] || '').replace(/"/g, '&quot;')}">
                    </div>
                </div>
                <div class="input-group">
                    <label>${isEn ? 'Correct Answer' : 'सही उत्तर'}</label>
                    <select class="edit-ans-select">
                        <option value="0" ${item.correctAnswer === item.options[0] ? 'selected' : ''}>A</option>
                        <option value="1" ${item.correctAnswer === item.options[1] ? 'selected' : ''}>B</option>
                        <option value="2" ${item.correctAnswer === item.options[2] ? 'selected' : ''}>C</option>
                        <option value="3" ${item.correctAnswer === item.options[3] ? 'selected' : ''}>D</option>
                    </select>
                </div>
                <div class="drawer-actions">
                    <button type="button" class="btn-drawer-sub btn-drawer-cancel">${isEn ? 'Cancel' : 'रद्द करें'}</button>
                    <button type="button" class="btn-drawer-sub btn-drawer-sub-save btn-drawer-save">${isEn ? 'Save' : 'सहेजें'}</button>
                </div>
            </div>
        `;
    } else {
        editDrawerHtml = `
            <div class="edit-drawer drawer-edit" style="display: none;">
                <div class="input-group">
                    <label>${isEn ? 'Question Text' : 'प्रश्न पाठ'}</label>
                    <input type="text" class="edit-q-input" value="${(item.question || '').replace(/"/g, '&quot;')}">
                </div>
                <div class="input-group">
                    <label>${isEn ? 'Answer Text' : 'उत्तर पाठ'}</label>
                    <textarea class="edit-a-input" rows="2">${item.answer || ''}</textarea>
                </div>
                <div class="drawer-actions">
                    <button type="button" class="btn-drawer-sub btn-drawer-cancel">${isEn ? 'Cancel' : 'रद्द करें'}</button>
                    <button type="button" class="btn-drawer-sub btn-drawer-sub-save btn-drawer-save">${isEn ? 'Save' : 'सहेजें'}</button>
                </div>
            </div>
        `;
    }

    card.innerHTML = displayHtml + toolbarHtml + refineDrawerHtml + editDrawerHtml;

    // 5. Setup Action Listeners
    const refineDrawer = card.querySelector('.drawer-refine');
    const editDrawer = card.querySelector('.drawer-edit');

    card.querySelector('.btn-action-refine').addEventListener('click', () => {
        const isShown = refineDrawer.style.display === 'flex';
        refineDrawer.style.display = isShown ? 'none' : 'flex';
        editDrawer.style.display = 'none';
    });

    card.querySelector('.btn-action-edit').addEventListener('click', () => {
        const isShown = editDrawer.style.display === 'flex';
        editDrawer.style.display = isShown ? 'none' : 'flex';
        refineDrawer.style.display = 'none';
    });

    // AI Refine Drawer Actions
    refineDrawer.querySelector('.btn-drawer-cancel').addEventListener('click', () => {
        refineDrawer.style.display = 'none';
        refineDrawer.querySelector('.refine-prompt-input').value = '';
    });

    refineDrawer.querySelector('.btn-drawer-sub-refine').addEventListener('click', () => {
        const promptVal = refineDrawer.querySelector('.refine-prompt-input').value.trim();
        if (!promptVal) {
            alert(isEn ? "Please enter a refinement prompt!" : "कृपया परिष्करण प्रॉम्प्ट दर्ज करें!");
            return;
        }
        handleAIRegeneration(index, type, language, promptVal);
    });

    // Manual Edit Drawer Actions
    editDrawer.querySelector('.btn-drawer-cancel').addEventListener('click', () => {
        editDrawer.style.display = 'none';
    });

    editDrawer.querySelector('.btn-drawer-sub-save').addEventListener('click', () => {
        let updatedData = {};
        if (isMcq) {
            const correctIdx = parseInt(editDrawer.querySelector('.edit-ans-select').value);
            const opts = [
                editDrawer.querySelector('.edit-opt-0').value.trim(),
                editDrawer.querySelector('.edit-opt-1').value.trim(),
                editDrawer.querySelector('.edit-opt-2').value.trim(),
                editDrawer.querySelector('.edit-opt-3').value.trim()
            ];
            updatedData = {
                question: editDrawer.querySelector('.edit-q-input').value.trim(),
                options: opts,
                correctAnswer: opts[correctIdx]
            };
        } else {
            updatedData = {
                question: editDrawer.querySelector('.edit-q-input').value.trim(),
                answer: editDrawer.querySelector('.edit-a-input').value.trim()
            };
        }
        handleManualSave(index, type, language, updatedData);
    });

    return card;
}

// Global Handlers for Interactive Editing
async function handleAIRegeneration(index, type, language, promptVal) {
    const apiKey = GEMINI_API_KEY.trim();
    if (!apiKey) {
        alert("Gemini API key is not configured! Please click the settings gear icon to configure it.");
        settingsModal.style.display = 'flex';
        return;
    }

    // 1. Show loader overlay on both English and Hindi cards for this question
    const cards = document.querySelectorAll(`.preview-q-card[data-index="${index}"][data-type="${type}"]`);
    cards.forEach(card => {
        const overlay = document.createElement('div');
        overlay.className = 'card-loader-overlay';
        overlay.innerHTML = `<div class="spinner"></div><span>AI Refinement...</span>`;
        card.appendChild(overlay);
    });

    try {
        const englishQuestion = generatedData.english[type][index];
        const hindiQuestion = generatedData.hindi[type][index];

        const result = await callGeminiAPIForSingleQuestion(apiKey, type, englishQuestion, hindiQuestion, promptVal);
        
        // 2. Update state
        generatedData.english[type][index] = result.english;
        generatedData.hindi[type][index] = result.hindi;

        // 3. Save to localStorage
        saveDataToLocalStorage();

        // 4. Re-render
        renderPreview();
    } catch (error) {
        alert("Failed to regenerate question: " + error.message);
        console.error(error);
        
        // Remove overlays in case of error
        cards.forEach(card => {
            const overlay = card.querySelector('.card-loader-overlay');
            if (overlay) overlay.remove();
        });
    }
}

async function callGeminiAPIForSingleQuestion(apiKey, type, englishQuestion, hindiQuestion, promptVal) {
    const isMcq = type === 'mcqs';
    let contextPrompt = '';
    
    if (isMcq) {
        contextPrompt = `
        Current English Question:
        Question: ${englishQuestion.question}
        Options:
        A) ${englishQuestion.options[0]}
        B) ${englishQuestion.options[1]}
        C) ${englishQuestion.options[2]}
        D) ${englishQuestion.options[3]}
        Correct Answer: ${englishQuestion.correctAnswer}

        Current Hindi Question:
        Question: ${hindiQuestion.question}
        Options:
        A) ${hindiQuestion.options[0]}
        B) ${hindiQuestion.options[1]}
        C) ${hindiQuestion.options[2]}
        D) ${hindiQuestion.options[3]}
        Correct Answer: ${hindiQuestion.correctAnswer}
        `;
    } else {
        contextPrompt = `
        Current English Question:
        Question: ${englishQuestion.question}
        Answer: ${englishQuestion.answer}

        Current Hindi Question:
        Question: ${hindiQuestion.question}
        Answer: ${hindiQuestion.answer}
        `;
    }

    const promptText = `
    You are an expert educational content writer. Your task is to modify a single question from a test bank based on the user's instructions.
    
    User Refinement Instruction: "${promptVal}"
    
    ${contextPrompt}
    
    Requirements:
    1. Regenerate BOTH the English and Hindi versions.
    2. Maintain 100% exact parity (translations) between the English and Hindi versions.
    3. Return the response using the exact JSON schema provided.
    `;

    let responseSchema;
    if (isMcq) {
        responseSchema = {
            type: "OBJECT",
            properties: {
                english: {
                    type: "OBJECT",
                    properties: {
                        question: { type: "STRING" },
                        options: { type: "ARRAY", items: { type: "STRING" } },
                        correctAnswer: { type: "STRING" }
                    },
                    required: ["question", "options", "correctAnswer"]
                },
                hindi: {
                    type: "OBJECT",
                    properties: {
                        question: { type: "STRING" },
                        options: { type: "ARRAY", items: { type: "STRING" } },
                        correctAnswer: { type: "STRING" }
                    },
                    required: ["question", "options", "correctAnswer"]
                }
            },
            required: ["english", "hindi"]
        };
    } else {
        responseSchema = {
            type: "OBJECT",
            properties: {
                english: {
                    type: "OBJECT",
                    properties: {
                        question: { type: "STRING" },
                        answer: { type: "STRING" }
                    },
                    required: ["question", "answer"]
                },
                hindi: {
                    type: "OBJECT",
                    properties: {
                        question: { type: "STRING" },
                        answer: { type: "STRING" }
                    },
                    required: ["question", "answer"]
                }
            },
            required: ["english", "hindi"]
        };
    }

    const requestBody = {
        contents: [{ parts: [{ text: promptText }] }],
        generationConfig: {
            responseMimeType: "application/json",
            responseSchema: responseSchema
        }
    };

    const model = modelSelect.value;
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const response = await fetch(apiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        const errorMessage = errorData.error?.message || `HTTP ${response.status} Error`;
        throw new Error(`Gemini API Error: ${errorMessage}`);
    }

    const data = await response.json();
    try {
        const textContent = data.candidates[0].content.parts[0].text;
        return JSON.parse(textContent);
    } catch (err) {
        throw new Error("Gemini successfully returned, but failed to format the response into the proper JSON structure.");
    }
}

function handleManualSave(index, type, language, updatedData) {
    // Update local state for specific language
    generatedData[language][type][index] = updatedData;
    
    // Save to localStorage
    saveDataToLocalStorage();

    // Re-render
    renderPreview();
}

function saveDataToLocalStorage() {
    if (generatedData) {
        localStorage.setItem('generated_paper_data', JSON.stringify(generatedData));
    }
}

function loadSavedData() {
    const saved = localStorage.getItem('generated_paper_data');
    if (saved) {
        try {
            generatedData = JSON.parse(saved);
            if (generatedData && generatedData.english && generatedData.hindi) {
                if (generatedData.examHeader) {
                    currentMode = 'fa';
                    // Toggle active classes on tab buttons
                    modeFaBtn.classList.add('active');
                    modeSingleBtn.classList.remove('active');
                    promptInputGroup.style.display = 'none';
                    faSettingsContainer.style.display = 'flex';
                    
                    // Update texts
                    document.querySelector('.card-upload .card-header h2').textContent = "Source Generated Documents (.docx / .pdf / .xlsx / .xls)";
                    document.querySelector('.dropzone-subtext').textContent = "Upload previously generated documents (.docx, .pdf, .xlsx, .xls) (Max 50MB per file)";
                    fileInput.setAttribute('accept', '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/pdf,.xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel');
                    generateBtn.querySelector('span').textContent = "Generate FA Exam Paper";
                    
                    // Update empty state text
                    emptyState.querySelector('p').textContent = "Enter your Gemini API key, upload 4-5 generated files (.docx, .pdf, .xlsx, .xls), fill the exam metadata, and click \"Generate FA Exam Paper\" to start.";
                    
                    // Populate inputs
                    const h = generatedData.examHeader;
                    if (faExamName) faExamName.value = h.examName || "FA - 1";
                    if (faSession) faSession.value = h.session || "2026-27";
                    if (faClass) faClass.value = h.class || "Class 6th";
                    if (faSubject) faSubject.value = h.subject || "Drawing";
                    if (faTime) faTime.value = h.time || "90 minutes";
                    if (faMaxMarks) faMaxMarks.value = h.maxMarks || "15";
                    if (faMcqCount) faMcqCount.value = h.mcqCount || 6;
                    if (faMcqMarks) faMcqMarks.value = h.mcqMarks || "1 * 6 = 6 Marks";
                    if (faShortTotal) faShortTotal.value = h.shortTotal || 4;
                    if (faShortLimit) faShortLimit.value = h.shortLimit || 2;
                    if (faShortMarks) faShortMarks.value = h.shortMarks || "2 * 2 = 4 Marks";
                    if (faLongHeader) faLongHeader.value = h.longHeader || "Section - C (Long Answer / Practical)";
                    if (faLongMarks) faLongMarks.value = h.longMarks || "(5 marks)";
                } else {
                    currentMode = 'single';
                    modeSingleBtn.classList.add('active');
                    modeFaBtn.classList.remove('active');
                    promptInputGroup.style.display = 'flex';
                    faSettingsContainer.style.display = 'none';
                    
                    // Update texts
                    document.querySelector('.card-upload .card-header h2').textContent = "Source Materials (Images / PDFs / Word)";
                    document.querySelector('.dropzone-subtext').textContent = "Upload study pages, images (JPG/PNG), PDFs, or Word docs (.docx) (Max 50MB per file)";
                    fileInput.setAttribute('accept', 'image/*,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document');
                    generateBtn.querySelector('span').textContent = "Generate Bilingual Documents";
                    
                    // Update empty state text
                    emptyState.querySelector('p').textContent = "Enter your Gemini API key, upload 3-4 images/PDFs/Word docs, and click \"Generate Bilingual Documents\" to start the process.";
                }

                // Show results card, hide emptyState
                resultsCard.style.display = 'block';
                emptyState.style.display = 'none';
                renderPreview();
            }
        } catch (e) {
            console.error("Failed to parse saved data from localStorage", e);
            localStorage.removeItem('generated_paper_data');
        }
    }
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
        
        // Dynamic live read of FA header inputs to merge into state at download time if it is an FA paper
        if (generatedData.examHeader) {
            generatedData.examHeader = {
                examName: faExamName.value.trim(),
                session: faSession.value.trim(),
                class: faClass.value.trim(),
                subject: faSubject.value.trim(),
                time: faTime.value.trim(),
                maxMarks: faMaxMarks.value.trim(),
                mcqCount: parseInt(faMcqCount.value) || 6,
                mcqMarks: faMcqMarks.value.trim(),
                shortTotal: parseInt(faShortTotal.value) || 4,
                shortLimit: parseInt(faShortLimit.value) || 2,
                shortMarks: faShortMarks.value.trim(),
                longHeader: faLongHeader.value.trim(),
                longMarks: faLongMarks.value.trim()
            };
            saveDataToLocalStorage();
        }

        const { Document, Paragraph, TextRun, PageBreak, HeadingLevel, AlignmentType } = docx;
        const children = [];

        if (generatedData.examHeader) {
            // ==========================================
            // FA EXAM PAPER FORMAT
            // ==========================================
            const h = generatedData.examHeader;
            
            // Header Text: "FA - 1    Session (2026-27).   Class 6th  subject Drawing"
            // Bold & Italic
            const headerText = `${h.examName}    Session (${h.session}).   ${h.class}  subject ${h.subject}`;
            const subheaderText = `Time:   ${h.time}.    | Max. Marks: ${h.maxMarks}`;
            
            children.push(
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 120 },
                    children: [
                        new TextRun({
                            text: headerText,
                            bold: true,
                            italics: true,
                            size: 26, // 13pt
                            font: "Calibri",
                            color: "000000"
                        })
                    ]
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 360 },
                    children: [
                        new TextRun({
                            text: subheaderText,
                            bold: true,
                            size: 24, // 12pt
                            font: "Calibri",
                            color: "000000"
                        })
                    ]
                })
            );
            
            // Section A: MCQ Header
            const mcqSectionHeader = isEnglish ? "Section - A    (MCQ)" : "विभाग - अ    (बहुविकल्पीय)";
            const mcqInstruction = isEnglish 
                ? `Q1 Choose the correct answer:    (${h.mcqMarks})` 
                : `प्रश्न 1 सही उत्तर चुनिए:    (${h.mcqMarks})`;
                
            children.push(
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 180, after: 100 },
                    children: [
                        new TextRun({
                            text: mcqSectionHeader,
                            bold: true,
                            size: 28, // 14pt
                            font: "Calibri",
                            color: "2B6CB0" // Blue color from template
                        })
                    ]
                }),
                new Paragraph({
                    spacing: { after: 120 },
                    children: [
                        new TextRun({
                            text: mcqInstruction,
                            bold: true,
                            size: 24,
                            font: "Calibri",
                            color: "000000"
                        })
                    ]
                })
            );
            
            // Render MCQs with letters (A), (B), (C)... and options 1), 2), 3), 4)...
            const questionLetters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"];
            dataset.mcqs.forEach((mcq, idx) => {
                const qLetter = questionLetters[idx] || (idx + 1).toString();
                
                children.push(
                    new Paragraph({
                        spacing: { before: 100, after: 60 },
                        children: [
                            new TextRun({
                                text: `(${qLetter}) ${mcq.question}`,
                                bold: true,
                                font: "Calibri",
                                size: 22
                            })
                        ]
                    })
                );
                
                // Render options in 2x2 grid to save space
                children.push(
                    new Paragraph({
                        indent: { left: 540 },
                        tabStops: [{ position: 4320 }], // 3 inches from margin
                        spacing: { after: 40 },
                        children: [
                            new TextRun({
                                text: `1)  ${mcq.options[0]}`,
                                font: "Calibri",
                                size: 22
                            }),
                            new TextRun({
                                text: `\t2)  ${mcq.options[1]}`,
                                font: "Calibri",
                                size: 22
                            })
                        ]
                    }),
                    new Paragraph({
                        indent: { left: 540 },
                        tabStops: [{ position: 4320 }],
                        spacing: { after: 40 },
                        children: [
                            new TextRun({
                                text: `3)  ${mcq.options[2]}`,
                                font: "Calibri",
                                size: 22
                            }),
                            new TextRun({
                                text: `\t4)  ${mcq.options[3]}`,
                                font: "Calibri",
                                size: 22
                            })
                        ]
                    })
                );
            });
            
            // Section B: Short Answer Header
            const shortSectionHeader = isEnglish ? "Section - B (Short Answer)" : "विभाग - ब (लघु उत्तरीय)";
            const shortInstruction = isEnglish 
                ? `Answer any ${h.shortLimit} questions:        (${h.shortMarks})` 
                : `किन्हीं ${h.shortLimit} प्रश्नों के उत्तर दीजिए:        (${h.shortMarks})`;
                
            children.push(
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 240, after: 100 },
                    children: [
                        new TextRun({
                            text: shortSectionHeader,
                            bold: true,
                            size: 28,
                            font: "Calibri",
                            color: "2B6CB0"
                        })
                    ]
                }),
                new Paragraph({
                    spacing: { after: 120 },
                    children: [
                        new TextRun({
                            text: shortInstruction,
                            bold: true,
                            size: 24,
                            font: "Calibri",
                            color: "000000"
                        })
                    ]
                })
            );
            
            // Render Short Answers (unnumbered, paragraph list)
            dataset.shortAnswers.forEach((sa) => {
                children.push(
                    new Paragraph({
                        spacing: { before: 100, after: 60 },
                        children: [
                            new TextRun({
                                text: sa.question,
                                font: "Calibri",
                                size: 22
                            })
                        ]
                    })
                );
            });
            
            // Section - C (Long Answer / Practical)
            const longSectionHeader = h.longHeader || "Section - C (Long Answer / Practical)";
            const longMarksText = h.longMarks || "(5 marks)";
            
            children.push(
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 240, after: 100 },
                    children: [
                        new TextRun({
                            text: `${longSectionHeader}        ${longMarksText}`,
                            bold: true,
                            size: 28,
                            font: "Calibri",
                            color: "2B6CB0"
                        })
                    ]
                }),
                // Add empty paragraphs to serve as manual typing space for Section C
                new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "" })] }),
                new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "" })] }),
                new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "" })] })
            );
            
            // ANSWER KEY ON SEPARATE PAGE
            children.push(new PageBreak());
            
            const keyHeader = isEnglish ? "ANSWER KEY (FOR TEACHERS ONLY)" : "उत्तर कुंजी (केवल शिक्षकों के लिए)";
            children.push(
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 120, after: 240 },
                    children: [
                        new TextRun({
                            text: keyHeader,
                            bold: true,
                            size: 32,
                            font: "Calibri",
                            color: "2F855A"
                        })
                    ]
                }),
                new Paragraph({
                    spacing: { before: 120, after: 80 },
                    children: [
                        new TextRun({
                            text: isEnglish ? "Section A MCQs Correct Answers:" : "विभाग अ बहुविकल्पीय सही उत्तर:",
                            bold: true,
                            size: 24,
                            font: "Calibri"
                        })
                    ]
                })
            );
            
            dataset.mcqs.forEach((mcq, idx) => {
                const qLetter = questionLetters[idx] || (idx + 1).toString();
                children.push(
                    new Paragraph({
                        indent: { left: 360 },
                        spacing: { after: 40 },
                        children: [
                            new TextRun({
                                text: `(${qLetter})  Correct Answer: `,
                                bold: true,
                                size: 22,
                                font: "Calibri"
                            }),
                            new TextRun({
                                text: mcq.correctAnswer,
                                size: 22,
                                font: "Calibri",
                                color: "2F855A"
                            })
                        ]
                    })
                );
            });
            
            children.push(
                new Paragraph({
                    spacing: { before: 240, after: 80 },
                    children: [
                        new TextRun({
                            text: isEnglish ? "Section B Short Answers Model Answers:" : "विभाग ब लघु उत्तरीय मॉडल उत्तर:",
                            bold: true,
                            size: 24,
                            font: "Calibri"
                        })
                    ]
                })
            );
            
            dataset.shortAnswers.forEach((sa) => {
                children.push(
                    new Paragraph({
                        spacing: { before: 120, after: 40 },
                        children: [
                            new TextRun({
                                text: `Q. ${sa.question}`,
                                bold: true,
                                size: 22,
                                font: "Calibri"
                            })
                        ]
                    }),
                    new Paragraph({
                        indent: { left: 360 },
                        spacing: { after: 120 },
                        children: [
                            new TextRun({
                                text: isEnglish ? "Model Answer: " : "मॉडल उत्तर: ",
                                bold: true,
                                size: 22,
                                font: "Calibri",
                                color: "4A5568"
                            }),
                            new TextRun({
                                text: sa.answer,
                                size: 22,
                                font: "Calibri",
                                color: "2D3748"
                            })
                        ]
                    })
                );
            });

            // Section C: Long Answer / Practical Answer Key Placeholder
            children.push(
                new Paragraph({
                    spacing: { before: 240, after: 80 },
                    children: [
                        new TextRun({
                            text: isEnglish ? "Section C Long Answer / Practical Model Answers:" : "विभाग स दीर्घ उत्तरीय / व्यावहारिक मॉडल उत्तर:",
                            bold: true,
                            size: 24,
                            font: "Calibri"
                        })
                    ]
                }),
                new Paragraph({
                    indent: { left: 360 },
                    spacing: { after: 120 },
                    children: [
                        new TextRun({
                            text: isEnglish 
                                ? "[Teacher Note: Write or grade this section based on the custom question details you manually type after downloading.]" 
                                : "[शिक्षक टिप्पणी: इस अनुभाग को डाउनलोड करने के बाद मैन्युअल रूप से लिखे जाने वाले कस्टम प्रश्न के आधार पर भरें या मूल्यांकित करें।]",
                            italics: true,
                            size: 22,
                            font: "Calibri",
                            color: "718096"
                        })
                    ]
                })
            );

        } else {
            // ==========================================
            // INDIVIDUAL STANDARD GENERATOR FORMATTING
            // ==========================================
            const titleText = isEnglish ? "ASSESSMENT QUESTION BANK" : "मूल्यांकन प्रश्न बैंक";
            const mcqHeadingText = isEnglish ? "Section A: Multiple Choice Questions (MCQs)" : "भाग अ: बहुविकल्पीय प्रश्न (MCQs)";
            const shortHeadingText = isEnglish ? "Section B: Short Answer Questions" : "भाग ब: लघु उत्तरीय प्रश्न";
            const answerLabel = isEnglish ? "Correct Answer" : "सही उत्तर";

            children.push(
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 360 },
                    children: [
                        new TextRun({
                            text: titleText,
                            bold: true,
                            size: 36, // 18pt
                            font: "Calibri",
                            color: "1A365D"
                        })
                    ]
                })
            );

            // MCQs Page Heading
            children.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 240, after: 180 },
                    children: [
                        new TextRun({
                            text: mcqHeadingText,
                            bold: true,
                            size: 28,
                            font: "Calibri",
                            color: "2B6CB0"
                        })
                    ]
                })
            );

            dataset.mcqs.forEach((mcq, idx) => {
                children.push(
                    new Paragraph({
                        spacing: { before: 120, after: 80 },
                        children: [
                            new TextRun({
                                text: `${idx + 1}. `,
                                bold: true,
                                font: "Calibri",
                                size: 24
                            }),
                            new TextRun({
                                text: mcq.question,
                                font: "Calibri",
                                size: 24
                            })
                        ]
                    })
                );

                const letters = ['A', 'B', 'C', 'D'];
                mcq.options.forEach((opt, optIdx) => {
                    children.push(
                        new Paragraph({
                            indent: { left: 720 },
                            spacing: { after: 40 },
                            children: [
                                new TextRun({
                                    text: `(${letters[optIdx]}) `,
                                    bold: true,
                                    font: "Calibri",
                                    size: 22
                                }),
                                new TextRun({
                                    text: opt,
                                    font: "Calibri",
                                    size: 22
                                })
                            ]
                        })
                    );
                });

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
                                color: "2F855A"
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

            // Short Questions Page Break & Heading
            children.push(new PageBreak());
            children.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 240, after: 180 },
                    children: [
                        new TextRun({
                            text: shortHeadingText,
                            bold: true,
                            size: 28,
                            font: "Calibri",
                            color: "2B6CB0"
                        })
                    ]
                })
            );

            dataset.shortAnswers.forEach((sa, idx) => {
                children.push(
                    new Paragraph({
                        spacing: { before: 120, after: 80 },
                        children: [
                            new TextRun({
                                text: `Q${idx + 1}. `,
                                bold: true,
                                font: "Calibri",
                                size: 24
                            }),
                            new TextRun({
                                text: sa.question,
                                font: "Calibri",
                                size: 24
                            })
                        ]
                    })
                );

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
        }

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
        
        let filename = isEnglish ? "assessment_english.docx" : "assessment_hindi.docx";
        if (generatedData.examHeader) {
            const h = generatedData.examHeader;
            const prefix = `${h.examName}_${h.class}_${h.subject}`.toLowerCase().replace(/[^a-z0-9]/g, '_');
            filename = isEnglish ? `${prefix}_english.docx` : `${prefix}_hindi.docx`;
        }
        
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        
        document.body.removeChild(a);
        
        // Delay revoking the object URL to give the browser time to complete the download
        setTimeout(() => {
            URL.revokeObjectURL(url);
        }, 10000);

        addLog(`Successfully generated and downloaded ${language} document.`, 'success');

    } catch (err) {
        console.error("DOCX generation error:", err);
        alert(`Failed to generate DOCX file: ${err.message}`);
    }
}

// Initialize the application after all variables and functions are declared
if (document.readyState === "loading") {
    window.addEventListener('DOMContentLoaded', initApp);
} else {
    initApp();
}

function initApp() {
    console.log("docx import:", docx);
    lucide.createIcons();
    setupEventListeners();
    loadSavedData();
}
