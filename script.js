let excelData = null;
let images = new Map();
let defaultMaleImage = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAAIAAgDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAb/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAX/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwCdABmX/9k=';
let defaultFemaleImage = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAAIAAgDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAb/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAX/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwCdABmX/9k=';
const LOGO_BASE64 = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAAIAAgDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAb/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAX/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwCdABmX/9k=';
let uploadedLogoBase64 = '';
let deletedCards = new Set(); // Track deleted cards by index
let manuallyAddedCards = []; // Track manually added cards

document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('excelFile').addEventListener('change', handleExcelUpload);
    document.getElementById('imageFiles').addEventListener('change', handleImageUpload);
    document.getElementById('generateBtn').addEventListener('click', generatePDF);
    document.getElementById('logoFile').addEventListener('change', handleLogoUpload);
    document.getElementById('fieldsForm').addEventListener('change', updateSelectedFields);
    document.getElementById('cardWidth').addEventListener('input', renderPreview);
    document.getElementById('cardHeight').addEventListener('input', renderPreview);
    document.getElementById('addCardBtn').addEventListener('click', addCardManually);
    document.getElementById('restoreBtn').addEventListener('click', restoreAllCards);
    document.getElementById('exportExcelBtn').addEventListener('click', exportManualDataToExcel);
    document.getElementById('exportImagesBtn').addEventListener('click', exportImagesAsZip);
    document.getElementById('generateQrCodesBtn').addEventListener('click', generateQrCodeImages);

    const fieldsForm = document.getElementById('fieldsForm');
    fieldsForm.addEventListener('dragstart', handleDragStart);
    fieldsForm.addEventListener('dragover', handleDragOver);
    fieldsForm.addEventListener('drop', handleDrop);
    fieldsForm.addEventListener('dragend', handleDragEnd);

    document.getElementById('logoAlign').addEventListener('input', renderPreview);
    document.getElementById('logoSize').addEventListener('input', renderPreview);
    document.getElementById('headerField').addEventListener('change', renderPreview);
    document.getElementById('mainTitleField').addEventListener('change', renderPreview);

    // QR Code Settings
    document.getElementById('includeQRCode').addEventListener('change', renderPreview);
    document.getElementById('qrCodeSize').addEventListener('input', renderPreview);
    document.getElementById('qrCodePosition').addEventListener('change', renderPreview);
    
    // Email Button
    document.getElementById('emailCardsBtn').addEventListener('click', openEmailModal);
    
    // Email Modal Listeners
    document.getElementById('saveEmailConfigBtn').addEventListener('click', sendAllEmails);
    document.getElementById('cancelEmailConfigBtn').addEventListener('click', () => {
        document.getElementById('emailConfigModal').classList.remove('visible');
    });
});

let draggedItem = null;

function handleDragStart(e) {
    if (e.target.classList.contains('draggable-field')) {
        draggedItem = e.target;
        setTimeout(() => {
            if(draggedItem) draggedItem.classList.add('dragging');
        }, 0);
    }
}

function handleDragOver(e) {
    e.preventDefault();
    const form = document.getElementById('fieldsForm');
    if (!form || !draggedItem) return;

    const afterElement = getDragAfterElement(form, e.clientY);
    if (afterElement == null) {
        form.appendChild(draggedItem);
    } else {
        form.insertBefore(draggedItem, afterElement);
    }
}

function getDragAfterElement(container, y) {
    const draggableElements = [...container.querySelectorAll('.draggable-field:not(.dragging)')];

    return draggableElements.reduce((closest, child) => {
        const box = child.getBoundingClientRect();
        const offset = y - box.top - box.height / 2;
        if (offset < 0 && offset > closest.offset) {
            return { offset: offset, element: child };
        } else {
            return closest;
        }
    }, { offset: Number.NEGATIVE_INFINITY }).element;
}

function handleDrop(e) {
    e.preventDefault();
    if (draggedItem) {
        draggedItem.classList.remove('dragging');
        updateSelectedFields();
        draggedItem = null;
    }
}

function handleDragEnd() {
    if (draggedItem) {
        draggedItem.classList.remove('dragging');
        draggedItem = null;
    }
}

function handleExcelUpload(event) {
    const file = event.target.files[0];
    if (file) {
        document.getElementById('excelInfo').textContent = file.name;
        // Clear deleted cards and manually added cards when new data is loaded
        deletedCards.clear();
        manuallyAddedCards = [];
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                // Read as 2D array
                const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                // Find the header row (row with 'NOM & PRENOM')
                let headerRowIndex = rows.findIndex(row => row.includes('NOM & PRENOM'));
                if (headerRowIndex === -1) {
                    alert('Could not find header row (NOM & PRENOM) in the Excel file.');
                    return;
                }
                const headers = rows[headerRowIndex];
                // Parse data rows
                excelData = rows.slice(headerRowIndex + 1).map(row => {
                        const obj = {};
                        headers.forEach((header, i) => {
                        obj[header] = row[i] || '';
                        });
                        return obj;
                }).filter(row => Object.values(row).some(val => val !== ''));
                // Populate field selection UI
                const fieldSection = document.getElementById('fieldSelection');
                const fieldsForm = document.getElementById('fieldsForm');
                const imageNameField = document.getElementById('imageNameField');
                const headerFieldSelect = document.getElementById('headerField');
                const mainTitleFieldSelect = document.getElementById('mainTitleField');

                fieldsForm.innerHTML = '';
                imageNameField.innerHTML = '';
                headerFieldSelect.innerHTML = '<option value="">None</option>';
                mainTitleFieldSelect.innerHTML = '<option value="">None</option>';

                headers.forEach(header => {
                    const label = document.createElement('label');
                    label.className = 'draggable-field';
                    label.draggable = true;
                    
                    const checkbox = document.createElement('input');
                    checkbox.type = 'checkbox';
                    checkbox.name = 'field';
                    checkbox.value = header;
                    checkbox.checked = true;
                    label.appendChild(checkbox);
                    label.appendChild(document.createTextNode(' ' + header));
                    fieldsForm.appendChild(label);

                    const option = document.createElement('option');
                    option.value = header;
                    option.textContent = header;
                    imageNameField.appendChild(option.cloneNode(true));
                    headerFieldSelect.appendChild(option.cloneNode(true));
                    mainTitleFieldSelect.appendChild(option.cloneNode(true));
                });

                // Set default selections for layout fields
                if(headers.includes('ADHERENT')) headerFieldSelect.value = 'ADHERENT';
                if(headers.includes('NOM & PRENOM')) mainTitleFieldSelect.value = 'NOM & PRENOM';

                fieldSection.style.display = headers.length ? 'block' : 'none';
                updateSelectedFields();
                updateGenerateButton();
            } catch (error) {
                console.error('Error reading Excel file:', error);
                alert('Error reading Excel file. Please make sure it is a valid Excel file.');
            }
        };
        reader.onerror = function() {
            alert('Error reading the file. Please try again.');
        };
        reader.readAsArrayBuffer(file);
    }
    
    updateCardCount();
}

function handleImageUpload(event) {
    const files = event.target.files;
    images.clear();
    if (files.length > 0) {
        document.getElementById('imageInfo').textContent = `Loaded ${files.length} images`;
        Array.from(files).forEach((file) => {
            const reader = new FileReader();
            reader.onload = function(e) {
                const id = file.name.replace(/\.[^/.]+$/, "").toLowerCase();
                if (id === 'f') {
                    defaultFemaleImage = e.target.result;
                } else if (id === 'm') {
                    defaultMaleImage = e.target.result;
                } else {
                    images.set(id, e.target.result);
                }
                updateGenerateButton();
                if (typeof renderUploadedImagesList === 'function') renderUploadedImagesList();
            };
            reader.onerror = function() {
                console.error('Error loading image:', file.name);
            };
            reader.readAsDataURL(file);
        });
    } else {
        document.getElementById('imageInfo').textContent = '';
        updateGenerateButton();
        if (typeof renderUploadedImagesList === 'function') renderUploadedImagesList();
    }
}

function handleLogoUpload(event) {
    const file = event.target.files[0];
    if (file) {
        document.getElementById('logoInfo').textContent = file.name;
        const reader = new FileReader();
        reader.onload = function(e) {
            uploadedLogoBase64 = e.target.result;
        };
        reader.readAsDataURL(file);
    }
}

function updateGenerateButton() {
    const generateBtn = document.getElementById('generateBtn');
    const hasExcelData = excelData && excelData.length > 0;
    const hasImages = images.size > 0;
    
    console.log('Update button state:', { hasExcelData, hasImages });
    generateBtn.disabled = !(hasExcelData && hasImages);
    
    if (!generateBtn.disabled) {
        generateBtn.style.backgroundColor = '#4CAF50';
    } else {
        generateBtn.style.backgroundColor = '#cccccc';
    }
}

function withLeadingZero(num) {
    if (!num) return '';
    num = String(num).trim();
    return num.startsWith('0') ? num : '0' + num;
}

let selectedFields = [];
function updateSelectedFields() {
    const checked = Array.from(document.querySelectorAll('#fieldsForm input[type="checkbox"]:checked'));
    selectedFields = checked.map(cb => cb.value);
    renderPreview();
}

function createIDCard(student, index, showDeleteButton = true, overridePhotoSrc = null) {
    const cardWidth = document.getElementById('cardWidth').value;
    const cardHeight = document.getElementById('cardHeight').value;
    const baseUnit = cardWidth / 325; // Scale everything relative to the original width

    // Layout settings
    const logoAlign = document.getElementById('logoAlign').value;
    const logoSize = document.getElementById('logoSize').value;
    const headerField = document.getElementById('headerField').value;
    const mainTitleField = document.getElementById('mainTitleField').value;
    
    const card = document.createElement('div');
    card.className = 'id-card-container';
    card.style.width = `${cardWidth}px`;
    card.style.height = `${cardHeight}px`;

    const cardInner = document.createElement('div');
    cardInner.className = 'id-card';
    card.appendChild(cardInner);

    const logoSrc = uploadedLogoBase64 || LOGO_BASE64;

    const headerText = student[headerField] || `ID: ${String(index + 1).padStart(3, '0')}`;
    const mainTitleText = student[mainTitleField] || ' ';

    cardInner.innerHTML = `
        <div class="card-header" style="justify-content: ${logoAlign === 'center' ? 'center' : (logoAlign === 'right' ? 'flex-end' : 'flex-start')};">
            <img src="${logoSrc}" alt="Logo" class="logo" style="height:${logoSize}px;">
            <div class="header-text">${headerText}</div>
        </div>
        <div class="card-body">
            <div class="index-number-rotated">${String(index + 1).padStart(3, '0')}</div>
        </div>
    `;

    const cardBody = cardInner.querySelector('.card-body');

    // Photo
    const id = String(student['ADHERENT'] || '').trim().toLowerCase();
    const type = String(student['type'] || student['TYPE'] || '').trim().toLowerCase();
    let photoSrc = overridePhotoSrc || images.get(id) || ((type === 'f') ? defaultFemaleImage : defaultMaleImage);

    if (photoSrc) {
        const photo = document.createElement('img');
        photo.src = photoSrc;
        photo.className = 'photo';
        photo.style.height = `${baseUnit * 130}px`; // Scaled height
        cardBody.appendChild(photo);
    }
    
    const textContainer = document.createElement('div');
    textContainer.className = 'text-container';
    
    const mainTitleDiv = document.createElement('div');
    mainTitleDiv.className = 'student-name';
    mainTitleDiv.textContent = mainTitleText;
    mainTitleDiv.style.fontSize = `${baseUnit * 22}px`;
    textContainer.appendChild(mainTitleDiv);

    const fieldsContainer = document.createElement('div');
    fieldsContainer.className = 'fields-container';
    
    // Use the order from the draggable list
    selectedFields.forEach(field => {
        // Exclude header and main title if they are selected
        if (student[field] && field !== headerField && field !== mainTitleField) {
            const fieldDiv = document.createElement('div');
            fieldDiv.className = 'field';
            fieldDiv.innerHTML = `<span class="field-name">${field}:</span> ${student[field]}`;
            fieldDiv.style.fontSize = `${baseUnit * 12}px`;
            fieldsContainer.appendChild(fieldDiv);
        }
    });

    textContainer.appendChild(fieldsContainer);
    cardBody.appendChild(textContainer);
    
    // QR Code Generation
    if (document.getElementById('includeQRCode').checked) {
        const qrCodeSize = document.getElementById('qrCodeSize').value;
        const qrCodePosition = document.getElementById('qrCodePosition').value;

        // Create a simpler URL with just the essential data
        const simpleData = {
            id: String(index + 1).padStart(3, '0'),
            name: mainTitleText.substring(0, 50), // Limit name length
            adherent: headerText.substring(0, 20) // Limit header length
        };

        // Store full student data in localStorage for the digital card to access
        const fullStudentData = {
            header: headerText,
            title: mainTitleText,
            photo: photoSrc,
            fields: selectedFields.reduce((obj, field) => {
                if (field !== headerField && field !== mainTitleField) {
                    obj[field] = student[field] || '';
                }
                return obj;
            }, {})
        };
        
        localStorage.setItem(`student_${index}`, JSON.stringify(fullStudentData));

        const jsonString = JSON.stringify(simpleData);
        const encodedData = encodeURIComponent(jsonString);
        
        // Use a simpler URL structure that works locally
        // TODO: Replace with your hosting URL when deployed
        // Example: const digitalCardUrl = `https://yourusername.github.io/id-card-generator/digital_card.html?id=${index}&data=${encodedData}`;
        const digitalCardUrl = `http://localhost:8000/digital_card.html?id=${index}&data=${encodedData}`;

        const qrCodeContainer = document.createElement('div');
        qrCodeContainer.className = 'qr-code-container';
        qrCodeContainer.style.width = `${qrCodeSize}px`;
        qrCodeContainer.style.height = `${qrCodeSize}px`;
        
        // Position the QR code
        const [y, x] = qrCodePosition.split('-');
        qrCodeContainer.style.position = 'absolute';
        qrCodeContainer.style[y] = '5px';
        qrCodeContainer.style[x] = '5px';
        
        cardInner.appendChild(qrCodeContainer);

        // Generate QR code with error handling
        try {
            new QRCode(qrCodeContainer, {
                text: digitalCardUrl,
                width: parseInt(qrCodeSize),
                height: parseInt(qrCodeSize),
                correctLevel: QRCode.CorrectLevel.L // Use lower correction level for shorter URLs
            });
        } catch (error) {
            console.error('Error generating QR code:', error);
            // Fallback: create a simple text element instead
            qrCodeContainer.innerHTML = `<div style="text-align: center; font-size: 8px; padding: 2px;">QR Code Error</div>`;
        }
    }

    if (showDeleteButton) {
        const deleteButton = document.createElement('button');
        deleteButton.className = 'delete-card-btn';
        deleteButton.innerHTML = '&times;';
        deleteButton.onclick = () => deleteCard(index);
        card.appendChild(deleteButton);

        card.addEventListener('dblclick', () => deleteCard(index));
    }

    return card;
}

async function generatePDF() {
    if (!excelData || excelData.length === 0) {
        alert('No Excel data loaded. Please upload an Excel file.');
        return;
    }
    if (images.size === 0) {
        alert('No images loaded. Please upload student photos first.');
        return;
    }

    let fileName = prompt("Enter a name for the PDF file:", "id-cards");
    if (fileName === null || fileName.trim() === "") {
        console.log("PDF generation cancelled by user.");
        return;
    }
    // Sanitize filename
    fileName = fileName.replace(/[^a-z0-9_-\s]/gi, '').trim();
    if (fileName === "") {
        fileName = "id-cards";
    }
    fileName += ".pdf";

    console.log('Starting PDF generation...');
    const modalOverlay = document.createElement('div');
    modalOverlay.id = 'pdf-modal-overlay';
    modalOverlay.style.position = 'fixed';
    modalOverlay.style.top = '0';
    modalOverlay.style.left = '0';
    modalOverlay.style.width = '100%';
    modalOverlay.style.height = '100%';
    modalOverlay.style.backgroundColor = 'rgba(0, 0, 0, 0.5)';
    modalOverlay.style.display = 'flex';
    modalOverlay.style.justifyContent = 'center';
    modalOverlay.style.alignItems = 'center';
    modalOverlay.style.zIndex = '9999';

    const modalContent = document.createElement('div');
    modalContent.id = 'pdf-modal';
    modalContent.style.backgroundColor = 'white';
    modalContent.style.padding = '2rem';
    modalContent.style.borderRadius = '8px';
    modalContent.style.boxShadow = '0 2px 10px rgba(0, 0, 0, 0.1)';
    modalContent.style.textAlign = 'center';
    modalContent.style.maxWidth = '90%';
    modalContent.style.width = 'auto';

    const loadingText = document.createElement('div');
    loadingText.textContent = 'Generating PDF, please wait...';
    loadingText.style.fontSize = '1.2rem';
    loadingText.style.color = '#1a73e8';
    loadingText.style.marginBottom = '1rem';

    const spinner = document.createElement('div');
    spinner.style.border = '4px solid #f3f3f3';
    spinner.style.borderTop = '4px solid #1a73e8';
    spinner.style.borderRadius = '50%';
    spinner.style.width = '40px';
    spinner.style.height = '40px';
    spinner.style.animation = 'spin 1s linear infinite';
    spinner.style.margin = '0 auto';

    const style = document.createElement('style');
    style.textContent = `@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }`;
    document.head.appendChild(style);

    modalContent.appendChild(loadingText);
    modalContent.appendChild(spinner);
    modalOverlay.appendChild(modalContent);
    document.body.appendChild(modalOverlay);

    const filteredData = excelData.filter((student, index) => !deletedCards.has(index) && (student['NOM & PRENOM'] || student['TEL ADH'] || student['TEL PARENT'] || student['LYCEE'] || student['ADHERENT']));
    
    const allCardsData = [...filteredData, ...manuallyAddedCards];
    document.getElementById('cardCount').textContent = `Total cards generated: ${allCardsData.length}`;

    try {
        // Check if jsPDF is available
        if (typeof window.jspdf === 'undefined') {
            throw new Error('jsPDF library not loaded. Please refresh the page and try again.');
        }
        
        const { jsPDF } = window.jspdf;
        if (!jsPDF) {
            throw new Error('jsPDF constructor not found. Please refresh the page and try again.');
        }
        
        const doc = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4' });
        console.log('PDF document created successfully');

        const cardsPerPage = 8;
        const pageW = 210;
        const pageH = 297;
        const cols = 2;
        const rows = 4;
        
        // Use minimal gaps to maximize card size
        const hGap = 2; // mm
        const vGap = 2; // mm

        // Calculate card size to fill the page with the given gaps
        const cardWidthMM = (pageW - ((cols - 1) * hGap)) / cols;
        const cardHeightMM = (pageH - ((rows - 1) * vGap)) / rows;

        console.log(`Processing ${allCardsData.length} cards...`);

        for (let i = 0; i < allCardsData.length; i++) {
            const student = allCardsData[i];
            console.log(`Processing card ${i + 1}/${allCardsData.length} for student:`, student['NOM & PRENOM'] || student['ADHERENT'] || `Student ${i + 1}`);
            
            if (i > 0 && i % cardsPerPage === 0) {
                doc.addPage();
                console.log(`Added new page for card ${i + 1}`);
            }

            const pdfContainer = document.createElement('div');
            pdfContainer.style.position = 'absolute';
            pdfContainer.style.left = '-9999px';
            pdfContainer.style.top = '-9999px';
            pdfContainer.style.width = '325px';
            pdfContainer.style.height = '200px';
            document.body.appendChild(pdfContainer);

            // Create card without delete buttons for PDF
            const cardElement = createIDCard(student, i, false);
            pdfContainer.appendChild(cardElement);

            const indexOnPage = i % cardsPerPage;
            const col = indexOnPage % 2;
            const row = Math.floor(indexOnPage / 2);
            
            // Position cards in a simple grid, starting from top-left
            const x = col * (cardWidthMM + hGap);
            const y = row * (cardHeightMM + vGap);

            try {
                console.log(`Generating canvas for card ${i + 1}...`);
                const canvas = await html2canvas(cardElement, { 
                    scale: 2, 
                    useCORS: true, 
                    backgroundColor: '#ffffff',
                    width: 325,
                    height: 200,
                    allowTaint: true
                });
                console.log(`Canvas generated for card ${i + 1}`);
                
                const imgData = canvas.toDataURL('image/jpeg', 0.8);
                // Add the image to the PDF, forcing it into the calculated grid space
                doc.addImage(imgData, 'JPEG', x, y, cardWidthMM, cardHeightMM);
                console.log(`Image added to PDF for card ${i + 1}`);
            } catch (canvasError) {
                console.error('Error generating canvas for card', i, canvasError);
                // Continue with next card instead of failing completely
            }
            
            document.body.removeChild(pdfContainer);

            loadingText.textContent = `Generating PDF... ${i + 1}/${allCardsData.length} cards processed`;
        }

        console.log('Saving PDF...');
        doc.save(fileName);
        modalOverlay.remove();
        document.head.removeChild(style);
        console.log('PDF saved successfully');

    } catch (error) {
        console.error('Error generating PDF:', error);
        alert(`An error occurred while generating the PDF: ${error.message}\n\nPlease check console for details.`);
        if (document.getElementById('pdf-modal-overlay')) {
            document.body.removeChild(document.getElementById('pdf-modal-overlay'));
        }
        if (style.parentElement) {
             document.head.removeChild(style);
        }
    } finally {
        showExportModal();
    }
}

function renderPreview() {
    const preview = document.getElementById('preview');
    const manualCardsSection = document.getElementById('manualCardsSection');
    
    preview.innerHTML = '';
    
    const allCards = [
        ...((excelData || []).map((student, index) => ({ ...student, originalIndex: index, isManual: false })).filter(card => !deletedCards.has(card.originalIndex))),
        ...(manuallyAddedCards.map((student, index) => ({ ...student, originalIndex: index, isManual: true })))
    ];

    if (allCards.length === 0) {
        preview.innerHTML = `
            <div style="text-align: center; padding: 3rem; color: #666;">
                <div style="font-size: 4rem; margin-bottom: 1rem;">📋</div>
                <h3>No cards to display</h3>
                <p>Upload an Excel file or add cards manually to see them here.</p>
                <div style="margin-top: 2rem; padding: 1rem; background: #f8fafc; border-radius: 8px; max-width: 500px; margin-left: auto; margin-right: auto;">
                    <h4 style="color: #1e293b; margin-bottom: 1rem;">How to get started:</h4>
                    <ul style="text-align: left; color: #64748b;">
                        <li>📊 Upload an Excel file with student data</li>
                        <li>📸 Upload student photos (optional)</li>
                        <li>🏷️ Select which fields to display on cards</li>
                        <li>➕ Add cards manually if needed</li>
                        <li>👀 Preview your cards here</li>
                    </ul>
                </div>
            </div>
        `;
        return;
    }
    
    // Helper function to display cards, defined once at the top of renderPreview
    function displayCards(cards, page = 1) {
        const cardsPerPage = 8;
        const startIndex = (page - 1) * cardsPerPage;
        const endIndex = startIndex + cardsPerPage;
        const pageCards = cards.slice(startIndex, endIndex);
        
        const cardsGrid = document.getElementById('cardsGrid');
        if (!cardsGrid) return; // Exit if the grid isn't on the page yet
        cardsGrid.innerHTML = '';
        
        pageCards.forEach((cardData, index) => {
            const student = cardData;
            const { isManual, originalIndex } = cardData;

            const globalIndex = startIndex + index;
            
            const card = createIDCard(student, globalIndex, true);
            
            const cardTypeIndicator = document.createElement('div');
            cardTypeIndicator.style.cssText = `
                position: absolute;
                top: 10px;
                left: 10px;
                padding: 0.25rem 0.5rem;
                border-radius: 4px;
                font-size: 0.75rem;
                font-weight: 600;
                z-index: 5;
                ${isManual ? 'background: #34d399; color: white;' : 'background: #3b82f6; color: white;'}
            `;
            cardTypeIndicator.textContent = isManual ? 'MANUAL' : 'EXCEL';
            card.appendChild(cardTypeIndicator);
            
            card.style.cursor = 'pointer';
            card.style.transition = 'transform 0.2s ease, box-shadow 0.2s ease';
            
            card.addEventListener('mouseenter', () => {
                card.style.transform = 'scale(1.02)';
                card.style.boxShadow = '0 8px 25px rgba(0, 0, 0, 0.15)';
            });
            
            card.addEventListener('mouseleave', () => {
                card.style.transform = 'scale(1)';
                card.style.boxShadow = '0 4px 6px -1px rgba(0, 0, 0, 0.1)';
            });
            
            card.addEventListener('click', (e) => {
                if (e.target.classList.contains('card-delete-btn') || e.target.parentElement.classList.contains('card-delete-btn')) return;
                
                card.style.transform = 'scale(0.98)';
                setTimeout(() => {
                    card.style.transform = 'scale(1.02)';
                    setTimeout(() => {
                        card.style.transform = 'scale(1)';
                    }, 100);
                }, 100);
                
                showCardDetails(student, globalIndex + 1, isManual);
            });
            
            const deleteFn = (e) => {
                e.stopPropagation();
                if (isManual) {
                    deleteManualCard(originalIndex);
                } else {
                    deleteCard(originalIndex);
                }
            };

            const deleteBtn = card.querySelector('.card-delete-btn');
            if (deleteBtn) {
                deleteBtn.onclick = deleteFn;
            }
            card.addEventListener('dblclick', deleteFn);
            
            cardsGrid.appendChild(card);
        });
        
        const pageInfoElement = document.getElementById('pageInfo');
        if (pageInfoElement) {
            pageInfoElement.textContent = `Page ${page} of ${Math.ceil(cards.length / cardsPerPage)}`;
        }
        
        const prevPageElement = document.getElementById('prevPage');
        const nextPageElement = document.getElementById('nextPage');
        
        if (prevPageElement) {
            prevPageElement.disabled = page <= 1;
            prevPageElement.style.opacity = page <= 1 ? '0.5' : '1';
        }
        if (nextPageElement) {
            nextPageElement.disabled = page >= Math.ceil(cards.length / cardsPerPage);
            nextPageElement.style.opacity = page >= Math.ceil(cards.length / cardsPerPage) ? '0.5' : '1';
        }
    }

    // Create enhanced preview container
    const previewContainer = document.createElement('div');
    previewContainer.style.cssText = `
        background: white;
        border-radius: 12px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
        overflow: hidden;
        margin: 1rem 0;
    `;
    
    // Create header with controls
    const header = document.createElement('div');
    header.style.cssText = `
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1.5rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 1rem;
    `;
    
    // Add instruction text
    const instructionText = document.createElement('div');
    instructionText.style.cssText = `
        font-size: 0.9rem;
        opacity: 0.9;
        margin-bottom: 1rem;
        text-align: center;
        width: 100%;
    `;
    instructionText.innerHTML = `
        <span style="margin-right: 1rem;">🔍 Click on any card to view details</span>
        <span style="margin-right: 1rem;">🗑️ Use delete buttons or double-click to remove cards</span>
        <span>📄 Use pagination to navigate through pages</span>
    `;
    header.appendChild(instructionText);
    
    // Search and filter section
    const searchSection = document.createElement('div');
    searchSection.style.cssText = `
        display: flex;
        gap: 1rem;
        align-items: center;
        flex-wrap: wrap;
    `;
    
    searchSection.innerHTML = `
        <div style="position: relative;">
            <input type="text" id="cardSearch" placeholder="Search cards..." style="
                padding: 0.75rem 1rem 0.75rem 2.5rem;
                border: none;
                border-radius: 8px;
                font-size: 1rem;
                width: 250px;
                background: rgba(255, 255, 255, 0.9);
            ">
            <span style="position: absolute; left: 0.75rem; top: 50%; transform: translateY(-50%); color: #666;">🔍</span>
        </div>
        <select id="cardFilter" style="
            padding: 0.75rem 1rem;
            border: none;
            border-radius: 8px;
            font-size: 1rem;
            background: rgba(255, 255, 255, 0.9);
            cursor: pointer;
        ">
            <option value="all">All Cards</option>
            <option value="excel">Excel Cards</option>
            <option value="manual">Manual Cards</option>
        </select>
    `;
    
    // Pagination controls
    const paginationSection = document.createElement('div');
    paginationSection.style.cssText = `
        display: flex;
        gap: 0.5rem;
        align-items: center;
    `;
    
    const cardsPerPage = 8;
    const totalPages = Math.ceil(allCards.length / cardsPerPage);
    
    paginationSection.innerHTML = `
        <button id="prevPage" style="
            padding: 0.5rem 1rem;
            border: none;
            border-radius: 6px;
            background: rgba(255, 255, 255, 0.2);
            color: white;
            cursor: pointer;
            font-weight: 600;
        ">← Previous</button>
        <span id="pageInfo" style="
            padding: 0.5rem 1rem;
            background: rgba(255, 255, 255, 0.2);
            border-radius: 6px;
            font-weight: 600;
        ">Page 1 of ${totalPages}</span>
        <button id="nextPage" style="
            padding: 0.5rem 1rem;
            border: none;
            border-radius: 6px;
            background: rgba(255, 255, 255, 0.2);
            color: white;
            cursor: pointer;
            font-weight: 600;
        ">Next →</button>
    `;
    
    header.appendChild(searchSection);
    header.appendChild(paginationSection);
    previewContainer.appendChild(header);
    
    // Cards display area
    const cardsArea = document.createElement('div');
    cardsArea.id = 'cardsDisplayArea';
    cardsArea.style.cssText = `
        padding: 2rem;
        min-height: 400px;
        background: #f8fafc;
    `;
    
    // Create cards grid
    const cardsGrid = document.createElement('div');
    cardsGrid.id = 'cardsGrid';
    cardsGrid.style.cssText = `
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
        gap: 1.5rem;
        margin-bottom: 2rem;
    `;
    
    cardsArea.appendChild(cardsGrid);
    previewContainer.appendChild(cardsArea);
    
    preview.appendChild(previewContainer);

    // Initial display
    displayCards(allCards, 1);
    
    // Add event listeners
    let currentPage = 1;
    let filteredCards = allCards;

    try {
        const prevPageElement = document.getElementById('prevPage');
        const nextPageElement = document.getElementById('nextPage');
        const cardSearchElement = document.getElementById('cardSearch');
        const cardFilterElement = document.getElementById('cardFilter');

        const updateFiltering = () => {
            const searchTerm = (cardSearchElement ? cardSearchElement.value : '').toLowerCase();
            const filterType = cardFilterElement ? cardFilterElement.value : 'all';
            
            filteredCards = allCards.filter(cardData => {
                const { isManual } = cardData;
                const matchesFilter = filterType === 'all' || 
                                    (filterType === 'excel' && !isManual) || 
                                    (filterType === 'manual' && isManual);
                
                const studentName = (cardData['NOM & PRENOM'] || '').toLowerCase();
                const adherent = (cardData['ADHERENT'] || '').toLowerCase();

                const matchesSearch = studentName.includes(searchTerm) || adherent.includes(searchTerm);

                return matchesFilter && matchesSearch;
            });
            
            currentPage = 1;
            displayCards(filteredCards, currentPage);
        };

        if (prevPageElement) {
            prevPageElement.addEventListener('click', () => {
                if (currentPage > 1) {
                    currentPage--;
                    displayCards(filteredCards, currentPage);
                }
            });
        }

        if (nextPageElement) {
            nextPageElement.addEventListener('click', () => {
                if (currentPage < Math.ceil(filteredCards.length / cardsPerPage)) {
                    currentPage++;
                    displayCards(filteredCards, currentPage);
                }
            });
        }

        if (cardSearchElement) {
            cardSearchElement.addEventListener('input', updateFiltering);
        }

        if (cardFilterElement) {
            cardFilterElement.addEventListener('change', updateFiltering);
        }
    } catch (error) {
        console.error("Error setting up preview event listeners:", error);
        alert("A critical error occurred while setting up the preview controls. Please check the console for details.");
    }
    
    // Manual cards section (simplified)
    if (manuallyAddedCards.length > 0) {
        if (manualCardsSection) {
            manualCardsSection.style.display = 'block';
            manualCardsSection.innerHTML = `
                <div style="
                    background: #f0f9ff;
                    border: 1px solid #0ea5e9;
                    border-radius: 8px;
                    padding: 1rem;
                    margin: 1rem 0;
                    text-align: center;
                ">
                    <h3 style="color: #0c4a6e; margin: 0;">Manual Cards Included</h3>
                    <p style="color: #0369a1; margin: 0.5rem 0 0 0;">
                        ${manuallyAddedCards.length} manually added card(s) are included in the preview above.
                    </p>
                </div>
            `;
        }
    } else {
        if (manualCardsSection) {
            manualCardsSection.style.display = 'none';
        }
    }
    
    updateCardCount();
}

function deleteCard(index) {
    deletedCards.add(index);
    renderPreview();
    updateCardCount();
}

function deleteManualCard(index) {
    manuallyAddedCards.splice(index, 1);
    renderPreview();
    updateCardCount();
}

function restoreAllCards() {
    deletedCards.clear();
    renderPreview();
    updateCardCount();
}

function updateCardCount() {
    const filteredData = excelData.filter((student, index) => 
        !deletedCards.has(index) && (
        student['NOM & PRENOM'] ||
        student['TEL ADH'] ||
        student['TEL PARENT'] ||
        student['LYCEE'] ||
        student['ADHERENT']
        )
    );
    const totalCards = excelData.length;
    const deletedCount = deletedCards.size;
    const activeCards = filteredData.length + manuallyAddedCards.length;
    document.getElementById('cardCount').textContent = `Active cards: ${activeCards} | Deleted: ${deletedCount} | Total: ${totalCards} | Added: ${manuallyAddedCards.length}`;
    
    // Show/hide restore button
    const restoreBtn = document.getElementById('restoreBtn');
    if (deletedCount > 0) {
        restoreBtn.style.display = 'flex';
    } else {
        restoreBtn.style.display = 'none';
    }
}

function addCardManually() {
    console.log('addCardManually function called');
    
    // Check if we have the required data
    if (!selectedFields || selectedFields.length === 0) {
        if (excelData && excelData.length > 0) {
             // This can happen if the excel file has no headers that were parsed
             alert('Could not determine fields from Excel file. Please check the file format.');
        } else {
            alert('Please upload an Excel file first to set up the field structure.');
        }
        return;
    }
    
    // Show options modal first
    const optionsModal = document.createElement('div');
    optionsModal.style.position = 'fixed';
    optionsModal.style.top = '0';
    optionsModal.style.left = '0';
    optionsModal.style.width = '100%';
    optionsModal.style.height = '100%';
    optionsModal.style.backgroundColor = 'rgba(0, 0, 0, 0.5)';
    optionsModal.style.display = 'flex';
    optionsModal.style.justifyContent = 'center';
    optionsModal.style.alignItems = 'center';
    optionsModal.style.zIndex = '10000';

    const optionsContent = document.createElement('div');
    optionsContent.style.backgroundColor = 'white';
    optionsContent.style.padding = '2rem';
    optionsContent.style.borderRadius = '12px';
    optionsContent.style.boxShadow = '0 4px 20px rgba(0, 0, 0, 0.3)';
    optionsContent.style.maxWidth = '400px';
    optionsContent.style.width = '90%';
    optionsContent.style.textAlign = 'center';

    optionsContent.innerHTML = `
        <h2 style="margin-bottom: 1.5rem; color: #1a73e8;">Add Cards</h2>
        <p style="margin-bottom: 2rem; color: #666;">Choose how you want to add cards:</p>
        
        <div style="display: flex; flex-direction: column; gap: 1rem;">
            <button id="singleCardBtn" style="padding: 1rem; background: #1a73e8; color: white; border: none; border-radius: 8px; cursor: pointer; font-weight: 600; font-size: 1rem;">
                Add Single Card
            </button>
            
            <button id="bulkCardBtn" style="padding: 1rem; background: #34d399; color: white; border: none; border-radius: 8px; cursor: pointer; font-weight: 600; font-size: 1rem;">
                Add Multiple Cards
            </button>
            
            <button id="cancelOptionsBtn" style="padding: 1rem; border: 2px solid #e2e8f0; background: white; border-radius: 8px; cursor: pointer; font-weight: 600; font-size: 1rem;">
                Cancel
            </button>
        </div>
    `;

    optionsModal.appendChild(optionsContent);
    document.body.appendChild(optionsModal);

    // Event listeners for options
    document.getElementById('singleCardBtn').addEventListener('click', () => {
        document.body.removeChild(optionsModal);
        addSingleCard();
    });

    document.getElementById('bulkCardBtn').addEventListener('click', () => {
        document.body.removeChild(optionsModal);
        addBulkCards();
    });

    document.getElementById('cancelOptionsBtn').addEventListener('click', () => {
        document.body.removeChild(optionsModal);
    });
}

function addSingleCard() {
    // Calculate next index
    const activeExcelCards = (excelData || []).filter((_, index) => !deletedCards.has(index)).length;
    const nextIndex = activeExcelCards + manuallyAddedCards.length;
    
    createCardForm(nextIndex, false);
}

function addBulkCards() {
    // Show bulk input modal
    const bulkModal = document.createElement('div');
    bulkModal.style.position = 'fixed';
    bulkModal.style.top = '0';
    bulkModal.style.left = '0';
    bulkModal.style.width = '100%';
    bulkModal.style.height = '100%';
    bulkModal.style.backgroundColor = 'rgba(0, 0, 0, 0.5)';
    bulkModal.style.display = 'flex';
    bulkModal.style.justifyContent = 'center';
    bulkModal.style.alignItems = 'center';
    bulkModal.style.zIndex = '10000';

    const bulkContent = document.createElement('div');
    bulkContent.style.backgroundColor = 'white';
    bulkContent.style.padding = '2rem';
    bulkContent.style.borderRadius = '12px';
    bulkContent.style.boxShadow = '0 4px 20px rgba(0, 0, 0, 0.3)';
    bulkContent.style.maxWidth = '400px';
    bulkContent.style.width = '90%';
    bulkContent.style.textAlign = 'center';

    bulkContent.innerHTML = `
        <h2 style="margin-bottom: 1.5rem; color: #1a73e8;">Add Multiple Cards</h2>
        <p style="margin-bottom: 1rem; color: #666;">How many cards do you want to create?</p>
        
        <div style="margin-bottom: 2rem;">
            <input type="number" id="cardCount" min="1" max="50" value="1" style="width: 100%; padding: 0.75rem; border: 2px solid #e2e8f0; border-radius: 8px; font-size: 1rem; text-align: center;">
        </div>
        
        <div style="display: flex; gap: 1rem; justify-content: center;">
            <button id="cancelBulkBtn" style="padding: 0.75rem 1.5rem; border: 2px solid #e2e8f0; background: white; border-radius: 8px; cursor: pointer; font-weight: 600;">Cancel</button>
            <button id="startBulkBtn" style="padding: 0.75rem 1.5rem; background: #34d399; color: white; border: none; border-radius: 8px; cursor: pointer; font-weight: 600;">Start</button>
        </div>
    `;

    bulkModal.appendChild(bulkContent);
    document.body.appendChild(bulkModal);

    document.getElementById('startBulkBtn').addEventListener('click', () => {
        const count = parseInt(document.getElementById('cardCount').value);
        if (count < 1 || count > 50) {
            alert('Please enter a number between 1 and 50.');
            return;
        }
        document.body.removeChild(bulkModal);
        startBulkCardCreation(count);
    });

    document.getElementById('cancelBulkBtn').addEventListener('click', () => {
        document.body.removeChild(bulkModal);
    });
}

function startBulkCardCreation(totalCards) {
    let currentCard = 0;
    let bulkCards = [];
    let isStopped = false;
    
    function createNextCard() {
        if (currentCard >= totalCards || isStopped) {
            // All cards created or process stopped
            if (bulkCards.length > 0) {
                manuallyAddedCards.push(...bulkCards);
                renderPreview();
                updateCardCount();
                if (isStopped) {
                    alert(`Bulk creation stopped. ${bulkCards.length} cards were successfully created.`);
            } else {
                    alert(`Successfully created ${totalCards} cards!`);
                }
            }
            return;
        }
        
        currentCard++;
        const totalExcelCards = excelData ? excelData.length : 0;
        const activeExcelCards = excelData ? excelData.filter((student, index) => !deletedCards.has(index)).length : 0;
        const nextIndex = activeExcelCards + manuallyAddedCards.length + bulkCards.length + 1;
        
        createCardForm(nextIndex, true, (cardData) => {
            bulkCards.push(cardData);
            createNextCard();
        }, () => {
            // Stop callback
            isStopped = true;
            createNextCard();
        });
    }
    
    createNextCard();
}

function createCardForm(nextIndex, isBulk = false, onComplete = null, onStop = null) {
    let tempPhotoDataURL = null;

    const modal = document.createElement('div');
    modal.className = 'add-card-modal';

    const modalContent = document.createElement('div');
    modalContent.className = 'add-card-modal-content';

    const formContainer = document.createElement('div');
    formContainer.className = 'add-card-form-container';

    const form = document.createElement('form');
    const bulkProgress = isBulk ? `<div class="bulk-progress">
        <strong>Bulk Mode:</strong> Creating card ${nextIndex}
    </div>` : '';
    
    let formHTML = `
        <h2 class="form-title">Add New Card (Index: ${nextIndex})</h2>
        ${bulkProgress}
        
        <div class="form-group">
            <label>Photo:</label>
            <div class="photo-upload-area">
                <input type="file" id="addPhoto_${nextIndex}" accept="image/*" class="file-input">
                <button type="button" id="uploadPhotoBtn_${nextIndex}" class="btn btn-primary">
                    <svg width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M.5 9.9a.5.5 0 0 1 .5.5v2.5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-2.5a.5.5 0 0 1 1 0v2.5a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2v-2.5a.5.5 0 0 1 .5-.5z"/><path d="M7.646 11.854a.5.5 0 0 0 .708 0l3-3a.5.5 0 0 0-.708-.708L8.5 10.293V1.5a.5.5 0 0 0-1 0v8.793L5.354 8.146a.5.5 0 1 0-.708.708l3 3z"/></svg>
                    Upload Photo
                </button>
                <span id="photoFileName_${nextIndex}" class="file-name"></span>
            </div>
            <small class="form-text">Click the button to upload a photo for this card</small>
        </div>
    `;

    selectedFields.forEach(field => {
        formHTML += `
            <div class="form-group">
                <label for="add_${field.replace(/[^a-zA-Z0-9]/g, '_')}_${nextIndex}">${field}:</label>
                <input type="text" id="add_${field.replace(/[^a-zA-Z0-9]/g, '_')}_${nextIndex}" class="form-control">
            </div>
        `;
    });

    const submitButtonText = isBulk ? 'Add & Next' : 'Add Card';
    const stopButton = isBulk ? `<button type="button" id="stopBulkBtn_${nextIndex}" class="btn btn-danger">Stop & Save</button>` : '';
    
    formHTML += `
        <div class="form-actions">
            ${stopButton}
            <button type="button" id="cancelAdd_${nextIndex}" class="btn btn-secondary">Cancel</button>
            <button type="submit" class="btn btn-primary">${submitButtonText}</button>
        </div>
    `;
    form.innerHTML = formHTML;

    const previewContainer = document.createElement('div');
    previewContainer.className = 'add-card-preview-container';
    previewContainer.innerHTML = `
        <h3 class="preview-title">Live Preview</h3>
        <div class="live-preview-wrapper">
            <div id="livePreviewCard_${nextIndex}"></div>
        </div>
    `;

    function updateLivePreview() {
        const previewDiv = document.getElementById(`livePreviewCard_${nextIndex}`);
        if (!previewDiv) return;

        const tempCard = {};
        selectedFields.forEach(field => {
            const fieldId = `add_${field.replace(/[^a-zA-Z0-9]/g, '_')}_${nextIndex}`;
            const input = form.querySelector('#' + fieldId);
            if (input) {
                tempCard[field] = input.value;
            }
        });
        
        const card = createIDCard(tempCard, nextIndex, false, tempPhotoDataURL);
        previewDiv.innerHTML = '';
        previewDiv.appendChild(card);
    }

    formContainer.appendChild(form);
    modalContent.appendChild(formContainer);
    modalContent.appendChild(previewContainer);
    modal.appendChild(modalContent);
    
    document.body.appendChild(modal);

    // Defer event listener setup
    setTimeout(() => {
        modal.classList.add('visible'); // Make the modal visible with animation

        const photoInput = modal.querySelector(`#addPhoto_${nextIndex}`);
        const uploadBtn = modal.querySelector(`#uploadPhotoBtn_${nextIndex}`);
        const cancelBtn = modal.querySelector(`#cancelAdd_${nextIndex}`);

        if (uploadBtn) {
            uploadBtn.addEventListener('click', () => photoInput.click());
        }

        if (photoInput) {
            photoInput.addEventListener('change', function(e) {
                if (e.target.files.length > 0) {
                    const file = e.target.files[0];
                    modal.querySelector(`#photoFileName_${nextIndex}`).textContent = file.name;
                    const reader = new FileReader();
                    reader.onload = function(e) {
                        tempPhotoDataURL = e.target.result;
                        updateLivePreview();
                    };
                    reader.readAsDataURL(file);
                }
            });
        }
        
        const closeModal = () => {
            modal.classList.remove('visible');
            setTimeout(() => {
                if (document.body.contains(modal)) {
                    document.body.removeChild(modal);
                }
            }, 300); // Wait for transition to finish
        };

        if (cancelBtn) {
            cancelBtn.addEventListener('click', closeModal);
        }

        form.addEventListener('input', updateLivePreview);
        form.onsubmit = (e) => {
            e.preventDefault();
            
            const newCard = {};
            selectedFields.forEach(field => {
                const fieldId = `add_${field.replace(/[^a-zA-Z0-9]/g, '_')}_${nextIndex}`;
                const input = form.querySelector('#' + fieldId);
                if (input) {
                    newCard[field] = input.value;
                }
            });

            if (tempPhotoDataURL) {
                const photoId = (newCard['ADHERENT'] || `manual_${Date.now()}`).toLowerCase();
                images.set(photoId, tempPhotoDataURL);
            }

            if (isBulk && onComplete) {
                onComplete(newCard);
            } else {
                manuallyAddedCards.push(newCard);
                renderPreview();
                updateCardCount();
            }
            closeModal();
        };
        
        if (isBulk && onStop) {
            const stopBtn = modal.querySelector('#stopBulkBtn_'+nextIndex);
            if (stopBtn) {
                stopBtn.addEventListener('click', () => {
                    modal.remove(); 
                    onStop();
                });
            }
        }

        updateLivePreview();
    }, 0);
}

function exportManualDataToExcel() {
    const activeExcelCards = (excelData || []).filter((_, index) => !deletedCards.has(index));
    const allCardsData = [...activeExcelCards, ...manuallyAddedCards];

    if (allCardsData.length === 0) {
        alert('No cards to export.');
        return;
    }

    const headers = selectedFields.length > 0 ? selectedFields : Object.keys(allCardsData[0] || {});
    if (headers.length === 0) {
        alert('No fields selected or available to export.');
        return;
    }
    
    const worksheetData = allCardsData.map(card => {
        return headers.map(header => card[header] || '');
    });

    worksheetData.unshift(headers);

    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

    const headerStyle = {
        font: { bold: true, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "4F81BD" } }
    };

    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for(let c = range.s.c; c <= range.e.c; ++c) {
        const address = XLSX.utils.encode_cell({r: 0, c: c});
        if(!worksheet[address]) continue;
        worksheet[address].s = headerStyle;
    }

    const colWidths = headers.map((header, i) => ({
        wch: Math.max(
            header.length,
            ...(worksheetData.slice(1).map(row => (row[i] || '').toString().length))
        ) + 2
    }));
    worksheet['!cols'] = colWidths;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'All Cards');
    XLSX.writeFile(workbook, 'all_cards_export.xlsx');
}

function exportImagesAsZip() {
    const imageNameField = document.getElementById('imageNameField').value;
    if (!imageNameField) {
        alert('Please select a field to name the images by.');
        return;
    }

    if (manuallyAddedCards.length === 0 || images.size === 0) {
        alert('No manually added cards or images available to export.');
        return;
    }

    const zip = new JSZip();
    let imageCount = 0;

    manuallyAddedCards.forEach(card => {
        const adherentId = String(card['ADHERENT'] || '').trim().toLowerCase();
        const photoSrc = images.get(adherentId);

        if (photoSrc) {
            const fileNameValue = card[imageNameField] || `card_${adherentId || (imageCount + 1)}`;
            const safeFileName = String(fileNameValue).replace(/[^a-z0-9_-\s\.]/gi, '').trim();
            const fileExtension = photoSrc.substring(photoSrc.indexOf('/') + 1, photoSrc.indexOf(';'));
            const finalFileName = `${safeFileName}.${fileExtension}`;
            
            const base64Data = photoSrc.split(';base64,').pop();
            zip.file(finalFileName, base64Data, { base64: true });
            imageCount++;
        }
    });

    if (imageCount === 0) {
        alert('No matching images found for the manually added cards.');
        return;
    }

    zip.generateAsync({ type: 'blob' }).then(function(content) {
        const link = document.createElement('a');
        link.href = URL.createObjectURL(content);
        link.download = 'exported_images.zip';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });
}

function showExportModal() {
    const modalOverlay = document.createElement('div');
    modalOverlay.className = 'export-modal-overlay';
    
    const modalContent = document.createElement('div');
    modalContent.className = 'export-modal-content';

    modalContent.innerHTML = `
        <h2>PDF Generated!</h2>
        <p>Would you like to export associated data?</p>
        <div class="export-modal-buttons">
            <button id="modalExportExcelBtn" class="btn btn-primary">Export All Cards to Excel</button>
            <div class="image-export-group">
                <button id="modalExportImagesBtn" class="btn btn-primary">Export Images as .zip</button>
                <div class="input-group">
                    <label for="modalImageNameField">Name images by:</label>
                    <select id="modalImageNameField"></select>
                </div>
            </div>
            <button id="modalCloseBtn" class="btn btn-secondary">Close</button>
        </div>
    `;

    modalOverlay.appendChild(modalContent);
    document.body.appendChild(modalOverlay);

    // Populate dropdown
    const originalDropdown = document.getElementById('imageNameField');
    const modalDropdown = document.getElementById('modalImageNameField');
    modalDropdown.innerHTML = originalDropdown.innerHTML;

    // Add event listeners
    document.getElementById('modalExportExcelBtn').addEventListener('click', exportManualDataToExcel);
    document.getElementById('modalExportImagesBtn').addEventListener('click', () => {
        // Temporarily move the selected value for the export function
        const originalSelect = document.getElementById('imageNameField');
        const modalSelect = document.getElementById('modalImageNameField');
        originalSelect.value = modalSelect.value;
        exportImagesAsZip();
    });
    document.getElementById('modalCloseBtn').addEventListener('click', () => {
        modalOverlay.classList.remove('visible');
        setTimeout(() => {
        document.body.removeChild(modalOverlay);
        }, 300); // Wait for animation
    });

    // Make it visible
    setTimeout(() => {
        modalOverlay.classList.add('visible');
    }, 10);
}

function openEmailModal() {
    if (!excelData || excelData.length === 0) {
        alert('Please upload an Excel file first.');
        return;
    }

    const modal = document.getElementById('emailConfigModal');
    const emailColumnSelect = document.getElementById('emailColumn');
    
    // Populate email column selector
    emailColumnSelect.innerHTML = '';
    const headers = Object.keys(excelData[0]);
    headers.forEach(header => {
        const option = document.createElement('option');
        option.value = header;
        option.textContent = header;
        if (header.toLowerCase().includes('email')) {
            option.selected = true;
        }
        emailColumnSelect.appendChild(option);
    });

    // Load saved config from localStorage
    document.getElementById('emailServiceId').value = localStorage.getItem('emailJS_serviceId') || '';
    document.getElementById('emailTemplateId').value = localStorage.getItem('emailJS_templateId') || '';
    document.getElementById('emailUserId').value = localStorage.getItem('emailJS_userId') || '';

    modal.classList.add('visible');
}

async function sendAllEmails() {
    const serviceId = document.getElementById('emailServiceId').value.trim();
    const templateId = document.getElementById('emailTemplateId').value.trim();
    const userId = document.getElementById('emailUserId').value.trim();
    const emailColumn = document.getElementById('emailColumn').value;

    if (!serviceId || !templateId || !userId || !emailColumn) {
        alert('Please fill out all EmailJS configuration fields and select the email column.');
        return;
    }

    // Save config to localStorage
    localStorage.setItem('emailJS_serviceId', serviceId);
    localStorage.setItem('emailJS_templateId', templateId);
    localStorage.setItem('emailJS_userId', userId);

    const mainTitleField = document.getElementById('mainTitleField').value;
    const activeCards = excelData.filter((_, index) => !deletedCards.has(index));
    const allCards = [...activeCards, ...manuallyAddedCards];
    
    let successCount = 0;
    let errorCount = 0;

    const statusDiv = document.createElement('div');
    statusDiv.style.marginTop = '1rem';
    statusDiv.className = 'form-text';
    document.querySelector('#emailConfigModal .add-card-form-container').appendChild(statusDiv);

    for (let i = 0; i < allCards.length; i++) {
        const cardData = allCards[i];
        const recipientEmail = cardData[emailColumn];
        const studentName = cardData[mainTitleField] || 'Student';

        if (!recipientEmail) {
            console.warn(`Skipping card ${i+1} due to missing email.`);
            continue;
        }

        statusDiv.textContent = `Sending email ${i + 1} of ${allCards.length} to ${recipientEmail}...`;

        try {
            const qrCodeDataUrl = await generateQrCodeDataUrl(cardData, i);
            
            const templateParams = {
                to_email: recipientEmail,
                student_name: studentName,
                qr_code_image: qrCodeDataUrl,
            };

            await emailjs.send(serviceId, templateId, templateParams, userId);
            successCount++;

        } catch (error) {
            console.error(`Failed to send email to ${recipientEmail}:`, error);
            errorCount++;
        }
    }

    statusDiv.innerHTML = `
        <strong>Email sending complete!</strong><br>
        Successful: ${successCount}<br>
        Failed: ${errorCount}
    `;

    setTimeout(() => {
        document.getElementById('emailConfigModal').classList.remove('visible');
        statusDiv.remove();
    }, 5000);
}

function generateQrCodeDataUrl(student, index) {
    return new Promise((resolve, reject) => {
        try {
            const headerField = document.getElementById('headerField').value;
            const mainTitleField = document.getElementById('mainTitleField').value;
            const id = String(student['ADHERENT'] || '').trim().toLowerCase();
            const type = String(student['type'] || student['TYPE'] || '').trim().toLowerCase();
            let photoSrc = images.get(id) || ((type === 'f') ? defaultFemaleImage : defaultMaleImage);

            const headerText = student[headerField] || `ID: ${String(index + 1).padStart(3, '0')}`;
            const mainTitleText = student[mainTitleField] || ' ';

            const digitalCardData = {
                header: headerText,
                title: mainTitleText,
                photo: photoSrc,
                fields: selectedFields.reduce((obj, field) => {
                    if (field !== headerField && field !== mainTitleField) {
                        obj[field] = student[field] || '';
                    }
                    return obj;
                }, {})
            };

            const jsonString = JSON.stringify(digitalCardData);
            const encodedData = encodeURIComponent(jsonString);
            // TODO: Replace with your hosting URL when deployed
            // Example: const digitalCardUrl = `https://yourusername.github.io/id-card-generator/digital_card.html?id=${index}&data=${encodedData}`;
            const digitalCardUrl = `http://localhost:8000/digital_card.html?id=${index}&data=${encodedData}`;

            const tempContainer = document.createElement('div');
            new QRCode(tempContainer, {
                text: digitalCardUrl,
                width: 256,
                height: 256,
                correctLevel: QRCode.CorrectLevel.L
            });

            setTimeout(() => {
                const img = tempContainer.querySelector('img');
                if (img && img.src) {
                    resolve(img.src);
                } else {
                    reject('Could not generate QR code image.');
                }
            }, 100);

        } catch (error) {
            reject(error);
        }
    });
}

function generateQrCodeImages() {
    if (!excelData || excelData.length === 0) {
        alert('Please upload an Excel file first.');
        return;
    }

    const mainTitleField = document.getElementById('mainTitleField').value;
    const headerField = document.getElementById('headerField').value;
    const activeCards = excelData.filter((_, index) => !deletedCards.has(index));
    const allCards = [...activeCards, ...manuallyAddedCards];

    // Create a zip file to contain all QR codes
    const zip = new JSZip();
    let processedCount = 0;
    
    allCards.forEach((student, index) => {
        const studentName = student[mainTitleField] || `Student_${index + 1}`;
        const headerText = student[headerField] || `ID: ${String(index + 1).padStart(3, '0')}`;
        const mainTitleText = student[mainTitleField] || ' ';

        // Store full student data in localStorage
        const fullStudentData = {
            header: headerText,
            title: mainTitleText,
            photo: images.get(String(student['ADHERENT'] || '').trim().toLowerCase()) || 
                   ((String(student['type'] || student['TYPE'] || '').trim().toLowerCase() === 'f') ? defaultFemaleImage : defaultMaleImage),
            fields: selectedFields.reduce((obj, field) => {
                if (field !== headerField && field !== mainTitleField) {
                    obj[field] = student[field] || '';
                }
                return obj;
            }, {})
        };
        
        localStorage.setItem(`student_${index}`, JSON.stringify(fullStudentData));

        // Create a simpler data structure to avoid URL length issues
        const simpleData = {
            id: String(index + 1).padStart(3, '0'),
            name: mainTitleText.substring(0, 50),
            adherent: headerText.substring(0, 20)
        };

        const jsonString = JSON.stringify(simpleData);
        const encodedData = encodeURIComponent(jsonString);
        const digitalCardUrl = `http://localhost:8000/digital_card.html?id=${index}&data=${encodedData}`;

        // Create QR code as canvas with error handling
        try {
            const canvas = document.createElement('canvas');
            canvas.width = 256;
            canvas.height = 256;
            
            new QRCode(canvas, {
                text: digitalCardUrl,
                width: 256,
                height: 256,
                correctLevel: QRCode.CorrectLevel.L
            });

            // Convert canvas to blob
            canvas.toBlob((blob) => {
                const fileName = `${studentName.replace(/[^a-zA-Z0-9]/g, '_')}_QR.png`;
                zip.file(fileName, blob);
                processedCount++;
                
                // Check if all QR codes have been processed
                if (processedCount === allCards.length) {
                    // Generate and download the zip file
                    zip.generateAsync({type: 'blob'}).then((content) => {
                        const url = URL.createObjectURL(content);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = 'qr_codes.zip';
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                        URL.revokeObjectURL(url);
                        alert(`QR codes generated successfully! ${processedCount} QR codes created.\n\nTo test: Open digital_card.html in your browser and scan a QR code.`);
                    });
                }
            }, 'image/png');
            
        } catch (error) {
            console.error(`Error generating QR code for ${studentName}:`, error);
            processedCount++;
            
            // Create a fallback image or text file
            const fallbackContent = `QR Code generation failed for ${studentName}. URL: ${digitalCardUrl}`;
            zip.file(`${studentName.replace(/[^a-zA-Z0-9]/g, '_')}_ERROR.txt`, fallbackContent);
            
            if (processedCount === allCards.length) {
                zip.generateAsync({type: 'blob'}).then((content) => {
                    const url = URL.createObjectURL(content);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = 'qr_codes.zip';
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                    alert(`QR codes generated with some errors. ${processedCount} files created.`);
                });
            }
        }
    });
}

// Add a test function to generate a single QR code for testing
function generateTestQRCode() {
    const testData = {
        id: "001",
        name: "Test Student",
        adherent: "TEST001"
    };
    
    const testUrl = `digital_card.html?id=0&data=${encodeURIComponent(JSON.stringify(testData))}`;
    
    // Store test data
    localStorage.setItem('student_0', JSON.stringify({
        header: "Test ID Card",
        title: "Test Student",
        photo: defaultMaleImage,
        fields: {
            "ID": "001",
            "Name": "Test Student",
            "Department": "Testing"
        }
    }));
    
    // Create test QR code
    const testContainer = document.createElement('div');
    testContainer.style.position = 'fixed';
    testContainer.style.top = '50%';
    testContainer.style.left = '50%';
    testContainer.style.transform = 'translate(-50%, -50%)';
    testContainer.style.backgroundColor = 'white';
    testContainer.style.padding = '20px';
    testContainer.style.border = '2px solid #333';
    testContainer.style.zIndex = '10000';
    
    const closeBtn = document.createElement('button');
    closeBtn.textContent = 'Close';
    closeBtn.onclick = () => document.body.removeChild(testContainer);
    closeBtn.style.marginBottom = '10px';
    
    testContainer.appendChild(closeBtn);
    testContainer.appendChild(document.createElement('br'));
    
    new QRCode(testContainer, {
        text: testUrl,
        width: 200,
        height: 200,
        correctLevel: QRCode.CorrectLevel.L
    });
    
    document.body.appendChild(testContainer);
} 