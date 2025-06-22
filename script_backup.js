document.addEventListener('DOMContentLoaded', () => {
    let excelData = null;
    let images = new Map();
    let uploadedLogoBase64 = '';
    let selectedFields = [];

    const excelFileInput = document.getElementById('excelFile');
    const logoFileInput = document.getElementById('logoFile');
    const imageFilesInput = document.getElementById('imageFiles');
    const imageNameFieldSelect = document.getElementById('imageNameField');
    const generateBtn = document.getElementById('generateBtn');
    const fieldSelectionContainer = document.getElementById('fieldSelection');
    const previewContainer = document.getElementById('preview');

    excelFileInput.addEventListener('change', handleExcelUpload);
    logoFileInput.addEventListener('change', handleLogoUpload);
    imageFilesInput.addEventListener('change', handleImageUpload);
    generateBtn.addEventListener('click', generatePDF);

    function handleExcelUpload(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            
            let headerRow = -1;
            const range = XLSX.utils.decode_range(firstSheet['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                const cell = firstSheet[XLSX.utils.encode_cell({c:0, r:R})];
                if (cell && cell.v && String(cell.v).toLowerCase().includes('nom')) {
                    headerRow = R;
                    break;
                }
            }

            if (headerRow === -1) {
                alert('Could not find a header row containing "nom". Trying to read from the first row.');
                headerRow = 0; 
            }
            
            excelData = XLSX.utils.sheet_to_json(firstSheet, { range: headerRow });
            
            const headers = Object.keys(excelData[0]);
            
            fieldSelectionContainer.innerHTML = '';
            headers.forEach(header => {
                const p = document.createElement('p');
                const label = document.createElement('label');
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.value = header;
                checkbox.checked = true;
                checkbox.addEventListener('change', updateSelectedFields);
                label.appendChild(checkbox);
                label.appendChild(document.createTextNode(` ${header}`));
                p.appendChild(label);
                fieldSelectionContainer.appendChild(p);
            });
            
            imageNameFieldSelect.innerHTML = '';
            headers.forEach(header => {
                const option = document.createElement('option');
                option.value = header;
                option.textContent = header;
                imageNameFieldSelect.appendChild(option);
            });
            const defaultImageField = headers.find(h => h.toUpperCase().includes('ADHERENT')) || headers.find(h => h.toUpperCase().includes('NOM')) || headers[0];
            if (defaultImageField) {
                imageNameFieldSelect.value = defaultImageField;
            }

            updateSelectedFields();
            generateBtn.disabled = false;
        };
        reader.readAsArrayBuffer(file);
    }

    function handleLogoUpload(event) {
        const file = event.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            uploadedLogoBase64 = e.target.result;
            renderPreview();
        };
        reader.readAsDataURL(file);
    }

    function handleImageUpload(event) {
        const files = event.target.files;
        if (!files.length) return;
        let loadedCount = 0;
        images.clear();
        for (const file of files) {
            const reader = new FileReader();
            reader.onload = (e) => {
                const fileName = file.name.split('.').slice(0, -1).join('.').toLowerCase();
                images.set(fileName, e.target.result);
                loadedCount++;
                if (loadedCount === files.length) {
                    renderPreview();
                }
            };
            reader.readAsDataURL(file);
        }
    }

    function updateSelectedFields() {
        selectedFields = Array.from(fieldSelectionContainer.querySelectorAll('input:checked')).map(cb => cb.value);
        renderPreview();
    }

    function renderPreview() {
        if (!excelData) return;
        previewContainer.innerHTML = '';
        excelData.forEach((student, index) => {
            const card = createIDCard(student, index);
            previewContainer.appendChild(card);
        });
    }

    function createIDCard(student, index) {
        const card = document.createElement('div');
        card.className = 'id-card';

        const indexNumber = document.createElement('div');
        indexNumber.className = 'id-card-index-number';
        indexNumber.textContent = String(index + 1).padStart(2, '0');
        card.appendChild(indexNumber);

        const mainContainer = document.createElement('div');
        mainContainer.className = 'id-card-main-container';

        const header = document.createElement('div');
        header.className = 'id-card-header';
        const logo = document.createElement('img');
        logo.className = 'id-card-logo';
        logo.src = uploadedLogoBase64 || 'ticket logo.png';
        header.appendChild(logo);
        const adherentId = document.createElement('div');
        adherentId.className = 'id-card-adherent';
        adherentId.textContent = student['ADHERENT'] || '';
        header.appendChild(adherentId);
        mainContainer.appendChild(header);

        const photoWrapper = document.createElement('div');
        photoWrapper.className = 'id-card-photo-wrapper';
        const photo = document.createElement('img');
        photo.className = 'id-card-photo';

        const imageNameField = imageNameFieldSelect.value;
        const imageName = (student[imageNameField] || '').toLowerCase();
        photo.src = images.get(imageName) || '';
        photoWrapper.appendChild(photo);
        mainContainer.appendChild(photoWrapper);

        const content = document.createElement('div');
        content.className = 'id-card-content';
        selectedFields.forEach(field => {
            if (field.toUpperCase() === 'ADHERENT') return;
            let value = student[field] || '';
            const row = document.createElement('div');
            
            if (field.toUpperCase() === 'NOM & PRENOM') {
                row.className = 'id-card-name';
                row.textContent = `${index + 1}. ${value}`;
            } else {
                row.className = 'id-card-info-row';
                row.innerHTML = `<div class="id-card-label">${field.toUpperCase()}:</div> <div class="id-card-value">${value}</div>`;
            }
            content.appendChild(row);
        });
        mainContainer.appendChild(content);
        card.appendChild(mainContainer);

        return card;
    }

    async function generatePDF() {
        const { jsPDF } = window.jspdf;
        if (!jsPDF) {
            alert('PDF library not loaded!');
            return;
        }
        if (!excelData) {
            alert('No data to generate PDF.');
            return;
        }

        const pdf = new jsPDF({
            orientation: 'portrait',
            unit: 'mm',
            format: 'a4'
        });

        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const cardsPerPage = 8;
        const cardsPerRow = 2;
        const rowsPerPage = 4;
        const cardWidth = pageWidth / cardsPerRow;
        const cardHeight = pageHeight / rowsPerPage;

        for (let i = 0; i < excelData.length; i++) {
            const pageIndex = Math.floor(i / cardsPerPage);
            if (i > 0 && i % cardsPerPage === 0) {
                pdf.addPage();
            }

            const student = excelData[i];
            const cardElement = createIDCard(student, i);
            cardElement.style.width = '325px'; 
            document.body.appendChild(cardElement);

            const cardIndexOnPage = i % cardsPerPage;
            const row = Math.floor(cardIndexOnPage / cardsPerRow);
            const col = cardIndexOnPage % cardsPerRow;
            const x = col * cardWidth;
            const y = row * cardHeight;

            try {
                const canvas = await html2canvas(cardElement, { scale: 3, useCORS: true });
                const imgData = canvas.toDataURL('image/png');
                pdf.addImage(imgData, 'PNG', x, y, cardWidth, cardHeight);
            } catch (error) {
                console.error('Error generating canvas for card:', error);
            } finally {
                document.body.removeChild(cardElement);
            }
        }

        const fileName = document.getElementById('pdfFileName').value || 'ID_Cards';
        pdf.save(fileName + '.pdf');
    }
});