class ExcelProcessor {
    constructor() {
        this.init();
        this.originalData = [];
        this.processedData = [];
    }

    init() {
        const fileInput = document.getElementById('fileInput');
        const uploadArea = document.getElementById('uploadArea');

        fileInput.addEventListener('change', (e) => this.handleFile(e.target.files[0]));

        // Drag and drop functionality
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                this.handleFile(files[0]);
            }
        });

        uploadArea.addEventListener('click', () => {
            fileInput.click();
        });
    }

    async handleFile(file) {
        if (!file) return;

        // Validate file type
        const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                           'application/vnd.ms-excel'];
        if (!validTypes.includes(file.type)) {
            this.showNotification('يرجى اختيار ملف Excel صحيح (.xlsx أو .xls)', 'error');
            return;
        }

        this.showLoading(true);

        try {
            const data = await this.readExcelFile(file);
            this.originalData = data;
            this.processData();
            this.displayResults();
        } catch (error) {
            console.error('Error processing file:', error);
            this.showNotification('حدث خطأ في معالجة الملف. يرجى التأكد من تنسيق الملف.', 'error');
        } finally {
            this.showLoading(false);
        }
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    resolve(jsonData);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = () => reject(new Error('فشل في قراءة الملف'));
            reader.readAsArrayBuffer(file);
        });
    }

    processData() {
        if (this.originalData.length < 2) {
            throw new Error('الملف لا يحتوي على بيانات كافية');
        }

        // Skip header row
        const dataRows = this.originalData.slice(1);
        
        // Use Map to track latest occurrence of each item code
        const itemCodeMap = new Map();
        let removedCount = 0;
        let duplicateCount = 0;

        dataRows.forEach((row, index) => {
            try {
                // Extract columns (0-indexed)
                const itemCode = row[0]; // كود الصنف
                const itemName = row[1]; // اسم الصنف
                const unitPrice = parseFloat(row[2]) || 0; // سعر الوحدة
                const unit = row[3]; // الوحدة
                const section = row[8]; // القسم

                // Skip rows where unit is not 1 or 4
                if (unit != 1 && unit != 4) {
                    removedCount++;
                    return;
                }

                // If this item code already exists, mark the previous one as duplicate
                if (itemCodeMap.has(itemCode)) {
                    duplicateCount++;
                }

                // Always keep the latest occurrence (overwrite if exists)
                itemCodeMap.set(itemCode, {
                    itemCode: itemCode || '',
                    itemName: itemName || '',
                    originalPrice: unitPrice,
                    unit: unit,
                    section: section,
                    rowIndex: index + 2 // +2 because we skipped header and arrays are 0-indexed
                });

            } catch (error) {
                console.warn(`خطأ في السطر ${index + 2}:`, error);
            }
        });

        // Process the final unique items
        this.processedData = [];
        let priceIncreasedCount = 0;

        itemCodeMap.forEach((item) => {
            // Calculate new price with different percentages based on original price
            let newPrice = item.originalPrice;
            let priceIncreased = false;
            let percentage = 0;

            // Add percentage if section is not 52
            if (item.section != 52) {
                if (item.originalPrice >= 150) {
                    percentage = 7; // 7% for prices >= 150
                    newPrice = item.originalPrice * 1.07;
                } else {
                    percentage = 7.5; // 7.5% for prices < 150
                    newPrice = item.originalPrice * 1.075;
                }
                priceIncreased = true;
                priceIncreasedCount++;
            }

            // Apply custom rounding
            newPrice = this.customRound(newPrice);

            this.processedData.push({
                itemCode: item.itemCode,
                itemName: item.itemName,
                originalPrice: item.originalPrice,
                newPrice: newPrice,
                priceIncreased: priceIncreased,
                section: item.section,
                percentage: percentage,
                index: this.processedData.length // Add index for tracking
            });
        });

        // Calculate percentage statistics
        const sevenPercentCount = this.processedData.filter(item => item.percentage === 7).length;
        const sevenFivePercentCount = this.processedData.filter(item => item.percentage === 7.5).length;

        this.stats = {
            total: this.processedData.length,
            removed: removedCount,
            duplicates: duplicateCount,
            priceIncreased: priceIncreasedCount,
            originalTotal: dataRows.length,
            sevenPercent: sevenPercentCount,
            sevenFivePercent: sevenFivePercentCount
        };
    }

    // Custom rounding function
    customRound(price) {
        // Convert to number to avoid any string issues
        const numPrice = Number(price);
        
        // Get the whole part and decimal part
        const wholePart = Math.floor(numPrice);
        const decimalPart = numPrice - wholePart;
        
        // If already a whole number, return as is
        if (decimalPart === 0) {
            return numPrice;
        }
        
        // Round decimal part based on the rules:
        // 0.01-0.49 -> 0.50
        // 0.50-0.99 -> 0.95
        if (decimalPart > 0 && decimalPart < 0.50) {
            return wholePart + 0.50;
        } else if (decimalPart >= 0.50) {
            return wholePart + 0.95;
        } else {
            return numPrice;
        }
    }

    // Format price for display without additional rounding
    formatPrice(price) {
        // If it's a whole number, show without decimals
        if (price % 1 === 0) {
            return price.toString();
        }
        
        // For decimal numbers, check if it ends with .5 or .95
        const priceStr = price.toString();
        if (priceStr.includes('.')) {
            const decimalPart = priceStr.split('.')[1];
            if (decimalPart === '5') {
                return price.toString() + '0'; // Make .5 into .50
            } else if (decimalPart === '95') {
                return price.toString(); // Keep .95 as is
            } else {
                // For other decimals, format to 2 places but avoid over-rounding
                return parseFloat(price.toFixed(2)).toString();
            }
        }
        return price.toString();
    }

    displayResults() {
        // Show stats
        const statsDiv = document.getElementById('stats');
        statsDiv.innerHTML = `
            <div class="stat-item">
                <span class="stat-number">${this.stats.originalTotal}</span>
                <div class="stat-label">إجمالي السطور الأصلية</div>
            </div>
            <div class="stat-item">
                <span class="stat-number">${this.stats.removed}</span>
                <div class="stat-label">السطور المحذوفة (وحدة ≠ 1 أو 4)</div>
            </div>
            <div class="stat-item">
                <span class="stat-number">${this.stats.duplicates}</span>
                <div class="stat-label">الأكواد المكررة المحذوفة</div>
            </div>
            <div class="stat-item">
                <span class="stat-number">${this.stats.total}</span>
                <div class="stat-label">السطور المعروضة النهائية</div>
            </div>
            <div class="stat-item">
                <span class="stat-number">${this.stats.sevenPercent}</span>
                <div class="stat-label">زيادة 7% (أسعار ≥ 150)</div>
            </div>
            <div class="stat-item">
                <span class="stat-number">${this.stats.sevenFivePercent}</span>
                <div class="stat-label">زيادة 7.5% (أسعار < 150)</div>
            </div>
        `;

        // Display table
        const tableBody = document.getElementById('tableBody');
        tableBody.innerHTML = '';

        this.processedData.forEach((item, index) => {
            const row = document.createElement('tr');
            
            const priceClass = item.priceIncreased ? 'price-increased' : 'price-original';
            // Display price: ensure custom rounding is preserved without additional rounding
            const priceDisplay = this.formatPrice(item.newPrice);
            
            row.innerHTML = `
                <td>
                    <span class="clickable" onclick="copyToClipboard('${item.itemCode}', this)">
                        ${item.itemCode}
                    </span>
                </td>
                <td>${item.itemName}</td>
                <td>
                    <span class="clickable ${priceClass} price-cell" 
                          data-index="${index}" 
                          data-original-price="${priceDisplay}"
                          onclick="handlePriceClick(this, '${priceDisplay}', ${index})">
                        ${priceDisplay}
                    </span>
                </td>
            `;
            
            tableBody.appendChild(row);
        });

        // Show results section
        document.getElementById('resultsSection').style.display = 'block';
        
        let successMessage = `تم معالجة ${this.stats.total} صنف بنجاح`;
        if (this.stats.duplicates > 0) {
            successMessage += ` (تم حذف ${this.stats.duplicates} كود مكرر)`;
        }
        this.showNotification(successMessage, 'success');
    }

    showLoading(show) {
        document.getElementById('loading').style.display = show ? 'block' : 'none';
    }

    showNotification(message, type = 'success') {
        const notification = document.getElementById('notification');
        notification.textContent = message;
        notification.className = `notification ${type} show`;
        
        setTimeout(() => {
            notification.classList.remove('show');
        }, 3000);
    }

    // Export processed data to Excel
    exportToExcel() {
        if (this.processedData.length === 0) {
            this.showNotification('لا توجد بيانات للتصدير', 'error');
            return;
        }

        // Prepare data for export
        const exportData = [
            ['كود الصنف', 'اسم الصنف', 'سعر الوحدة'] // Header
        ];

        this.processedData.forEach(item => {
            const priceDisplay = this.formatPrice(item.newPrice);
            exportData.push([
                item.itemCode,
                item.itemName,
                parseFloat(priceDisplay)
            ]);
        });

        // Create workbook
        const ws = XLSX.utils.aoa_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'الأصناف المعدلة');

        // Generate filename with timestamp
        const now = new Date();
        const timestamp = now.toISOString().slice(0, 19).replace(/[:.]/g, '-');
        const filename = `الأصناف_المعدلة_${timestamp}.xlsx`;

        // Download file
        XLSX.writeFile(wb, filename);
        
        this.showNotification('تم تحميل الملف بنجاح', 'success');
    }
}

// Global variables for click tracking
let clickTimeouts = {};
let clickCounts = {};

// Global variable to access processor instance
let processorInstance = null;

// Copy to clipboard function with visual feedback
function copyToClipboard(text, element = null) {
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(text).then(() => {
            showCopyNotification('تم نسخ النص: ' + text);
            if (element) addCopyVisualFeedback(element);
        }).catch(err => {
            console.error('فشل في النسخ:', err);
            fallbackCopyTextToClipboard(text, element);
        });
    } else {
        fallbackCopyTextToClipboard(text, element);
    }
}

// Fallback copy function for older browsers
function fallbackCopyTextToClipboard(text, element = null) {
    const textArea = document.createElement('textarea');
    textArea.value = text;
    textArea.style.top = '0';
    textArea.style.left = '0';
    textArea.style.position = 'fixed';
    
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    
    try {
        const successful = document.execCommand('copy');
        if (successful) {
            showCopyNotification('تم نسخ النص: ' + text);
            if (element) addCopyVisualFeedback(element);
        } else {
            showCopyNotification('فشل في النسخ', 'error');
        }
    } catch (err) {
        console.error('فشل في النسخ:', err);
        showCopyNotification('فشل في النسخ', 'error');
    }
    
    document.body.removeChild(textArea);
}

// Handle price cell clicks (1-2 clicks = copy, 3 clicks = edit)
function handlePriceClick(element, priceValue, index) {
    const elementId = `price_${index}`;
    
    // Initialize click count if not exists
    if (!clickCounts[elementId]) {
        clickCounts[elementId] = 0;
    }
    
    clickCounts[elementId]++;
    
    // Clear previous timeout
    if (clickTimeouts[elementId]) {
        clearTimeout(clickTimeouts[elementId]);
    }
    
    // Check for triple click immediately if count reaches 3
    if (clickCounts[elementId] === 3) {
        // Triple click - enable editing immediately
        enablePriceEditing(element, priceValue, index);
        clickCounts[elementId] = 0; // Reset
        return;
    }
    
    // Set timeout to process single/double clicks
    clickTimeouts[elementId] = setTimeout(() => {
        const clickCount = clickCounts[elementId];
        
        if (clickCount > 0 && clickCount < 3) {
            // Single or double click - copy to clipboard
            copyToClipboard(priceValue, element);
        }
        
        // Reset click count
        clickCounts[elementId] = 0;
    }, 300); // Reduced timeout for better responsiveness
}

// Enable price editing
function enablePriceEditing(element, currentPrice, index) {
    // Prevent multiple edits on same element
    if (element.querySelector('input')) {
        return;
    }
    
    // Add editing class
    element.classList.add('editing');
    
    // Create input element
    const input = document.createElement('input');
    input.type = 'number';
    input.step = '0.1';
    input.value = currentPrice;
    input.className = 'price-edit-input';
    input.style.width = '90px';
    input.style.textAlign = 'center';
    
    // Store original content
    const originalContent = element.innerHTML;
    
    // Replace span content with input
    element.innerHTML = '';
    element.appendChild(input);
    
    // Focus and select
    setTimeout(() => {
        input.focus();
        input.select();
    }, 50);
    
    // Handle save on Enter or blur
    const saveEdit = () => {
        const newPrice = parseFloat(input.value);
        
        if (isNaN(newPrice) || newPrice <= 0) {
            showCopyNotification('سعر غير صحيح. يجب أن يكون رقم موجب.', 'error');
            element.innerHTML = originalContent;
            element.classList.remove('editing');
            return;
        }
        
        // Update the processed data
        if (processorInstance && processorInstance.processedData[index]) {
            processorInstance.processedData[index].newPrice = newPrice;
            
            // Update display
            const displayPrice = processorInstance.formatPrice(newPrice);
            element.innerHTML = displayPrice;
            
            // Update onclick attribute
            element.setAttribute('onclick', `handlePriceClick(this, '${displayPrice}', ${index})`);
            element.setAttribute('data-original-price', displayPrice);
            
            element.classList.remove('editing');
            showCopyNotification('✅ تم تحديث السعر: ' + displayPrice + ' جنيه');
        }
    };
    
    // Handle cancel on Escape
    const cancelEdit = () => {
        element.innerHTML = originalContent;
        element.classList.remove('editing');
        showCopyNotification('❌ تم إلغاء التعديل', 'warning');
    };
    
    input.addEventListener('keydown', (e) => {
        e.stopPropagation();
        if (e.key === 'Enter') {
            saveEdit();
        } else if (e.key === 'Escape') {
            cancelEdit();
        }
    });
    
    input.addEventListener('blur', saveEdit);
    
    // Prevent click propagation
    input.addEventListener('click', (e) => {
        e.stopPropagation();
    });
    
    showCopyNotification('🔄 وضع التعديل مُفعل - اكتب السعر الجديد', 'info');
}

// Update price in processed data
function updateProcessedData(index, newPrice) {
    if (processorInstance && processorInstance.processedData[index]) {
        processorInstance.processedData[index].newPrice = newPrice;
        return true;
    }
    return false;
}

// Show copy notification with icons
function showCopyNotification(message, type = 'success') {
    const notification = document.getElementById('notification');
    
    // Add appropriate icon based on type
    let icon = '✅';
    if (type === 'error') icon = '❌';
    else if (type === 'warning') icon = '⚠️';
    else if (type === 'info') icon = 'ℹ️';
    
    notification.innerHTML = `<span style="margin-left: 8px;">${icon}</span>${message}`;
    notification.className = `notification ${type} show`;
    
    setTimeout(() => {
        notification.classList.remove('show');
    }, type === 'warning' ? 2500 : 2000);
}

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    processorInstance = new ExcelProcessor();
});

// Add visual feedback for copy action - permanent until page refresh
function addCopyVisualFeedback(element) {
    // Add permanent copy indicator
    element.classList.add('copied');
    
    // Create temporary indicator
    const indicator = document.createElement('span');
    indicator.className = 'copy-indicator';
    indicator.innerHTML = '✅';
    indicator.style.cssText = `
        position: absolute;
        top: -8px;
        right: -8px;
        background: var(--success-color);
        color: white;
        width: 20px;
        height: 20px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 10px;
        z-index: 100;
        animation: copyPulse 0.6s ease-out;
    `;
    
    // Add indicator to element
    element.style.position = 'relative';
    element.appendChild(indicator);
    
    // Remove indicator after animation but keep the 'copied' class permanently
    setTimeout(() => {
        if (indicator.parentNode) {
            indicator.parentNode.removeChild(indicator);
        }
    }, 1000);
    
    // DO NOT remove the 'copied' class - keep it permanent until page refresh
}

// Global download function
function downloadExcel() {
    if (processorInstance) {
        processorInstance.exportToExcel();
    }
}