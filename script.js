let excelData = null;
let customFontNonBoldName = null;
let customFontBoldName = null;
let backgroundImageURL = null; // Store background image URL

// Handle Excel file input and parse
function handleFile(e) {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        console.log('Excel data loaded', excelData); // Debugging log
    };

    reader.readAsArrayBuffer(file);
}

// Handle custom font upload for non-bold text
function handleCustomFontNonBoldUpload(e) {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function(event) {
        customFontNonBoldName = file.name.split('.')[0];
        const fontFace = new FontFace(customFontNonBoldName, `url(${event.target.result})`);
        fontFace.load().then(function(loadedFont) {
            document.fonts.add(loadedFont);
            console.log(`Non-Bold Font ${customFontNonBoldName} loaded successfully`);
        });
    };

    reader.readAsDataURL(file);
}

// Handle custom font upload for bold text
function handleCustomFontBoldUpload(e) {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function(event) {
        customFontBoldName = file.name.split('.')[0];
        const fontFace = new FontFace(customFontBoldName, `url(${event.target.result})`);
        fontFace.load().then(function(loadedFont) {
            document.fonts.add(loadedFont);
            console.log(`Bold Font ${customFontBoldName} loaded successfully`);
        });
    };

    reader.readAsDataURL(file);
}

// Handle background image upload
function handleBackgroundImageUpload(e) {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function(event) {
        backgroundImageURL = event.target.result; // Store the image URL
        console.log("Background image loaded successfully");
    };

    reader.readAsDataURL(file);
}

// Function to clean up text to be used as a valid file name
function cleanFileName(text) {
    return text.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 50); // Replace invalid characters and limit length
}

// Function to render and preview the text
function runTextFormatting(isDownload = false) {
    if (!excelData) {
        alert("Please upload an Excel file first.");
        return;
    }

    const outputArea = document.getElementById('outputArea');
    outputArea.innerHTML = ''; // Clear previous output

    // Get selected fonts, sizes, colors, text alignment, and background color
    const nonBoldFont = customFontNonBoldName || document.getElementById('nonBoldFont').value;
    const boldFont = customFontBoldName || document.getElementById('boldFont').value;
    const nonBoldColor = document.getElementById('nonBoldColor').value;
    const boldColor = document.getElementById('boldColor').value;
    const nonBoldSize = document.getElementById('nonBoldSize').value;
    const boldSize = document.getElementById('boldSize').value;
    const backgroundColor = document.getElementById('backgroundColor').value;

    const textAlign = document.querySelector('input[name="textAlign"]:checked').value;
    const centerHorizontally = document.getElementById('centerHorizontally').checked;
    const centerVertically = document.getElementById('centerVertically').checked;
    const boldToggle = document.getElementById('boldToggle').checked; // Check if bold word should be bolded

    // Check for transparent background
    const transparentBackground = document.getElementById('transparentBackground').checked;

    // Get frame size and fit options
    const frameWidthInput = document.getElementById('frameWidth').value;
    const frameHeightInput = document.getElementById('frameHeight').value;
    const fitToFrame = document.getElementById('fitToFrame').checked;

    const zip = new JSZip();

    excelData.forEach((row, index) => {
        let sentence = row[0]; // Get the sentence from the row

        // Create a div to display the sentence
        const textDiv = document.createElement('div');
        textDiv.style.fontFamily = nonBoldFont;
        textDiv.style.color = nonBoldColor;
        textDiv.style.fontSize = `${nonBoldSize}px`;

        // Handle horizontal centering
        if (centerHorizontally) {
            textDiv.style.textAlign = 'center';
        } else {
            textDiv.style.textAlign = textAlign; // Use the chosen alignment (left, right, center)
        }

        // Set fixed dimensions if input is provided, else auto-size
        let frameWidth = frameWidthInput ? `${frameWidthInput}px` : 'auto';
        let frameHeight = frameHeightInput ? `${frameHeightInput}px` : 'auto';

        // Handle background image, color, or transparency
        if (transparentBackground) {
            textDiv.style.backgroundColor = 'transparent'; // Apply transparent background
        } else if (backgroundImageURL) {
            textDiv.style.backgroundImage = `url(${backgroundImageURL})`;
            textDiv.style.backgroundPosition = 'center center'; // Center the background image both horizontally and vertically
            textDiv.style.backgroundRepeat = 'no-repeat'; // Ensure the image doesn't repeat
            textDiv.style.backgroundSize = 'cover'; // Scale the image to cover the frame
        } else {
            textDiv.style.backgroundColor = backgroundColor; // Apply selected background color
        }

        textDiv.style.padding = '20px'; // Add padding to ensure the text isn't cramped
        textDiv.style.marginBottom = '10px'; // Add margin for spacing between sentences

        // Apply vertical centering if enabled
        if (centerVertically) {
            textDiv.style.display = 'flex';
            textDiv.style.flexDirection = 'column'; // Keep each part of the sentence on a separate line
            textDiv.style.alignItems = centerHorizontally ? 'center' : 'flex-start'; // Align horizontally if selected
            textDiv.style.justifyContent = 'center'; // Vertically center the content
        } else {
            textDiv.style.display = 'block'; // Default to block-level display if vertical centering is disabled
        }

        // Split the sentence by the bold markers (*word*) and ensure the bold word is moved to a new line
        let parts = sentence.split('*');
        let formattedSentence = '';

        if (parts.length === 3) { // We expect a sentence with one bold word between two normal text parts
            formattedSentence = `<div>${parts[0]}</div>`;  // Text before the bold word (block element)
            
            // Apply bold styling only if the boldToggle checkbox is checked
            if (boldToggle) {
                formattedSentence += `<div style="font-weight: bold; font-family: ${boldFont}; font-size: ${boldSize}px; color: ${boldColor};">${parts[1]}</div>`;  // Bold word on its own line (block element)
            } else {
                formattedSentence += `<div style="font-family: ${boldFont}; font-size: ${boldSize}px; color: ${boldColor};">${parts[1]}</div>`;  // Non-bold word on its own line
            }

            formattedSentence += `<div>${parts[2]}</div>`;  // Text after the bold word (block element)
        } else {
            // If no bold word is found, just use the sentence as it is
            formattedSentence = `<div>${sentence}</div>`;
        }

        textDiv.innerHTML = formattedSentence; // Add the sentence to the div, bold word on a new line

        // If user inputs frame dimensions, apply them, otherwise auto-size to content
        if (fitToFrame && frameWidthInput && frameHeightInput) {
            textDiv.style.width = `${frameWidthInput}px`;
            textDiv.style.height = `${frameHeightInput}px`;
            textDiv.style.overflow = 'hidden'; // Ensure content does not overflow
        } else {
            textDiv.style.width = frameWidth;
            textDiv.style.height = frameHeight;
        }

        outputArea.appendChild(textDiv); // Add the sentence to the output area

        // If downloading, capture the image
        if (isDownload) {
            html2canvas(textDiv, { backgroundColor: null }).then(canvas => { // Set backgroundColor to null for transparency
                canvas.toBlob(function(blob) {
                    // Generate a cleaned-up file name based on the text content
                    const cleanFileNameText = cleanFileName(sentence);
                    const imageName = `${cleanFileNameText}.png`; // Save as PNG to support transparency
                    zip.file(imageName, blob); // Add image to ZIP

                    if (index === excelData.length - 1) {
                        zip.generateAsync({ type: 'blob' }).then(function(content) {
                            saveAs(content, 'formatted_texts.zip');
                        });
                    }
                });
            });
        }
    });
}

// Event listeners for file input and buttons
document.getElementById('fileInput').addEventListener('change', handleFile);
document.getElementById('customFontNonBoldUpload').addEventListener('change', handleCustomFontNonBoldUpload);
document.getElementById('customFontBoldUpload').addEventListener('change', handleCustomFontBoldUpload);
document.getElementById('backgroundImageUpload').addEventListener('change', handleBackgroundImageUpload);

// Preview button event listener
document.getElementById('previewButton').addEventListener('click', function() {
    runTextFormatting(false); // Just preview
});

// Download ZIP button event listener
document.getElementById('downloadZipButton').addEventListener('click', function() {
    runTextFormatting(true); // Download ZIP
});
