'use strict';

(function () {
    Office.onReady(function (info) {
        // Office is ready with information about the host
        $(document).ready(function () {
            // The document is ready

            // By default, hide header/footer buttons - we'll show them only for Word
            $('#insertSelectedDisclaimersAsHeader').hide();
            $('#insertSelectedDisclaimersAsFooter').hide();
            
            // Initialize the Rich Text toggle as disabled by default
            $('#useRichTextContainer').hide();
            
            // Hide the "Insert at Bottom" button by default - only show for Outlook
            $('#insertSelectedDisclaimersAtEnd').hide();

            // Check if the host is Word (which supports headers and footers)
            if (info.host === Office.HostType.Word) {
                // It's Word, so show the header and footer buttons
                $('#insertSelectedDisclaimersAsHeader').show();
                $('#insertSelectedDisclaimersAsFooter').show();
                $('#useRichTextContainer').show();
            } 
            // PowerPoint and Outlook also support rich text
            else if (info.host === Office.HostType.Outlook) {
                $('#useRichTextContainer').show();
                // Show "Insert at Bottom" button only for Outlook
                $('#insertSelectedDisclaimersAtEnd').show();
            } else if (info.host === Office.HostType.PowerPoint) {
                $('#useRichTextContainer').hide(); // Hide for PowerPoint
            }

            // Automatically load disclaimers from SharePoint list when page loads
            fetchDisclaimersSharePointList();

            // Add event listeners for insertion buttons
            $('#insertSelectedDisclaimersAtCursor').on('click', function () {
                insertSelectedDisclaimersAtCursor();
            });

            $('#insertSelectedDisclaimersAsHeader').on('click', function () {
                insertSelectedDisclaimersAsHeader();
            });

            $('#insertSelectedDisclaimersAsFooter').on('click', function () {
                insertSelectedDisclaimersAsFooter();
            });

            $('#insertSelectedDisclaimersAtEnd').on('click', function () {
                insertSelectedDisclaimersAtEnd();
            });
        });
    });

    // Fetch disclaimers from a Web API
    function fetchDisclaimers() {
        // Get base URL - in production this should be configurable
        const baseUrl = getBaseUrl();
        const apiUrl = `${baseUrl}/api/disclaimer`;

        try {
            fetch(apiUrl)
                .then(response => {
                    // Check if response is ok (status 200-299)
                    if (!response.ok) {
                        // Try to parse error message as JSON
                        return response.json()
                            .then(errorData => {
                                throw new Error(errorData.error || `Server error: ${response.status}`);
                            })
                            .catch(jsonError => {
                                // If error response isn't valid JSON
                                return response.text().then(errorText => {
                                    throw new Error(`Server error (${response.status}): ${errorText.slice(0, 100)}`);
                                });
                            });
                    }
                    return response.json();
                })
                .then(data => {
                    let disclaimers = data;
                    let disclaimerList = $('#disclaimer-list');
                    disclaimerList.empty();

                    // Create a <ul> element to hold the list of disclaimers with checkboxes
                    let list = $('<ul style="list-style-type: none;"></ul>');
                    disclaimerList.append(list);


                    disclaimers.forEach(disclaimer => {
                        // Create a list item for each disclaimer
                        let listItem = $('<li></li>');

                        // Create a checkbox and label for the disclaimer
                        let checkbox = $(`<input type="checkbox" id="disclaimer_${disclaimer.id}" value="${disclaimer.text}">`);
                        let label = $(`<label for="disclaimer_${disclaimer.id}">${disclaimer.description}</label>`);

                        // Append the checkbox and label to the list item
                        listItem.append(checkbox);
                        listItem.append(label);

                        // Append the list item to the list
                        list.append(listItem);
                    });
                })
                .catch(error => {
                    console.log('Error:', error);
                    showMessage(`Failed to load disclaimers from Web API: ${error.message}`);
                });
        } catch (error) {
            console.log('Error:', error);
            showMessage(`Failed to load disclaimers from Web API: ${error.message}`);
        }
    }

    // Fetch disclaimers from a SharePoint list
    async function fetchDisclaimersSharePointList() {
        // Get base URL - in production this should be configurable
        const baseUrl = getBaseUrl();
        const apiUrl = `${baseUrl}/api/sharepointlist`;

        try {
            const response = await fetch(apiUrl);
            
            // Check if response is ok (status 200-299)
            if (!response.ok) {
                // Try to parse error message as JSON
                try {
                    const errorData = await response.json();
                    throw new Error(errorData.error || `Server error: ${response.status}`);
                } catch (jsonError) {
                    // If error response isn't valid JSON
                    const errorText = await response.text();
                    throw new Error(`Server error (${response.status}): ${errorText.slice(0, 100)}`);
                }
            }
            
            const data = await response.json();
            let disclaimers = data;
            let disclaimerList = $('#disclaimer-list');
            disclaimerList.empty();

            // Create a <ul> element to hold the list of disclaimers with checkboxes
            let list = $('<ul style="list-style-type: none;"></ul>');
            disclaimerList.append(list);

            disclaimers.forEach(disclaimer => {
                // Create a list item for each disclaimer
                let listItem = $('<li></li>');

                // Create a checkbox and label for the disclaimer
                let checkbox = $(`<input type="checkbox" id="disclaimer_${disclaimer.id}" value="${disclaimer.text}">`);
                
                // Store the RichText content as a data attribute (if available)
                if (disclaimer.richText) {
                    checkbox.attr('data-richtext', disclaimer.richText);
                }
                
                let label = $(`<label for="disclaimer_${disclaimer.id}">${disclaimer.description}</label>`);

                // Add click event to label to toggle the checkbox
                label.on('click', function(event) {
                    event.preventDefault(); // Prevent default label behavior
                    checkbox.prop('checked', !checkbox.prop('checked'));
                });

                // Append the checkbox and label to the list item
                listItem.append(checkbox);
                listItem.append(label);

                // Append the list item to the list
                list.append(listItem);
            });
        }
        catch (error) {
            console.log('Error:', error);
            showMessage(`Failed to load disclaimers from SharePoint List: ${error.message}`);
        }
    }

    // Helper function to determine the base API URL
    function getBaseUrl() {
        // For production, the URLs should be configurable through environment variables
        // or other configuration mechanisms
        if (window.location.hostname === 'localhost') {
            return 'https://localhost:7057';
        }

        // For production - ideally this would come from configuration
        return 'https://mycompanydisclaimers.azurewebsites.us';
    }

    function showMessage(text) {
        const appendedText = $('#disclaimer-list').html() + text + "<br>---";
        $('#disclaimer-list').html(appendedText);
    }

    // A simple way to determine the Office app context of the running Add-in, you can check the value of Office.context.host.
    // Office.HostType.Outlook is the host type for Outlook.
    // Office.HostType.Word is the host type for Word.
    // Office.HostType.Excel is the host type for Excel.
    // Office.HostType.PowerPoint is the host type for PowerPoint.

    function insertText(text, isRichText = false) {
        // Check if the host is Outlook
        if (Office.context.mailbox) {
            // Use the setSelectedDataAsync method on the body object to insert text
            Office.context.mailbox.item.body.setSelectedDataAsync(text,
                { coercionType: isRichText ? Office.CoercionType.Html : Office.CoercionType.Text },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                    }
                }
            );
        } else {
            // For Word and PowerPoint, we can use HTML, but Excel only supports Text
            const isExcel = Office.context.host === Office.HostType.Excel;
            const finalCoercionType = isRichText && !isExcel ? Office.CoercionType.Html : Office.CoercionType.Text;
            
            // Fallback for other Office applications
            Office.context.document.setSelectedDataAsync(text,
                { coercionType: finalCoercionType },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                    }
                }
            );
        }
    }

    function insertSelectedDisclaimersAtCursor() {
        try {
            const selectedCheckboxes = document.querySelectorAll('#disclaimer-list input[type="checkbox"]:checked');
            const useRichText = document.getElementById('useRichText').checked;

            if (selectedCheckboxes.length === 0) {
                console.log("Please select at least one disclaimer to insert.");
                return;
            }

            const selectedTexts = [];
            const selectedRichTexts = [];
            
            selectedCheckboxes.forEach(checkbox => {
                if (checkbox.value) {
                    selectedTexts.push(checkbox.value);
                }
                
                if (useRichText && checkbox.dataset.richtext) {
                    selectedRichTexts.push(checkbox.dataset.richtext);
                }
            });

            const textToInsert = selectedTexts.join('\n\n');
            const richTextToInsert = selectedRichTexts.join('<br><br>');

            // Check if we're using rich text and have rich text content available
            const shouldUseRichText = useRichText && selectedRichTexts.length > 0;
            
            // Check if the host is Outlook
            if (Office.context.mailbox) {
                // Use the setSelectedDataAsync method on the body object for Outlook
                Office.context.mailbox.item.body.setSelectedDataAsync(
                    shouldUseRichText ? richTextToInsert : textToInsert,
                    { coercionType: shouldUseRichText ? Office.CoercionType.Html : Office.CoercionType.Text },
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log("Disclaimers inserted successfully in Outlook!");
                        } else {
                            console.log("Error inserting disclaimers in Outlook: " + asyncResult.error.message);
                        }
                    }
                );
            } else {
                // For Word and PowerPoint, we can use HTML, but Excel only supports Text
                const isExcel = Office.context.host === Office.HostType.Excel;
                
                // Fallback for other Office applications (Word, Excel, PowerPoint)
                Office.context.document.setSelectedDataAsync(
                    shouldUseRichText && !isExcel ? richTextToInsert : textToInsert,
                    { coercionType: shouldUseRichText && !isExcel ? Office.CoercionType.Html : Office.CoercionType.Text },
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log("Disclaimers inserted successfully!");
                        } else {
                            console.log("Error inserting disclaimers: " + asyncResult.error.message);
                        }
                    }
                );
            }
        } catch (error) {
            console.log("Error: " + error.message);
        }
    }

    function insertSelectedDisclaimersAsHeader() {
        // Get array of selected disclaimer texts
        let selectedDisclaimers = [];
        let selectedRichTextDisclaimers = [];
        const useRichText = document.getElementById('useRichText').checked;
        
        $('#disclaimer-list input[type="checkbox"]:checked').each(function () {
            selectedDisclaimers.push($(this).val());
            
            // Check if rich text is available and enabled
            if (useRichText && $(this).data('richtext')) {
                selectedRichTextDisclaimers.push($(this).data('richtext'));
            }
        });

        if (selectedDisclaimers.length > 0) {
            // First check if there's existing content in the header
            Word.run(function (context) {
                const header = context.document.sections.getFirst().getHeader("primary");
                header.load('text');
                
                return context.sync()
                    .then(function () {
                        const hasExistingText = header.text.trim() !== '';
                        
                        try {
                            // Show alignment options dialog using direct HTML manipulation
                            // This method doesn't rely on jQuery UI dialog
                            const overlayId = 'alignment-overlay';
                            const dialogId = 'alignment-dialog';
                            
                            // Remove any existing overlay/dialog
                            $('#' + overlayId).remove();
                            $('#' + dialogId).remove();
                            
                            // Create overlay
                            const overlay = $('<div id="' + overlayId + '" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000;"></div>');
                            
                            // Create dialog
                            const dialog = $('<div id="' + dialogId + '" style="position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 15px; border-radius: 5px; width: 300px; z-index: 1001;"></div>');
                            
                            // Add header
                            dialog.append('<h3 style="margin-top: 0; font-size: 14px;">Select Header Alignment</h3>');
                            
                            // Add warning if needed
                            if (hasExistingText) {
                                dialog.append(
                                    '<div style="margin-bottom: 15px; color: #d83b01; padding: 8px; background-color: #fff4ce; border-left: 4px solid #d83b01;">' +
                                    '<strong>Warning:</strong> There is existing text in the header that will be replaced.' +
                                    '</div>'
                                );
                            }
                            
                            // Add alignment options
                            const alignmentContainer = $('<div style="display: flex; justify-content: space-between; margin-bottom: 10px;"></div>');
                            
                            // Define alignment options
                            const alignmentOptions = [
                                { value: 'left', display: 'Left' },
                                { value: 'center', display: 'Center' },
                                { value: 'right', display: 'Right' }
                            ];
                            
                            // Create option buttons
                            alignmentOptions.forEach(option => {
                                const optionButton = $(
                                    `<div class="alignment-option" data-value="${option.value}" style="text-align: center; cursor: pointer; padding: 8px; border: 1px solid #ccc; flex: 1; margin: 0 3px; font-size: 12px;">
                                        <div style="text-align: ${option.value}; border-bottom: 1px solid #ccc; padding-bottom: 3px; margin-bottom: 3px;">
                                            <span>Text</span>
                                        </div>
                                        ${option.display}
                                    </div>`
                                );
                                alignmentContainer.append(optionButton);
                            });
                            
                            dialog.append(alignmentContainer);
                            
                            // Add buttons
                            const buttonContainer = $('<div style="margin-top: 8px; text-align: center;"></div>');
                            const confirmButton = $('<button id="confirmHeader" style="margin-right: 8px; padding: 6px 12px; cursor: pointer; font-size: 12px;">Insert</button>');
                            const cancelButton = $('<button id="cancelHeader" style="padding: 6px 12px; cursor: pointer; font-size: 12px;">Cancel</button>');
                            
                            buttonContainer.append(confirmButton);
                            buttonContainer.append(cancelButton);
                            dialog.append(buttonContainer);
                            
                            // Add to document
                            $('body').append(overlay);
                            $('body').append(dialog);
                            
                            // Set default selection
                            let selectedAlignment = 'left';
                            $('.alignment-option[data-value="left"]').css('background-color', '#e6f2ff');
                            
                            // Handle alignment options
                            $('.alignment-option').on('click', function() {
                                $('.alignment-option').css('background-color', '');
                                $(this).css('background-color', '#e6f2ff');
                                selectedAlignment = $(this).data('value');
                            });

                            // Handle confirmation
                            $('#confirmHeader').on('click', function() {
                                // Remove dialog elements
                                $('#' + overlayId).remove();
                                $('#' + dialogId).remove();

                                // Insert text with selected alignment
                                Word.run(function (context) {
                                    const header = context.document.sections.getFirst().getHeader("primary");
                                    header.clear();
                                    
                                    // Check if we should use rich text
                                    const shouldUseRichText = useRichText && selectedRichTextDisclaimers.length > 0;
                                    const disclaimersToUse = shouldUseRichText ? selectedRichTextDisclaimers : selectedDisclaimers;
                                    
                                    if (shouldUseRichText) {
                                        // For rich text, we need to use insertHtml and then set alignment
                                        const combinedHtml = disclaimersToUse.join('<br><br>');
                                        const htmlWithAlignment = `<div style="text-align: ${selectedAlignment};">${combinedHtml}</div>`;
                                        header.insertHtml(htmlWithAlignment, Word.InsertLocation.end);
                                        
                                        // Force alignment setting for rich text after HTML insertion
                                        const paragraphs = header.paragraphs;
                                        paragraphs.load("items");
                                        return context.sync().then(function() {
                                            for (let i = 0; i < paragraphs.items.length; i++) {
                                                if (selectedAlignment === 'left') {
                                                    paragraphs.items[i].alignment = Word.Alignment.left;
                                                } else if (selectedAlignment === 'center') {
                                                    paragraphs.items[i].alignment = Word.Alignment.centered;
                                                } else if (selectedAlignment === 'right') {
                                                    paragraphs.items[i].alignment = Word.Alignment.right;
                                                }
                                            }
                                            return context.sync().then(function() {
                                                console.log(`Header added with ${selectedAlignment} alignment.`);
                                            });
                                        });
                                    } else {
                                        // Insert each disclaimer as separate paragraph with the same alignment for plain text
                                        for (let i = 0; i < disclaimersToUse.length; i++) {
                                            const paragraph = header.insertParagraph(disclaimersToUse[i], Word.InsertLocation.end);
                                            
                                            if (selectedAlignment === 'left') {
                                                paragraph.alignment = Word.Alignment.left;
                                            } else if (selectedAlignment === 'center') {
                                                paragraph.alignment = Word.Alignment.centered;
                                            } else if (selectedAlignment === 'right') {
                                                paragraph.alignment = Word.Alignment.right;
                                            }
                                        }
                                    }
                                    
                                    return context.sync()
                                        .then(function() {
                                            console.log(`Header added with ${selectedAlignment} alignment.`);
                                        });
                                }).catch(function (error) {
                                    console.log('Error:', error);
                                });
                            });
                            
                            // Handle cancellation
                            $('#cancelHeader, #' + overlayId).on('click', function() {
                                $('#' + overlayId).remove();
                                $('#' + dialogId).remove();
                            });
                            
                            // Prevent clicks on the dialog from closing it
                            $('#' + dialogId).on('click', function(e) {
                                e.stopPropagation();
                            });
                            
                        } catch (err) {
                            console.log('Error showing dialog:', err);
                            
                            // Fallback to simple confirm
                            if (hasExistingText) {
                                if (!confirm('There is existing text in the header. Replace with selected disclaimers?')) {
                                    return;
                                }
                            }
                            
                            // Use left alignment by default
                            Word.run(function (context) {
                                const header = context.document.sections.getFirst().getHeader("primary");
                                header.clear();
                                
                                // Check if we should use rich text
                                const shouldUseRichText = useRichText && selectedRichTextDisclaimers.length > 0;
                                const disclaimersToUse = shouldUseRichText ? selectedRichTextDisclaimers : selectedDisclaimers;
                                
                                if (shouldUseRichText) {
                                    // For rich text, we need to use insertHtml with left alignment
                                    const combinedHtml = disclaimersToUse.join('<br><br>');
                                    const htmlWithAlignment = `<div style="text-align: left;">${combinedHtml}</div>`;
                                    header.insertHtml(htmlWithAlignment, Word.InsertLocation.end);
                                    
                                    // Force alignment setting for rich text after HTML insertion
                                    const paragraphs = header.paragraphs;
                                    paragraphs.load("items");
                                    return context.sync().then(function() {
                                        for (let i = 0; i < paragraphs.items.length; i++) {
                                            paragraphs.items[i].alignment = Word.Alignment.left;
                                        }
                                        return context.sync().then(function() {
                                            console.log('Header added with default alignment.');
                                        });
                                    });
                                } else {
                                    // Insert each disclaimer as separate paragraph with left alignment for plain text
                                    for (let i = 0; i < disclaimersToUse.length; i++) {
                                        const paragraph = header.insertParagraph(disclaimersToUse[i], Word.InsertLocation.end);
                                        paragraph.alignment = Word.Alignment.left;
                                    }
                                }
                                
                                return context.sync()
                                    .then(function() {
                                        console.log('Header added with default alignment.');
                                    });
                            }).catch(function (error) {
                                console.log('Error:', error);
                            });
                        }
                    });
            });
        } else {
            console.log('No disclaimers selected.');
        }
    }

    function insertSelectedDisclaimersAsFooter() {
        // Get array of selected disclaimer texts
        let selectedDisclaimers = [];
        let selectedRichTextDisclaimers = [];
        const useRichText = document.getElementById('useRichText').checked;
        
        $('#disclaimer-list input[type="checkbox"]:checked').each(function () {
            selectedDisclaimers.push($(this).val());
            
            // Check if rich text is available and enabled
            if (useRichText && $(this).data('richtext')) {
                selectedRichTextDisclaimers.push($(this).data('richtext'));
            }
        });

        if (selectedDisclaimers.length > 0) {
            // First check if there's existing content in the footer
            Word.run(function (context) {
                const footer = context.document.sections.getFirst().getFooter("primary");
                footer.load('text');
                
                return context.sync()
                    .then(function () {
                        const hasExistingText = footer.text.trim() !== '';
                        
                        try {
                            // Show alignment options dialog using direct HTML manipulation
                            // This method doesn't rely on jQuery UI dialog
                            const overlayId = 'alignment-overlay';
                            const dialogId = 'alignment-dialog';
                            
                            // Remove any existing overlay/dialog
                            $('#' + overlayId).remove();
                            $('#' + dialogId).remove();
                            
                            // Create overlay
                            const overlay = $('<div id="' + overlayId + '" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000;"></div>');
                            
                            // Create dialog
                            const dialog = $('<div id="' + dialogId + '" style="position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 15px; border-radius: 5px; width: 300px; z-index: 1001;"></div>');
                            
                            // Add header
                            dialog.append('<h3 style="margin-top: 0; font-size: 14px;">Select Footer Alignment</h3>');
                            
                            // Add warning if needed
                            if (hasExistingText) {
                                dialog.append(
                                    '<div style="margin-bottom: 15px; color: #d83b01; padding: 8px; background-color: #fff4ce; border-left: 4px solid #d83b01;">' +
                                    '<strong>Warning:</strong> There is existing text in the footer that will be replaced.' +
                                    '</div>'
                                );
                            }
                            
                            // Add alignment options
                            const alignmentContainer = $('<div style="display: flex; justify-content: space-between; margin-bottom: 10px;"></div>');
                            
                            // Define alignment options
                            const alignmentOptions = [
                                { value: 'left', display: 'Left' },
                                { value: 'center', display: 'Center' },
                                { value: 'right', display: 'Right' }
                            ];
                            
                            // Create option buttons
                            alignmentOptions.forEach(option => {
                                const optionButton = $(
                                    `<div class="alignment-option" data-value="${option.value}" style="text-align: center; cursor: pointer; padding: 8px; border: 1px solid #ccc; flex: 1; margin: 0 3px; font-size: 12px;">
                                        <div style="text-align: ${option.value}; border-bottom: 1px solid #ccc; padding-bottom: 3px; margin-bottom: 3px;">
                                            <span>Text</span>
                                        </div>
                                        ${option.display}
                                    </div>`
                                );
                                alignmentContainer.append(optionButton);
                            });
                            
                            dialog.append(alignmentContainer);
                            
                            // Add buttons
                            const buttonContainer = $('<div style="margin-top: 8px; text-align: center;"></div>');
                            const confirmButton = $('<button id="confirmFooter" style="margin-right: 8px; padding: 6px 12px; cursor: pointer; font-size: 12px;">Insert</button>');
                            const cancelButton = $('<button id="cancelFooter" style="padding: 6px 12px; cursor: pointer; font-size: 12px;">Cancel</button>');
                            
                            buttonContainer.append(confirmButton);
                            buttonContainer.append(cancelButton);
                            dialog.append(buttonContainer);
                            
                            // Add to document
                            $('body').append(overlay);
                            $('body').append(dialog);
                            
                            // Set default selection
                            let selectedAlignment = 'left';
                            $('.alignment-option[data-value="left"]').css('background-color', '#e6f2ff');
                            
                            // Handle alignment options
                            $('.alignment-option').on('click', function() {
                                $('.alignment-option').css('background-color', '');
                                $(this).css('background-color', '#e6f2ff');
                                selectedAlignment = $(this).data('value');
                            });

                            // Handle confirmation
                            $('#confirmFooter').on('click', function() {
                                // Remove dialog elements
                                $('#' + overlayId).remove();
                                $('#' + dialogId).remove();

                                // Insert text with selected alignment
                                Word.run(function (context) {
                                    const footer = context.document.sections.getFirst().getFooter("primary");
                                    footer.clear();
                                    
                                    // Check if we should use rich text
                                    const shouldUseRichText = useRichText && selectedRichTextDisclaimers.length > 0;
                                    const disclaimersToUse = shouldUseRichText ? selectedRichTextDisclaimers : selectedDisclaimers;
                                    
                                    if (shouldUseRichText) {
                                        // For rich text, we need to use insertHtml and then set alignment
                                        const combinedHtml = disclaimersToUse.join('<br><br>');
                                        const htmlWithAlignment = `<div style="text-align: ${selectedAlignment};">${combinedHtml}</div>`;
                                        footer.insertHtml(htmlWithAlignment, Word.InsertLocation.end);
                                        
                                        // Force alignment setting for rich text after HTML insertion
                                        const paragraphs = footer.paragraphs;
                                        paragraphs.load("items");
                                        return context.sync().then(function() {
                                            for (let i = 0; i < paragraphs.items.length; i++) {
                                                if (selectedAlignment === 'left') {
                                                    paragraphs.items[i].alignment = Word.Alignment.left;
                                                } else if (selectedAlignment === 'center') {
                                                    paragraphs.items[i].alignment = Word.Alignment.centered;
                                                } else if (selectedAlignment === 'right') {
                                                    paragraphs.items[i].alignment = Word.Alignment.right;
                                                }
                                            }
                                            return context.sync().then(function() {
                                                console.log(`Footer added with ${selectedAlignment} alignment.`);
                                            });
                                        });
                                    } else {
                                        // Insert each disclaimer as separate paragraph with the same alignment for plain text
                                        for (let i = 0; i < disclaimersToUse.length; i++) {
                                            const paragraph = footer.insertParagraph(disclaimersToUse[i], Word.InsertLocation.end);
                                            
                                            if (selectedAlignment === 'left') {
                                                paragraph.alignment = Word.Alignment.left;
                                            } else if (selectedAlignment === 'center') {
                                                paragraph.alignment = Word.Alignment.centered;
                                            } else if (selectedAlignment === 'right') {
                                                paragraph.alignment = Word.Alignment.right;
                                            }
                                        }
                                    }
                                    
                                    return context.sync()
                                        .then(function() {
                                            console.log(`Footer added with ${selectedAlignment} alignment.`);
                                        });
                                }).catch(function (error) {
                                    console.log('Error:', error);
                                });
                            });
                            
                            // Handle cancellation
                            $('#cancelFooter, #' + overlayId).on('click', function() {
                                $('#' + overlayId).remove();
                                $('#' + dialogId).remove();
                            });
                            
                            // Prevent clicks on the dialog from closing it
                            $('#' + dialogId).on('click', function(e) {
                                e.stopPropagation();
                            });
                            
                        } catch (err) {
                            console.log('Error showing dialog:', err);
                            
                            // Fallback to simple confirm
                            if (hasExistingText) {
                                if (!confirm('There is existing text in the footer. Replace with selected disclaimers?')) {
                                    return;
                                }
                            }
                            
                            // Use left alignment by default
                            Word.run(function (context) {
                                const footer = context.document.sections.getFirst().getFooter("primary");
                                footer.clear();
                                
                                // Check if we should use rich text
                                const shouldUseRichText = useRichText && selectedRichTextDisclaimers.length > 0;
                                const disclaimersToUse = shouldUseRichText ? selectedRichTextDisclaimers : selectedDisclaimers;
                                
                                if (shouldUseRichText) {
                                    // For rich text, we need to use insertHtml with left alignment
                                    const combinedHtml = disclaimersToUse.join('<br><br>');
                                    const htmlWithAlignment = `<div style="text-align: left;">${combinedHtml}</div>`;
                                    footer.insertHtml(htmlWithAlignment, Word.InsertLocation.end);
                                    
                                    // Force alignment setting for rich text after HTML insertion
                                    const paragraphs = footer.paragraphs;
                                    paragraphs.load("items");
                                    return context.sync().then(function() {
                                        for (let i = 0; i < paragraphs.items.length; i++) {
                                            paragraphs.items[i].alignment = Word.Alignment.left;
                                        }
                                        return context.sync().then(function() {
                                            console.log('Footer added with default alignment.');
                                        });
                                    });
                                } else {
                                    // Insert each disclaimer as separate paragraph with left alignment for plain text
                                    for (let i = 0; i < disclaimersToUse.length; i++) {
                                        const paragraph = footer.insertParagraph(disclaimersToUse[i], Word.InsertLocation.end);
                                        paragraph.alignment = Word.Alignment.left;
                                    }
                                }
                                
                                return context.sync()
                                    .then(function() {
                                        console.log('Footer added with default alignment.');
                                    });
                            }).catch(function (error) {
                                console.log('Error:', error);
                            });
                        }
                    });
            });
        } else {
            console.log('No disclaimers selected.');
        }
    }

    function insertSelectedDisclaimersAtEnd() {
        try {
            const selectedCheckboxes = document.querySelectorAll('#disclaimer-list input[type="checkbox"]:checked');
            const useRichText = document.getElementById('useRichText') && document.getElementById('useRichText').checked;

            if (selectedCheckboxes.length === 0) {
                console.log("Please select at least one disclaimer to insert.");
                return;
            }

            const selectedTexts = [];
            const selectedRichTexts = [];
            
            selectedCheckboxes.forEach(checkbox => {
                if (checkbox.value) {
                    selectedTexts.push(checkbox.value);
                }
                
                if (useRichText && checkbox.dataset.richtext) {
                    selectedRichTexts.push(checkbox.dataset.richtext);
                }
            });

            const textToInsert = selectedTexts.join('\n\n');
            const richTextToInsert = selectedRichTexts.join('<br><br>');

            // Check if we're using rich text and have rich text content available
            const shouldUseRichText = useRichText && selectedRichTexts.length > 0;
            
            // Check if the host is Outlook
            if (Office.context.mailbox) {
                // First check if the current email is in HTML format by getting it as HTML
                Office.context.mailbox.item.body.getAsync(
                    Office.CoercionType.Html,
                    function (htmlResult) {
                        if (htmlResult.status === Office.AsyncResultStatus.Succeeded) {
                            // Email supports HTML
                            const htmlContent = htmlResult.value || "";
                            const isHtmlEmail = htmlContent.includes('<') && htmlContent.includes('>');
                            
                            // Determine final coercion type based on both current email format and user selection
                            const finalCoercionType = isHtmlEmail ? Office.CoercionType.Html : 
                                                     (shouldUseRichText ? Office.CoercionType.Html : Office.CoercionType.Text);
                            
                            // Get the content in the appropriate format
                            Office.context.mailbox.item.body.getAsync(
                                finalCoercionType,
                                function (getResult) {
                                    if (getResult.status === Office.AsyncResultStatus.Succeeded) {
                                        let content = getResult.value || "";
                                        
                                        // Create a separator between existing content and the disclaimers
                                        let separator = "";
                                        if (content && content.trim() !== "") {
                                            separator = finalCoercionType === Office.CoercionType.Html ? "<br><br>" : "\n\n";
                                        }
                                        
                                        // Choose appropriate content for insertion
                                        let newContent;
                                        if (finalCoercionType === Office.CoercionType.Html) {
                                            // If we have rich text available and it's selected, use it
                                            if (shouldUseRichText) {
                                                newContent = richTextToInsert;
                                            } else {
                                                // Otherwise, convert plain text to HTML format
                                                newContent = textToInsert.replace(/\n\n/g, '<br><br>').replace(/\n/g, '<br>');
                                            }
                                        } else {
                                            newContent = textToInsert;
                                        }
                                        
                                        // Append the new content to the existing content
                                        const fullContent = content + separator + newContent;
                                        
                                        // Replace the entire body with the updated content
                                        Office.context.mailbox.item.body.setAsync(
                                            fullContent,
                                            { coercionType: finalCoercionType },
                                            function (setResult) {
                                                if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                                                    console.log("Disclaimers inserted successfully at the end of the email!");
                                                } else {
                                                    console.log("Error inserting disclaimers at the end: " + setResult.error.message);
                                                }
                                            }
                                        );
                                    } else {
                                        console.log("Error getting email body: " + getResult.error.message);
                                        
                                        // Fallback to insertion at cursor position with appropriate format
                                        const insertContent = finalCoercionType === Office.CoercionType.Html ? 
                                                             (shouldUseRichText ? richTextToInsert : textToInsert.replace(/\n\n/g, '<br><br>').replace(/\n/g, '<br>')) : 
                                                             textToInsert;
                                        
                                        Office.context.mailbox.item.body.setSelectedDataAsync(
                                            insertContent,
                                            { coercionType: finalCoercionType },
                                            function (asyncResult) {
                                                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                                                    console.log("Fallback: Disclaimers inserted at cursor position!");
                                                } else {
                                                    console.log("Error with fallback insert: " + asyncResult.error.message);
                                                }
                                            }
                                        );
                                    }
                                }
                            );
                        } else {
                            // Fallback to old approach if we can't determine HTML capability
                            const fallbackCoercionType = shouldUseRichText ? Office.CoercionType.Html : Office.CoercionType.Text;
                            const fallbackContent = shouldUseRichText ? richTextToInsert : textToInsert;
                            
                            Office.context.mailbox.item.body.getAsync(
                                fallbackCoercionType,
                                function (getResult) {
                                    if (getResult.status === Office.AsyncResultStatus.Succeeded) {
                                        let content = getResult.value || "";
                                        let separator = "";
                                        
                                        if (content && content.trim() !== "") {
                                            separator = fallbackCoercionType === Office.CoercionType.Html ? "<br><br>" : "\n\n";
                                        }
                                        
                                        const fullContent = content + separator + fallbackContent;
                                        
                                        Office.context.mailbox.item.body.setAsync(
                                            fullContent,
                                            { coercionType: fallbackCoercionType },
                                            function (setResult) {
                                                if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                                                    console.log("Disclaimers inserted successfully using fallback approach!");
                                                } else {
                                                    console.log("Error inserting disclaimers: " + setResult.error.message);
                                                }
                                            }
                                        );
                                    } else {
                                        console.log("Error getting email body: " + getResult.error.message);
                                    }
                                }
                            );
                        }
                    }
                );
            } else if (Office.context.host === Office.HostType.Word) {
                // For Word, we can append text to the end of the document
                Word.run(function (context) {
                    const contentToInsert = shouldUseRichText ? richTextToInsert : textToInsert;
                    const body = context.document.body;
                    
                    if (shouldUseRichText) {
                        body.insertHtml(contentToInsert, Word.InsertLocation.end);
                    } else {
                        body.insertText(contentToInsert, Word.InsertLocation.end);
                    }
                    
                    return context.sync()
                        .then(function() {
                            console.log("Disclaimers inserted successfully at the end of the document!");
                        });
                }).catch(function (error) {
                    console.log("Error: " + error.message);
                    
                    // Fallback to regular insertion at cursor
                    insertText(textToInsert, false);
                });
            } else if (Office.context.host === Office.HostType.PowerPoint) {
                // For PowerPoint, use text only as HTML might not be fully supported
                Office.context.document.setSelectedDataAsync(
                    textToInsert, 
                    { coercionType: Office.CoercionType.Text },
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log("Disclaimers inserted successfully in PowerPoint!");
                        } else {
                            console.log("Error inserting disclaimers in PowerPoint: " + asyncResult.error.message);
                        }
                    }
                );
            } else {
                // For Excel and other apps, always use Text coercion
                Office.context.document.setSelectedDataAsync(
                    textToInsert,
                    { coercionType: Office.CoercionType.Text },
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log("Disclaimers inserted successfully!");
                        } else {
                            console.log("Error inserting disclaimers: " + asyncResult.error.message);
                        }
                    }
                );
            }
        } catch (error) {
            console.log("Error: " + error.message);
        }
    }

    console.log('MyCompany Disclaimers Add-in loaded.');
})();