'use strict';

(function () {
    Office.onReady(function (info) {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            
            // Initialize the Rich Text toggle as disabled by default
            $('#useRichTextContainer').hide();

            // Check if the host supports rich text
            if (info && (info.host === Office.HostType.Word || 
                         info.host === Office.HostType.PowerPoint || 
                         info.host === Office.HostType.Outlook)) {
                $('#useRichTextContainer').show();
            }
            
            // Automatically load disclaimers from SharePoint list when page loads
            fetchDisclaimersSharePointList();

            // Add event listeners for the action buttons
            $('#insertSelectedDisclaimersAtCursor').on('click', function () {
                insertSelectedDisclaimersAtCursor();
            });

            $('#insertSelectedDisclaimersAsHeader').on('click', function () {
                insertSelectedDisclaimersAsHeader();
            });

            $('#insertSelectedDisclaimersAsFooter').on('click', function () {
                insertSelectedDisclaimersAsFooter();
            });
        });
    });

    //// Fetch disclaimers from a Web API
    //function fetchDisclaimers() {
    //    fetch('https://mycompanydisclaimers.azurewebsites.us/api/disclaimer')
    //        .then(response => response.json())
    //        .then(data => {
    //            let disclaimers = data.value;
    //            let disclaimerList = $('#disclaimer-list');
    //            disclaimerList.empty();
    //            disclaimers.forEach(disclaimer => {
    //                let disclaimerElementWrapper = $('<div class="disclaimer-button-wrapper"></div>');
    //                let disclaimerElement = $(`<button class="ms-Button">${disclaimer.description}</button>`);
    //                disclaimerElement.on("click", function () {
    //                    insertText(disclaimer.text);
    //                });
    //                disclaimerElementWrapper.append(disclaimerElement);
    //                disclaimerList.append(disclaimerElementWrapper);
    //            });
    //        })
    //        .catch(error => {
    //            console.log('Error:', error);
    //            showMessage(`Failed to load disclaimers from Web API: ${error.message}`);
    //        });
    //}

    // Fetch disclaimers from a SharePoint list
    async function fetchDisclaimersSharePointList() {
        fetch('https://mycompanydisclaimers.azurewebsites.us/api/sharepointlist')
            .then(response => response.json())
            .then(data => {
                let disclaimers = data.value;
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

                    // Append the checkbox and label to the list item
                    listItem.append(checkbox);
                    listItem.append(label);

                    // Append the list item to the list
                    list.append(listItem);
                });
            })
            .catch(error => {
                console.log('Error:', error);
                showMessage(`Failed to load disclaimers from SharePoint List: ${error.message}`);
            });
    }


    function showMessage(text) {
        const appendedText = $('#disclaimer-list').html() + text + "<br>---";
        $('#disclaimer-list').html(appendedText); // Targeting #disclaimer-list for displaying the message
    }

    function insertText(text, isRichText = false) {
        const coercionType = isRichText ? Office.CoercionType.Html : Office.CoercionType.Text;
        
        // Check if the host is Outlook
        if (Office.context.mailbox) {
            // Use the setSelectedDataAsync method on the body object to insert text
            Office.context.mailbox.item.body.setSelectedDataAsync(text,
                { coercionType: coercionType },
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
            
            insertText(shouldUseRichText ? richTextToInsert : textToInsert, shouldUseRichText);
        } catch (error) {
            console.log("Error: " + error.message);
        }
    }

})();
