var app = angular.module('ExhibitsApp', ['ngMaterial'], function ($mdThemingProvider) {
    $mdThemingProvider.theme('default')
        .primaryPalette('grey', {
            'default': '800',
        });
});
app.controller('ExhibitsCtrl', function ($scope, $mdToast, $log, $http, $mdDialog) {
    ProgressLinearActive();

    Office.onReady(function (info) {
        console.log("js file loaded");
        localStorage.removeItem("Confirm");

        //Word.run(function (context) {
        //    var body = context.document.body;
        //    var paragraphs = body.paragraphs;

        //    context.load(paragraphs, 'text, font');

        //    return context.sync().then(function () {
        //        var exhibitCount = 1; // Start with Exhibit A

        //        for (var i = 0; i < paragraphs.items.length; i++) {
        //            var paragraph = paragraphs.items[i];
        //            var text = paragraph.text;
        //            var font = paragraph.font;

        //            if (font.italic && font.bold) {
        //                text = text.replace(/Exhibit [A-Z]/, "Exhibit " + String.fromCharCode(65 + exhibitCount));
        //                paragraph.insertText(text, Word.InsertLocation.replace);
        //                exhibitCount++;
        //            }
        //        }

        //        return context.sync();
        //    });
        //}).catch(function (error) {
        //    console.error("Error: " + JSON.stringify(error));
        //});










        /////////////////////////////////////////////////////////////////////////////////////

        function replaceBoldItalicExhibits() {
            Word.run(function (context) {
                var body = context.document.body;
                var searchResults = body.search("Exhibit [A-Za-z]", { matchCase: false, matchWildcards: true, formatting: { bold: true, italic: true } });

                context.load(searchResults, 'text, font');

                return context.sync().then(function () {
                    if (searchResults.items.length > 0) {
                        var exhibitCounter = 0;

                        searchResults.items.forEach(function (result) {
                            var exhibitText = result.text;
                            var replacementText = "Exhibit " + String.fromCharCode(65 + exhibitCounter);

                            result.insertText(replacementText, Word.InsertLocation.replace);

                            exhibitCounter++;
                        });
                    } else {
                        console.log("No bold and italic Exhibits found.");
                    }
                });
            }).catch(function (error) {
                console.error("Error: " + JSON.stringify(error));
            });
        }

        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);

        function handler(evtArgs) {
            Word.run(function (context) {
                var range = context.document.getSelection();
                context.load(range);
                return context.sync().then(function () {
                    replaceBoldItalicExhibits();
                });
            });
        };

        //Word.run(function (context) {
        //    const paragraphs = context.document.body;
        //    paragraphs.load();
        //    return context.sync().then(function () {
        //        for (let i = 0; i < paragraphs.items.length; i++) {
        //            paragraphs.items[i].ParagraphChanged.add(onDocumentSelectionChanged)

        //        }


        //    });
        //}).catch(function (error) {
        //    console.error('Error: ' + error);
        //});


        //Word.run(function (context) {

        //    var document = Office.context.document;

        //    document.addHandlerAsync(Word.EventType.DocumentSelectionChanged, onDocumentSelectionChanged);
        //    return context.sync();
        //}).catch(function (error) {
        //    console.error("Error: " + JSON.stringify(error));
        //});

        //function onDocumentSelectionChanged(eventArgs) {
        //    var selectedContent = eventArgs.selection.text;
        //    if (selectedContent.trim() !== "") {
        //        console.log("Text pasted in the document:", selectedContent);
        //    }
        //}

        ////////////////////////////////////////




        //<------------------------------Add New Exhibit------------------------------>//


        $scope.AddNewExhibit = function(){
                Word.run(function (context) {
                    var selectedRange = context.document.getSelection();
                    selectedRange.load("text");

                    return context.sync().then(function () {
                        var selectedText = selectedRange.text;

                        // Check if the selected text contains the word "Exhibit"
                        var exhibitIndex = selectedText.indexOf("Exhibit");
                        if (exhibitIndex !== -1) {
                            var beforeExhibit = selectedText.substring(0, exhibitIndex);
                            var exhibitText = "Exhibit";
                            var afterExhibit = selectedText.substring(exhibitIndex + exhibitText.length);

                            console.log("Before Exhibit: " + beforeExhibit);
                            console.log("Exhibit: " + exhibitText);
                            console.log("After Exhibit: " + afterExhibit);

                            var addobj = {
                                AddLead: beforeExhibit,
                                AddExhibit: exhibitText,
                                AddDescription: afterExhibit,
                            }
                            localStorage.setItem("AddExhibit", JSON.stringify(addobj));
                            AddExhibitDialog();
                        } else {
                            console.log("Selected text does not contain the word 'Exhibit'");
                            loadToast("Please seclect a row with Exhibit Text");
                        }
                    });
                }).catch(function (error) {
                    console.error("Error: " + JSON.stringify(error));
                });
            
        }
     


        var adddialog;
        function AddExhibitDialog () {
            Office.context.ui.displayDialogAsync('https://aamirhusnain.github.io/EvidenceReference/EvidenceReferenceWeb/AddExhibit.html', { height: 60, width: 35 },
                function (asyncResult) {
                    adddialog = asyncResult.value;
                    adddialog.addEventHandler(Office.EventType.DialogMessageReceived, addprocessMessage);
                }
            );
        };

        function addprocessMessage(arg) {
            adddialog.close();
            var message = arg.message;
            //  console.log(message);
            //  var data = JSON.parse(mesage);
            //  console.log(data);
            ExhibitName(message);
        };

        function ExhibitName(message) {
            Word.run(function (context) {
                var documentBody = context.document.body;
                var selectedRange = context.document.getSelection();

                context.load(documentBody);
                context.load(selectedRange);

                return context.sync()
                    .then(function () {
                        var selectedText = selectedRange.text;

                        var regex = /Exhibit\s+\w+/g;
                        var matches = selectedText.match(regex);

                        if (matches && matches.length > 0) {
                            matches.forEach(function (match) {
                                console.log('Found: ' + match);
                                var Exhibitid = match;

                                var data = JSON.parse(message);

                                data.ExhibitName = Exhibitid;

                                console.log(data);
                                $scope.addNewPage(JSON.stringify(data), Exhibitid);

                            });
                        } else {
                            // If "Exhibit" pattern is not found, get the first word of the selected text
                            var firstWord = selectedText.split(/\s+/)[0];
                            console.log('No occurrences of "Exhibit" followed by the next word found.');
                            console.log('First word of selected text: ' + firstWord);
                            var Exhibitid = firstWord;

                            //<------------------------------------------->//
                            var data = JSON.parse(message);
                            data.ExhibitName = Exhibitid;
                            console.log(data);
                            $scope.addNewPage(JSON.stringify(data), Exhibitid)

                            //  $scope.showBindingData(message, Exhibitid);
                        }
                    })
                    .catch(function (error) {
                        console.log("Error: " + error);
                    });
            });

        };

        // <--------------------------- Add Binding ------------------------>//

        $scope.exhibits = [];
        $scope.Exhibitdatas = [];
        $scope.addNewPage = function (data, Exhibitid) {
            $scope.exhibits.push(Exhibitid);
            $scope.Exhibitdatas.push({ data: data, Exhibitid: Exhibitid });
            console.log($scope.Exhibitdatas);
            console.log($scope.exhibits);


            if (!$scope.$$phase) {
                $scope.$apply();
            }
            // addNewPageWithHeaderAndBody(data);
        };

        //<-----------------------------Get Data for Edit----------------------------------->//

        function findObjectByExhibitid(exhibit) {
            return $scope.Exhibitdatas.find(function (item) {
                return item.Exhibitid === exhibit;
            });
        }

        $scope.Edit = function (exhibit) {
            $scope.selectBoldItalicLineWithText(exhibit);
            localStorage.setItem("Id", exhibit);
            var foundExhibit = findObjectByExhibitid(exhibit);

            if (foundExhibit) {
                var exhibitData = JSON.parse(foundExhibit.data);
                console.log(exhibitData);
                localStorage.setItem('data', JSON.stringify(exhibitData));

                OpenEditDialog();
            } else {
                console.log('Exhibit not found');
            }

        };

        //<------------------------------------------------------------------------------------>//

        function editExhibit(exhibitIdToUpdate, newValues) {
            var exhibitToEdit = $scope.Exhibitdatas.find(function (item) {
                return item.Exhibitid === exhibitIdToUpdate;
            });

            if (exhibitToEdit) {
                var exhibitData = JSON.parse(exhibitToEdit.data);
                var newValues = JSON.parse(newValues);
                exhibitData.Description = newValues.Description || exhibitData.Description;
                exhibitData.Exhibit = newValues.Exhibit || exhibitData.Exhibit;
                exhibitData.Lead = newValues.Lead || exhibitData.Lead;
                exhibitData.Link = newValues.Link || exhibitData.Link;
                exhibitToEdit.data = JSON.stringify(exhibitData);
                console.log($scope.Exhibitdatas);
            }
        };

        //<-----------------------------Get Data for Edit end------------------------------>//


        //function addNewPageWithHeaderAndBody(data) {
        //    Word.run(function (context) {
        //        var body = context.document.body;

        //        var row = JSON.parse(data);

        //        // Get the primary header or create one if it doesn't exist
        //        var header = context.document.sections.getFirst().getHeader("primary");
        //        if (!header) {
        //            header = context.document.sections.getFirst().addHeader(Word.HeaderFooterType.primary);
        //        }

        //        // Clear the existing header
        //        header.clear();

        //        // Set the ExhibitName property in the center of the header

        //        var headerParagraph = header.insertParagraph(row.ExhibitName, Word.InsertLocation.start);
        //        //headerParagraph.alignment = Word.ParagraphAlignment.center;
        //        headerParagraph.font.size = 16;

        //        // Insert a page break to start a new page
        //        body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

        //        // Insert other properties in the body
        //        Object.keys(row).forEach(function (key) {
        //            if (key !== "Id" && key !== "ExhibitName") {
        //                var paragraph = body.insertParagraph(key + ": " + row[key], Word.InsertLocation.end);

        //                paragraph.font.size = 14;
        //                //paragraph.font.italic = true;
        //            }
        //        });

        //        return context.sync().then(function () {
        //            console.log("New page added with header and body data.");
        //        });
        //    }).catch(function (error) {
        //        console.error("Error: " + error);
        //    });
        //}


        //<------------------------------------------------------------>//

        $scope.selectBoldItalicLineWithText = function (exhibit) {
            Word.run(function (context) {
                var body = context.document.body;
                var paragraphs = body.paragraphs;

                paragraphs.load("text, font");

                return context.sync().then(function () {
                    var selectedParagraphs = [];
                    for (var i = 0; i < paragraphs.items.length; i++) {
                        var paragraph = paragraphs.items[i];
                        var text = paragraph.text;
                        var pattern = new RegExp(`.*\\b\\s*${exhibit}\\s*\\b.*`, "i");
                        if (pattern.test(text) && paragraph.font.bold && paragraph.font.italic) {
                            selectedParagraphs.push(paragraph);
                        }
                    }

                    if (selectedParagraphs.length > 0) {
                        // Get the ranges of all selected paragraphs
                        var selectedRanges = selectedParagraphs.map(function (paragraph) {
                            return paragraph.getRange();
                        });

                        // Combine the ranges into a single range and select it
                        var combinedRange = selectedRanges.reduce(function (range1, range2) {
                            return range1.expandTo(range2);
                        });
                        combinedRange.select();
                    }
                });
            }).catch(function (error) {
                console.error("Error: " + JSON.stringify(error));
            });
        };

        //$scope.selectBoldItalicLineWithText = function (exhibit) {
        //        Word.run(function (context) {
        //            var body = context.document.body;
        //            var paragraphs = body.paragraphs;

        //            paragraphs.load("text, font");

        //            return context.sync().then(function () {
        //                for (var i = 0; i < paragraphs.items.length; i++) {
        //                    var paragraph = paragraphs.items[i];
        //                    if (paragraph.text.includes(exhibit) && paragraph.font.bold && paragraph.font.italic) {
        //                        paragraph.select();
        //                        break; 
        //                    }
        //                }
        //            });
        //        }).catch(function (error) {
        //            console.error("Error: " + JSON.stringify(error));
        //        });
        //    }

        //<------------------------------------------------------------>//

        //<----------------open Edit Dialog------------>//

        var dialog;
        function OpenEditDialog() {
            Office.context.ui.displayDialogAsync('https://aamirhusnain.github.io/EvidenceReference/EvidenceReferenceWeb/EditExhibit.html', { height: 65, width: 35 },
                function (asyncResult) {
                    dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                }
            );
        };

        function processMessage(arg) {
            dialog.close();
            var message = arg.message;
            SaveExitdata(message);
        };

        function SaveExitdata(message) {
            //  var data = JSON.parse(message)
            //  console.log(data);
            var exhibitIdToUpdate = localStorage.getItem('Id');
            var newValues = message;
            editExhibit(exhibitIdToUpdate, newValues);
        }

        //<------------------Show Exhibit Text as Sclected------------------------>//
        //<-------------------Save Document----------------->//

        function saveDocument() {
            Word.run(function (context) {
                var doc = context.document;
                doc.save();
                return context.sync();
            }).catch(function (error) {
                console.log(error.message);
            });
        };

        //<-----------------------Get LocalStorage---------------------->//

        $scope.ChangText = function () {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (asyncResult) => {
                $scope.selectedText = asyncResult.value;
                if ($scope.selectedText) {
                    console.log($scope.selectedText);
                }
                $scope.$apply();
            });

            //<-----------------create Text Italic and bold------------------->//

            Word.run(function (context) {
                const selection = context.document.getSelection();
                selection.font.bold = true;
                selection.font.italic = true;
                return context.sync();
                console.log('The selection is now bold and italic.');
            });
        };

        //<----------------------------Set A New Custom Property----------------------------->//

        //<---------------------Delete Exhibit------------------->//
        $scope.DeleteExhibit = function (exhibit) {
            $scope.selectBoldItalicLineWithText(exhibit);
            var confirm = $mdDialog.confirm()
                .title('Warning')
                .textContent('Are you sure you want to delete this Exhibit?')
                .ariaLabel('Delete Confirmation')
                .ok('Yes')
                .cancel('No');

            $mdDialog.show(confirm).then(function () {
                var indexToDelete = $scope.Exhibitdatas.findIndex(function (item) {
                    return item.Exhibitid === exhibit;
                });

                if (indexToDelete !== -1) {
                    $scope.Exhibitdatas.splice(indexToDelete, 1);
                    // Reset all ExhibitId values
                    for (var i = 0; i < $scope.Exhibitdatas.length; i++) {
                        $scope.Exhibitdatas[i].Exhibitid = "Exhibit " + String.fromCharCode(65 + i); // A, B, C, ...
                    }
                } else {
                    console.log("not found");
                }
                var indexToDelete = $scope.exhibits.indexOf(exhibit);

                if (indexToDelete !== -1) {
                    $scope.exhibits.splice(indexToDelete, 1);
                    console.log($scope.exhibits);

                    for (var i = 0; i < $scope.exhibits.length; i++) {
                        $scope.exhibits[i] = "Exhibit " + String.fromCharCode(65 + i);
                    }
                } else {
                    console.log("Not found");
                }


                if (!$scope.$$phase) {
                    $scope.$apply();
                }

                Word.run(function (context) {
                    const range = context.document.getSelection();
                    range.clear();
                    return context.sync();
                });

                saveDocument();

            },
                function () {
                    console.log("No");
                });
        };

        //<--------------------------------Convert  to  PDF----------------------------->//

        //$scope.ConvertToPDF = function () {
        //    var document = Office.context.document;
        //    document.getFilePropertiesAsync(function (asyncResult) {
        //        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        //            console.log("The file hasn't been saved yet. Save the file and try again");
        //        } else {
        //            var fileUrl = asyncResult.value.url;
        //            var fileName = fileUrl.substring(fileUrl.lastIndexOf('/') + 1);
        //            console.log("File URL:", fileUrl);
        //            console.log("File Name:", fileName);
        //            // ConvertToPDF(fileUrl, fileName)
        //            if (fileName === "") {
        //                loadToast("The file hasn't been saved yet. Save the file and try again");
        //            }
        //        }
        //    });
        //};

        //function ConvertToPDF (fileUrl, fileName) {
        //    var settings = {
        //        "url": 'https://api.pdf.co/v1/pdf/convert/from/doc',
        //        "method": "POST",
        //        "timeout": 0,
        //        "headers": {
        //            "x-api-key": "defaulterj042@gmail.com_7a2de32bfb84270f3b6f6a5baf3061d3e99fcc386141dfd1d9f38aafeaec45b6c6796e5f",
        //            "Content-Type": "application/json",
        //        },
        //        "data": JSON.stringify({
        //            "name": fileName,
        //            "url": fileUrl
        //        }),
        //    };

        //    $.ajax(settings).done(function (response) {
        //        console.log(response);
        //        console.log(response.url)
        //    }).catch(function (error) {
        //        console.log(error);
        //    })

        //}

        //<----------------------------------------------------------------------------->//

        //$scope.convertObjectsToPDF = function () {
        //    var Confirmdata = localStorage.getItem("Confirm");
        //    console.log(Confirmdata);
        //    if (Confirmdata == null) {
        //        openDialogClick()
        //    } else {
        //        convertObjectsToPDF()
        //    }
        //}

        //function convertObjectsToPDF(Confirmdata) {
        //    var Cdata = localStorage.getItem("Confirm");
        //    var Confirmdata = JSON.parse(Cdata);
        //    var Affiant = Confirmdata.Affiant;
        //    var Commissioner = Confirmdata.commissioner;
        //    var CDate = Confirmdata.Date;
        //    function formatDate(CDate) {
        //        const date = new Date(CDate);
        //        const day = date.getDate();
        //        const month = date.toLocaleString('en-us', { month: 'long' });
        //        const year = date.getFullYear();

        //        return `${day}${getOrdinalSuffix(day)} Day of ${month}, ${year}`;
        //    }

        //    function getOrdinalSuffix(day) {
        //        if (day >= 11 && day <= 13) {
        //            return 'th';
        //        }
        //        switch (day % 10) {
        //            case 1:
        //                return 'st';
        //            case 2:
        //                return 'nd';
        //            case 3:
        //                return 'rd';
        //            default:
        //                return 'th';
        //        }
        //    }

        //    const formattedDate = formatDate(CDate);
        //    console.log(formattedDate);


        //    const docDefinition = {
        //        content: []
        //    };

        //    // Iterate through the objects array and add cover page + data page for each object
        //    $scope.Exhibitdatas.forEach((object, index) => {
        //        const coverPage = {
        //            text: [
        //                { text: "ThIS IS ", fontSize: 16, bold: true },
        //                { text: object.Exhibitid + " REFERRED" + '\n', fontSize: 14, bold: true, },
        //                { text: "TO IN THE AFFIDAVIT OF ", fontSize: 14, bold: true },
        //                { text: Affiant + '\n', fontSize: 16, bold: true, },
        //                { text: "SWORN THIS " + formattedDate + '\n', fontSize: 14, bold: true, },
        //                { text: '\n' },
        //                { text: '\n' },
        //                { text: "A Commissioner for Taking Affidavits, etc.\n", fontSize: 14, bold: true },
        //                { text: Commissioner, fontSize: 14, bold: true, }
        //            ],
        //            alignment: 'center',
        //            margin: [0, 100, 0, 0]
        //        };

        //        const formattedData = formatDataForDisplay(object.data);
        //        const dataPage = {
        //            text: formattedData,
        //            margin: [40, 40, 40, 40]
        //        };

        //        if (index !== 0) {
        //            // Add a page break before each object's content page (except the first one)
        //            docDefinition.content.push({ text: '\u000C', pageBreak: 'before' });
        //        }

        //        docDefinition.content.push(coverPage);
        //        docDefinition.content.push(dataPage);
        //    });

        //    const pdfDocGenerator = pdfMake.createPdf(docDefinition);

        //    pdfDocGenerator.getBuffer(function (buffer) {
        //        const pdfBlob = new Blob([buffer], { type: 'application/pdf' });
        //        const pdfUrl = URL.createObjectURL(pdfBlob);
        //        const link = document.createElement('a');
        //        link.href = pdfUrl;
        //        link.download = 'objects_document.pdf';
        //        link.click();
        //    });
        //};

        //function formatDataForDisplay(data) {
        //    const parsedData = JSON.parse(data);
        //    const formattedData = Object.keys(parsedData).map(key => {
        //        return `${key} : ${parsedData[key]}`;
        //    });
        //    return formattedData.join('\n');
        //}

        //<--------------------------------------------------------------->//

        function openDialogClick(ev) {
            $mdDialog.show({
                scope: $scope.$new(),
                templateUrl: 'https://aamirhusnain.github.io/EvidenceReference/EvidenceReferenceWeb/ConfirmDialog.html',
                clickOutsideToClose: true,
                targetEvent: ev,
                fullscreen: $scope.customFullscreen,
                controller: ['$scope', '$mdDialog', function ($scope, $mdDialog) {

                    $scope.confirm = function () {
                        var Affiant = $scope.Affiant;
                        var commissioner = $scope.Commissioner;
                        var Date = $scope.Date;
                        $mdDialog.hide({ Affiant: Affiant, commissioner: commissioner, Date: Date });
                    }

                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }
                }]
            }).then(function (result) {
                var Confirmdata = result;
                localStorage.setItem("Confirm", JSON.stringify(Confirmdata));
                $scope.convertToPDF();
            });

        };
        //<--------------------------------------------------------------->//
        //$scope.TOPDF = function () {
        //    Word.run(function (context) {
        //        const body = context.document.body;
        //        const bodyHTML = body.getHtml();

        //        return context.sync().then(function () {
        //            console.log("Body contents (HTML): " + bodyHTML.value);
        //            const htmlContent = bodyHTML.value;

        //            $.ajax({
        //                url: '/ConvertToPDF/ConvertToPdf', // Update the URL to match your controller action
        //                type: 'POST',
        //                contentType: 'application/json', // Set the content type to JSON
        //                data: JSON.stringify({ htmlContent: htmlContent }),// Pass the HTML content as data
        //                success: function (response) {
        //                    // Handle the response from the server (e.g., display or save the PDF)
        //                    console.log(response);
        //                },
        //                error: function (error) {
        //                    // Handle errors
        //                    console.log(error);
        //                }
        //            });
        //        });
        //    }).catch(function (error) {
        //        // Handle any Word API errors
        //        console.log(error);
        //    });

        //}


        $scope.convertToPDF = function () {
            return new Promise(function (resolve, reject) {
                Office.context.document.getFileAsync(Office.FileType.Pdf, { sliceSize: 4194304 }, function (result) {
                    try {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            var file = result.value;
                            file.getSliceAsync(0, function (sliceResult) {
                                try {
                                    if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                                        var dataSlice = sliceResult.value.data;
                                        var base64Data = btoa(String.fromCharCode.apply(null, new Uint8Array(dataSlice)));
                                        console.log(base64Data);
                                        resolve(base64Data);
                                        $scope.CompilingDocumentasPDF(base64Data);
                                       // convertToPdfUsingAjax(base64Data);
                                    } else {
                                        reject(sliceResult.error);
                                    }
                                } catch (error) {
                                    reject(error);
                                }
                            });
                        } else {
                            reject(result.error);
                        }
                    } catch (error) {
                        reject(error);
                    }
                });
            });
        };

        //function convertToPdfUsingAjax(base64Data,) {
        //    fetch('https://aamirhusnain.github.io/EvidenceReference/EvidenceReferenceWeb/Controllers/PDFTesting/ConvertToPdf', {
        //        method: 'POST',
        //        headers: {
        //            'Content-Type': 'application/json',
        //        },
        //        body: JSON.stringify({ data: base64Data }),
        //    })
        //        .then(response => response.blob())
        //        .then(blob => {
        //            const blobUrl = URL.createObjectURL(blob);
        //            const a = document.createElement('a');
        //            a.href = blobUrl;
        //            a.download = 'converted-document.pdf';
        //            a.style.display = 'none';
        //            document.body.appendChild(a);
        //            a.click();
        //            document.body.removeChild(a);
        //        })
        //        .catch(error => {
        //            console.error('Error:', error);
        //        });

        //   }


        //<--------------------------------------------------------------->//

        $scope.convertObjectsToPDF = function () {
            var Confirmdata = localStorage.getItem("Confirm");
            console.log(Confirmdata);
            if (Confirmdata == null) {
                openDialogClick()
            } else {
                $scope.convertToPDF();
            }
        }

        $scope.CompilingDocumentasPDF = function (base64Data) {

            var Cdata = localStorage.getItem("Confirm");
            var Confirmdata = JSON.parse(Cdata);
            var Affiant = Confirmdata.Affiant;
            var Commissioner = Confirmdata.commissioner;
            var CDate = Confirmdata.Date;
            function formatDate(CDate) {
                const date = new Date(CDate);
                const day = date.getDate();
                const month = date.toLocaleString('en-us', { month: 'long' });
                const year = date.getFullYear();

                return `${day}${getOrdinalSuffix(day)} Day of ${month}, ${year}`;
            }

            function getOrdinalSuffix(day) {
                if (day >= 11 && day <= 13) {
                    return 'th';
                }
                switch (day % 10) {
                    case 1:
                        return 'st';
                    case 2:
                        return 'nd';
                    case 3:
                        return 'rd';
                    default:
                        return 'th';
                }
            }

            const formattedDate = formatDate(CDate);
            console.log(formattedDate);

            var newArray = [];

            // Loop through the original array
            for (var i = 0; i < $scope.Exhibitdatas.length; i++) {
                var originalObject = $scope.Exhibitdatas[i];
                var newDataObject = {
                    Exhibitid: originalObject.Exhibitid,
                    Link: JSON.parse(originalObject.data).Link,
                    Affiant: Affiant,
                    Commissioner: Commissioner,
                    Date: formattedDate,

                };
                newArray.push(newDataObject);
            }
            console.log(newArray);

            fetch(''https://aamirhusnain.github.io/EvidenceReference/EvidenceReferenceWeb/Controllers//PDFTesting/ConvertToPdf', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    base64Data: base64Data, 
                    objectsArray: newArray
                }),
            })
                .then(response => response.blob())
                .then(blob => {
                    const blobUrl = URL.createObjectURL(blob);

                    const a = document.createElement('a');
                    a.href = blobUrl;
                    a.download = 'converted-document.pdf';
                    a.style.display = 'none';
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                })
                .catch(error => {
                    console.error('Error:', error);
                });

             };

        //<--------------------------------------------------------------->//

        ProgressLinearInActive();
       
    });

    //<----------------------------Loader & Toast----------------------------->//

    function ProgressLinearActive() {
        $("#StartProgressLinear").show(function () {
            $("#ProgressBgDiv").show();
            $scope.ddeterminateValue = 15;
            $scope.showProgressLinear = false;
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        });
    };

    function ProgressLinearInActive() {
        $("#StartProgressLinear").hide(function () {
            setTimeout(function () {
                $scope.ddeterminateValue = 0;
                $scope.showProgressLinear = true;
                $("#ProgressBgDiv").hide();
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            }, 500);
        });
    };

    function loadToast(alertMessage) {
        var el = document.querySelectorAll('#zoom');
        $mdToast.show(
            $mdToast.simple()
                .textContent(alertMessage)
                .position('bottom')
                .hideDelay(4000))
            .then(function () {
                console.log('Toast dismissed.');
            }).catch(function () {
                console.log('Toast failed or was forced to close early by another toast.');
            });
        if (!$scope.$$phase) {
            $scope.$apply();
        }
    };

    if (!$scope.$$phase) {
        $scope.$apply();
    }

});
