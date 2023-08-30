var app = angular.module('ExhibitsApp', ['ngMaterial'], function ($mdThemingProvider) {
    $mdThemingProvider.theme('default')
        .primaryPalette('blue', {
            'default': '500',
        });
});
app.controller('ExhibitsCtrl', function ($scope, $mdToast, $log, $http, $mdDialog) {
    ProgressLinearActive();


    Office.onReady(function (info) {
        console.log("js file loaded");


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

        var adddialog;
        $scope.AddExhibitDialog = function () {
            //  var URL = 'https://localhost:44326/AddExhibit.html';
            var URL = '/EvidenceReference/EvidenceReferenceWeb/AddExhibit.html';
            Office.context.ui.displayDialogAsync(URL, { height: 60, width: 35 },
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
                localStorage.setItem('data', JSON.stringify(exhibitData))
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
                // Parse the data property JSON string
                var exhibitData = JSON.parse(exhibitToEdit.data);
                var newValues = JSON.parse(newValues);
                // Update specific properties with new values
                exhibitData.Description = newValues.Description || exhibitData.Description;
                exhibitData.Exhibit = newValues.Exhibit || exhibitData.Exhibit;
                exhibitData.Lead = newValues.Lead || exhibitData.Lead;
                exhibitData.Link = newValues.Link || exhibitData.Link;
                exhibitToEdit.data = JSON.stringify(exhibitData);
                console.log($scope.Exhibitdatas);
            }
        };


        //<-----------------------------Get Data for Edit end----------------------------------->//



        function addNewPageWithHeaderAndBody(data) {
            Word.run(function (context) {
                var body = context.document.body;

                var row = JSON.parse(data);

                // Get the primary header or create one if it doesn't exist
                var header = context.document.sections.getFirst().getHeader("primary");
                if (!header) {
                    header = context.document.sections.getFirst().addHeader(Word.HeaderFooterType.primary);
                }

                // Clear the existing header
                header.clear();

                // Set the ExhibitName property in the center of the header

                var headerParagraph = header.insertParagraph(row.ExhibitName, Word.InsertLocation.start);
                //headerParagraph.alignment = Word.ParagraphAlignment.center;
                headerParagraph.font.size = 16;

                // Insert a page break to start a new page
                body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

                // Insert other properties in the body
                Object.keys(row).forEach(function (key) {
                    if (key !== "Id" && key !== "ExhibitName") {
                        var paragraph = body.insertParagraph(key + ": " + row[key], Word.InsertLocation.end);

                        paragraph.font.size = 14;
                        //paragraph.font.italic = true;
                    }
                });

                return context.sync().then(function () {
                    console.log("New page added with header and body data.");
                });
            }).catch(function (error) {
                console.error("Error: " + error);
            });
        }


        //<------------------------------------------------------------>//

        $scope.selectBoldItalicLineWithText = function (exhibit) {
            Word.run(function (context) {
                var body = context.document.body;
                var paragraphs = body.paragraphs;

                paragraphs.load("text, font");

                return context.sync().then(function () {
                    for (var i = 0; i < paragraphs.items.length; i++) {
                        var paragraph = paragraphs.items[i];
                        if (paragraph.text.includes(exhibit) && paragraph.font.bold && paragraph.font.italic) {
                            paragraph.select();
                            break;
                        }
                    }
                });
            }).catch(function (error) {
                console.error("Error: " + JSON.stringify(error));
            });
        }

        //<-------------------------------------------------->//

        //<----------------open Edit Dialog------------>//

        var dialog;
        function OpenEditDialog() {

            //  var URL = 'https://localhost:44326/AddExhibit.html';
            var URL = '/EvidenceReference/EvidenceReferenceWeb/EditExhibit.html';

            Office.context.ui.displayDialogAsync(URL, { height: 65, width: 35 },
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

        //<--------------Show Exhibit Text as Sclected------------------------>//


        //<--------------------------------------------------------------->//

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

            //<--------------------create Text Italic and bold------------------->//

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
                } else {
                    console.log("not found")
                };


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

        $scope.ConvertToPDF = function () {
            var document = Office.context.document;
            document.getFilePropertiesAsync(function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log("The file hasn't been saved yet. Save the file and try again");
                } else {
                    var fileUrl = asyncResult.value.url;
                    var fileName = fileUrl.substring(fileUrl.lastIndexOf('/') + 1);
                    console.log("File URL:", fileUrl);
                    console.log("File Name:", fileName);
                    // ConvertToPDF(fileUrl, fileName)
                    if (fileName === "") {
                        loadToast("The file hasn't been saved yet. Save the file and try again");
                    }
                }
            });
        };

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


        //<--------------------------------------------------------------->//

        ProgressLinearInActive();
        if (!$scope.$$phase) {
            $scope.$apply();
        }
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