(function () {
        "use strict";

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#test').click(test);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function test() {                        
            /*
            console.log("prima");
            $.ajax({
                url: 'https://localhost:9000/',
                method: "GET",
                success: function(data) {
                    console.log(data);
                },
                error: function() {
                    console.log('network error');
                }
            });
            console.log("dopo");
            */
            
            var request = {
                job: [
                    {
                        method: 'com.stata.sfi.Data.getObsCount'
                    }
                ]
            };

            console.log("prima");
            $.ajax({
                url: 'https://localhost:50000',
                data: swire.encode(request),
                method: "POST",
                success: function(data) {
                    console.log('SUCCESS');
                    var response = swire.decode(data);
                    if (response.status === 'ok')
                        //console.log(response.output[0].output);
                        //writeToWordDocument('prova');
                        writeToWordDocument(response.output[0].output.toString());
                },
                error: function() {
                    console.log('network error');
                }
            });
            console.log("dopo");
        }

        function writeToWordDocument(text) {
            Word.run(function (context) {
                var range = context.document.getSelection();    
                range.insertText(text, Word.InsertLocation.end);

                return context.sync().then(function () {
                     console.log('Text added to the end of the range.');
                });

            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();