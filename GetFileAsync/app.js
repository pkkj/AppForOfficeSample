(function () {
    "use strict";

    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#getFileAsync').click(getFileAsyncInternal);
        });
    };

    function encodeBase64(docData) {
        var s = "";
        for (var i = 0; i < docData.length; i++)
            s += String.fromCharCode(docData[i]);
        return window.btoa(s);
    }

    // Call getFileAsnyc() to start the retrieving file process.
    function getFileAsyncInternal() {
        Office.context.document.getFileAsync("compressed", { sliceSize: 10240 }, function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                document.getElementById("log").textContent = JSON.stringify(asyncResult);
            }
            else {
                getAllSlices(asyncResult.value);
            }
        });
    }

    // Get all the slices of file from the host after "getFileAsync" is done.
    function getAllSlices(file) {
        var sliceCount = file.sliceCount;
        var sliceIndex = 0;
        var docdata = [];
        var getSlice = function () {
            file.getSliceAsync(sliceIndex, function (asyncResult) {
                if (asyncResult.status == "succeeded") {
                    docdata = docdata.concat(asyncResult.value.data);
                    sliceIndex++;
                    if (sliceIndex == sliceCount) {
                        file.closeAsync();
                        onGetAllSlicesSucceeded(docdata);
                    }
                    else {
                        getSlice();
                    }
                }
                else {
                    file.closeAsync();
                    document.getElementById("log").textContent = JSON.stringify(asyncResult);
                }
            });
        };
        getSlice();
    }

    // Upload the docx file to server after obtaining all the bits from host.
    function onGetAllSlicesSucceeded(docxData) {
        $.ajax({
            type: "POST",
            url: "Handler.ashx",
            data: encodeBase64(docxData),
            contentType: "application/json; charset=utf-8",
        }).done(function (data) {
            document.getElementById("documentXmlContent").textContent = data;
        }).fail(function (jqXHR, textStatus) {
        });
    }
})();