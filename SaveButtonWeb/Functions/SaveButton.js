// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
        window.Promise = OfficeExtension.Promise;
    };
})();
var State = /** @class */ (function () {
    function State(file, counter, sliceCount) {
        this.file = file;
        this.counter = counter;
        this.sliceCount = sliceCount;
    }
    return State;
}());
;
// Get all of the content from a PowerPoint or Word document in 100-KB chunks of text.
function writeText(event) {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 100000 }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            // Get the File object from the result.
            var myFile = result.value;
            var state = new State(myFile, 0, myFile.sliceCount);
            updateStatus("Getting file of " + myFile.size + " bytes");
            getSlice(state);
        }
        else {
            updateStatus(result.status.toString());
        }
    });
}
// Create a function for writing to the log. 
function updateStatus(message) {
    console.log(message);
}
// Get a slice from the file and then call sendSlice.
function getSlice(state) {
    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
            sendSlice(result.value, state);
        }
        else {
            updateStatus(result.status.toString());
        }
    });
}
function sendSlice(slice, state) {
    var data = slice.data;
    // If the slice contains data, create an HTTP request.
    if (data) {
        console.log(typeof (data));
        var fileData = _arrayBufferToBase64(data);
        //var b = base64ToArrayBuffer(fileData);
        //var fileData = b64EncodeUnicode(data);
        // Create a new HTTP request. You need to send the request 
        // to a webpage that can receive a post.
        var request = new XMLHttpRequest();
        // Create a handler function to update the status 
        // when the request has been sent.
        request.onreadystatechange = function () {
            if (request.readyState === 4) {
                updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;
                if (state.counter < state.sliceCount) {
                    getSlice(state);
                }
                else {
                    closeFile(state);
                }
            }
        };
        //request.open("GET", "https://localhost:4443/api/v1/servicetemplates/types", false);
        //request.open("POST", "https://localhost:4443/v1/api/servicetemplatedocument/2/docx", false);
        //request.open("POST", "http://localhost:4065/api/v1/servicetemplatedocument/2/docx", true);
        //let auth = "Basic YWdpbGV0aG91Z2h0OkhlbGxvRE1DMQ==";
        //request.setRequestHeader("Authorization", auth);
        request.open("POST", "https://localhost:8000/api/Services", true);
        request.setRequestHeader("Slice-Number", slice.index);
        //request.setRequestHeader("Authorization", auth);
        request.setRequestHeader("Access-Control-Allow-Origin", "*");
        request.setRequestHeader("Access-Control-Allow-Headers", "slice-number");
        //request.setRequestHeader("Content-Type", "multipart/form-data");
        //request.setRequestHeader("Content-Type", "application/octet-stream");
        request.setRequestHeader("Content-Type", "application/octet-binary");
        // Send the file as the body of an HTTP POST 
        // request to the web server.
        //fileData = "test";
        //let blob = new Blob(fileData, { type: 'multipart/form-data' });
        //12k bytes
        //console.log('File length: ' + fileData.length);
        request.send(fileData);
    }
}
function closeFile(state) {
    // Close the file when you're done with it.
    state.file.closeAsync(function (result) {
        // If the result returns as a success, the
        // file has been successfully closed.
        if (result.status === "succeeded") {
            updateStatus("File closed.");
        }
        else {
            updateStatus("File couldn't be closed.");
        }
    });
}
function b64EncodeUnicode(str) {
    // first we use encodeURIComponent to get percent-encoded UTF-8,
    // then we convert the percent encodings into raw bytes which
    // can be fed into btoa.
    return btoa(encodeURIComponent(str).replace(/%([0-9A-F]{2})/g, function toSolidBytes(match, p1) {
        return String.fromCharCode(parseInt('0x', 8) + p1);
    }));
}
function _arrayBufferToBase64(buffer) {
    var binary = '';
    var bytes = new Uint8Array(buffer);
    var len = bytes.byteLength;
    for (var i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
}
function base64ToArrayBuffer(base64) {
    var binary_string = window.atob(base64);
    var len = binary_string.length;
    var bytes = new Uint8Array(len);
    for (var i = 0; i < len; i++) {
        bytes[i] = binary_string.charCodeAt(i);
    }
    return bytes.buffer;
}
function loadText(event) {
    Word.run(function (context) {
        var request = new XMLHttpRequest();
        var document;
        request.open('GET', 'https://localhost:8000/api/Services', false);
        request.send();
        if (request.status === 200) {
            document = context.application.createDocument(request.response);
        }
        //Office.context.document.setSelectedDataAsync(myXML, { coercionType: Office.CoercionType.Text });
        return context.sync()
            .then(function () {
            document.open();
            context.sync();
        })["catch"](function (error) {
            console.log(error);
        });
    });
    //Word.run(async(context) = > {
    //    await context.sync(); 
    //}).catch((error) => {
    //    console.log(error);
    //});   
}
