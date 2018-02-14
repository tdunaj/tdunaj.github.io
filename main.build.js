/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "/Dist";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports, __webpack_require__) {

__webpack_require__(1);
module.exports = __webpack_require__(2);


/***/ }),
/* 1 */
/***/ (function(module, exports) {

(function () {
    "use strict";
    var messageBanner;
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");
                return;
            }
        });
    };
    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();


/***/ }),
/* 2 */
/***/ (function(module, exports) {

// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
        window.Promise = OfficeExtension.Promise;
    };
})();
// Get all of the content from a PowerPoint or Word document in 4MB chunks of data.
function saveDocument(event) {
    Office.context.document.getFilePropertiesAsync(function (propertiesResult) {
        var fileUrl = propertiesResult.value.url;
        Office.context.document.getFileAsync(Office.FileType.Compressed, function (fileResult) {
            if (fileResult.status === Office.AsyncResultStatus.Succeeded) {
                // Get the File object from the result.
                var file = fileResult.value;
                //this should be moved out to a configuration; not really a great way to do that without better packaging
                sendFile(fileUrl, file, 'https://localhost:4443/api/v1/servicetemplates/2/docx');
            }
        });
    });
}
/**
 * Sends a file using HttpMultipart
 * @param fileUrl   The Url of the file (filename).
 * @param file      The file.
 * @param destURL   The Url where it will post the multipart request.
 */
function sendFile(fileUrl, file, destURL) {
    var BOUNDARY = 'HelloDMC_Multipart_Boundary';
    var BOUNDARY_DASHES = '--';
    var NEWLINE = '\r\n';
    var CONTENT_TYPE = 'Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    var CONTENT_DISPOSITION = "Content-Disposition: form-data; name=\"WriteUp\"; filename=\"" + fileUrl + "\"";
    var POST_DATA_START = [
        NEWLINE, BOUNDARY_DASHES, BOUNDARY, NEWLINE,
        CONTENT_DISPOSITION, NEWLINE, CONTENT_TYPE, NEWLINE, NEWLINE
    ].join('');
    var POST_DATA_END = [NEWLINE, BOUNDARY_DASHES, BOUNDARY, BOUNDARY_DASHES, NEWLINE].join('');
    var size = POST_DATA_START.length + file.size + POST_DATA_END.length;
    var unit8Array = new Uint8Array(size);
    var xhr = new XMLHttpRequest();
    xhr.open('POST', destURL, true);
    xhr.onreadystatechange = function () {
        if (xhr.readyState === 4) {
            file.closeAsync();
        }
    };
    //add start to byte array
    var i = 0;
    for (; i < POST_DATA_START.length; i++) {
        unit8Array[i] = POST_DATA_START.charCodeAt(i) & 0xFF;
    }
    //add slices to byte array
    //right now, this only works if there is a single slice (total size < 4MB)
    //we will need to modify this to chunk for sizes > 4MB in the future
    for (var counter = 0; counter < file.sliceCount; counter++) {
        file.getSliceAsync(counter, function (sliceResult) {
            if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                var slice = sliceResult.value;
                if (slice) {
                    var data = slice.data;
                    for (var j = 0; j < (data.length); i++, j++) {
                        unit8Array[i] = data[j];
                    }
                    //add end to byte array
                    for (var j = 0; i < size; i++, j++) {
                        unit8Array[i] = POST_DATA_END.charCodeAt(j) & 0xFF;
                    }
                    //set Basic Auth creds
                    xhr.setRequestHeader('Authorization', 'Basic ZWJlOGM3MDRjMjg3ZmFmNmQzZmM0YmQ3OTU5YzljYTY6JGMwMGJ5RDAw');
                    //send as multipart
                    xhr.setRequestHeader('Content-Type', 'multipart/form-data; boundary=' + BOUNDARY);
                    xhr.send(unit8Array.buffer);
                }
            }
        });
    }
}


/***/ })
/******/ ]);