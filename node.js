// js/default.js
(function () {
  "use strict";

  var openDocxBtn, openWebBtn, statusEl, docxContainer, webViewer;

  document.addEventListener("DOMContentLoaded", function () {
    openDocxBtn = document.getElementById("openDocxBtn");
    openWebBtn  = document.getElementById("openWebBtn");
    statusEl    = document.getElementById("status");
    docxContainer = document.getElementById("docxContainer");
    webViewer   = document.getElementById("webViewer");

    openDocxBtn.addEventListener("click", openDocxLocal);
    openWebBtn.addEventListener("click", openDocxWeb);
  });

  function setStatus(msg) {
    statusEl.textContent = msg || "";
  }

  function openDocxLocal() {
    setStatus("Selecting DOCX...");
    var picker = new Windows.Storage.Pickers.FileOpenPicker();
    picker.viewMode = Windows.Storage.Pickers.PickerViewMode.list;
    picker.suggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.documentsLibrary;
    picker.fileTypeFilter.replaceAll([".docx"]);

    picker.pickSingleFileAsync().done(function (file) {
      if (!file) { setStatus("No file selected."); return; }
      setStatus("Loading " + file.name + " ...");

      file.openReadAsync().done(function (stream) {
        var reader = new Windows.Storage.Streams.DataReader(stream);
        reader.loadAsync(stream.size).done(function (/*bytesLoaded*/) {
          var buffer = new Uint8Array(reader.readBytes(stream.size));
          reader.close();

          // Convert DOCX ArrayBuffer to HTML via Mammoth.js
          mammoth.convertToHtml({ arrayBuffer: buffer })
            .then(function (result) {
              docxContainer.innerHTML = result.value; // HTML string
              webViewer.style.display = "none";
              docxContainer.style.display = "block";
              setStatus("Rendered " + file.name);
            })
            .catch(function (err) {
              setStatus("Failed to render DOCX: " + err.message);
            });
        });
      }, function (err) {
        setStatus("Failed to open file: " + err.message);
      });
    });
  }

  function openDocxWeb() {
    setStatus("Select DOCX to view online...");
    var picker = new Windows.Storage.Pickers.FileOpenPicker();
    picker.viewMode = Windows.Storage.Pickers.PickerViewMode.list;
    picker.suggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.documentsLibrary;
    picker.fileTypeFilter.replaceAll([".docx"]);

    picker.pickSingleFileAsync().done(function (file) {
      if (!file) { setStatus("No file selected."); return; }
      // Save to temporary local folder and serve via ms-appdata URL if needed
      file.copyAsync(Windows.Storage.ApplicationData.current.temporaryFolder, file.name, Windows.Storage.NameCollisionOption.replaceExisting)
        .done(function (tempFile) {
          // Some online viewers require HTTP URLs. If you have a backend, upload and generate a viewer URL.
          // Fallback: If the file is accessible via a public URL, set it here:
          var publicUrl = ""; // TODO: supply your uploaded URL
          if (!publicUrl) {
            setStatus("Online viewer requires a public URL. Please integrate upload and set the viewer URL.");
            return;
          }
          var viewerUrl = "https://docs.google.com/gview?embedded=true&url=" + encodeURIComponent(publicUrl);
          webViewer.src = viewerUrl;
          docxContainer.style.display = "none";
          webViewer.style.display = "block";
          setStatus("Opening online viewer...");
        }, function (err) {
          setStatus("Failed to prepare file for online viewing: " + err.message);
        });
    });
  }

})();
