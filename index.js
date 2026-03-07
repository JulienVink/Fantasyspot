/**
 * Eco-growth Discovery Project
 * Entry point for the Node.js project.
 */

Office.onReady((info) => {
  // Initialize for PowerPoint or when testing in a browser
  if (info.host === Office.HostType.PowerPoint || !info.host) {
    console.log("Eco-growth Discovery initialized.");

    // Retrieve and display initials
    const initialsDisplay = document.getElementById("initials-display");
    if (initialsDisplay) {
      // Default initials from the project template metadata (Julien Vink)
      let initials = "JV";

      // Note: Office.context.userProfile is primarily for Outlook.
      // In PowerPoint, user information is not directly exposed via simple properties.
      // We use "JV" as the primary identifier for this specialized add-in.

      initialsDisplay.textContent = initials;
    }

    // Attach event listener to the "Read Active File" button
    const readBtn = document.getElementById("read-files-btn");
    if (readBtn) {
      readBtn.onclick = readActiveFile;
    }
  }
});

/**
 * Promisified wrapper for Office.context.document.getFileAsync.
 */
function getFileAsync(fileType, options) {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, options, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

/**
 * Promisified wrapper for file.getSliceAsync.
 */
function getSliceAsync(file, sliceIndex) {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(sliceIndex, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

/**
 * Promisified wrapper for file.closeAsync.
 */
function closeAsync(file) {
  return new Promise((resolve) => {
    file.closeAsync(() => {
      resolve();
    });
  });
}

/**
 * Reads the active PowerPoint file as a compressed byte stream.
 */
async function readActiveFile() {
  const status = document.getElementById("status");
  if (status) status.textContent = "Reading file...";

  if (typeof Office === "undefined" || !Office.context || !Office.context.document) {
    const errorMsg = "Office.js is not loaded or this is not an Office host.";
    console.error(errorMsg);
    if (status) status.textContent = errorMsg;
    return;
  }

  let file = null;
  try {
    // 1. Get the file handle
    file = await getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 });
    const sliceCount = file.sliceCount;
    const fileSize = file.size;

    if (status) status.textContent = `File size: ${fileSize} bytes. Reading ${sliceCount} slices...`;

    // 2. Pre-allocate Uint8Array for the file content
    const fileData = new Uint8Array(fileSize);
    let offset = 0;

    // 3. Read slices sequentially
    for (let i = 0; i < sliceCount; i++) {
      const slice = await getSliceAsync(file, i);
      fileData.set(slice.data, offset);
      offset += slice.data.length;

      if (status) {
        status.textContent = `Reading progress: ${Math.round(((i + 1) / sliceCount) * 100)}%`;
      }
    }

    if (status) {
      status.textContent = `Successfully read active file: ${fileSize} bytes.`;
    }
    console.log(`Read ${fileSize} bytes from the active presentation.`);
  } catch (error) {
    const errorMsg = `Error reading file: ${error.message || error}`;
    console.error(errorMsg);
    if (status) status.textContent = errorMsg;
  } finally {
    // 4. Always close the file handle
    if (file) {
      await closeAsync(file);
    }
  }
}
