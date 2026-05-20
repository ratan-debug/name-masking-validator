const form = document.getElementById("uploadForm");
const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const chooseButton = document.getElementById("chooseButton");
const processButton = document.getElementById("processButton");
const fileName = document.getElementById("fileName");
const summaryPanel = document.getElementById("summaryPanel");
const states = {
  idle: document.getElementById("idleState"),
  loading: document.getElementById("loadingState"),
  success: document.getElementById("successState"),
  error: document.getElementById("errorState"),
};

chooseButton.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", () => {
  updateSelectedFile();
  showState("idle");
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("drag-over");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("drag-over");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("drag-over");
  if (event.dataTransfer.files.length > 0) {
    fileInput.files = event.dataTransfer.files;
    updateSelectedFile();
    showState("idle");
  }
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!fileInput.files.length) {
    showError("Choose an Excel file before processing.");
    return;
  }

  const formData = new FormData();
  formData.append("file", fileInput.files[0]);

  setBusy(true);
  showState("loading");
  summaryPanel.classList.add("hidden");

  try {
    const response = await fetch("/process", {
      method: "POST",
      body: formData,
    });
    const payload = await response.json();

    if (!response.ok) {
      throw new Error(payload.error || "Processing failed.");
    }

    renderSuccess(payload);
  } catch (error) {
    showError(toUserMessage(error));
  } finally {
    setBusy(false);
  }
});

function updateSelectedFile() {
  const file = fileInput.files[0];
  fileName.textContent = file ? file.name : "No file selected";
  processButton.disabled = !file;
}

function setBusy(isBusy) {
  processButton.disabled = isBusy || !fileInput.files.length;
  chooseButton.disabled = isBusy;
}

function showState(activeState) {
  Object.entries(states).forEach(([name, element]) => {
    element.classList.toggle("hidden", name !== activeState);
  });
}

function showError(message) {
  document.getElementById("errorMessage").textContent = message;
  showState("error");
}

function toUserMessage(error) {
  if (error instanceof TypeError && error.message.toLowerCase().includes("fetch")) {
    return "Backend server reachable nahi hai. Flask app ko python app.py se chalu rakhein, phir page refresh karke retry karein.";
  }
  return error.message;
}

function renderSuccess(payload) {
  document.getElementById("successMessage").textContent = payload.message;
  const downloadButton = document.getElementById("downloadButton");
  downloadButton.href = payload.download_url;
  downloadButton.setAttribute("download", payload.download_name || "processed.xlsx");

  const summary = payload.summary || {};
  document.getElementById("totalRows").textContent = formatNumber(summary.total_rows);
  document.getElementById("fullMatches").textContent = formatNumber(summary.full_matches);
  document.getElementById("partialMatches").textContent = formatNumber(summary.partial_matches);
  document.getElementById("noMatches").textContent = formatNumber(summary.no_matches);
  document.getElementById("averageConfidence").textContent = `${formatNumber(summary.average_confidence)}%`;

  summaryPanel.classList.remove("hidden");
  showState("success");
}

function formatNumber(value) {
  const number = Number(value || 0);
  return Number.isInteger(number) ? number.toLocaleString() : number.toFixed(2);
}
