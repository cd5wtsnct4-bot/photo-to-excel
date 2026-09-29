const $ = (id) => document.getElementById(id);
const fields = ["documentDate","supplier","documentNumber","description","category","subtotal","tax","total","currency","paymentMethod","notes"];
let photoFile = null;
let workbookFile = null;

function showToast(message) {
  const toast = $("toast");
  toast.textContent = message;
  toast.classList.remove("hidden");
  setTimeout(() => toast.classList.add("hidden"), 3200);
}

function setPhoto(file) {
  if (!file || !file.type.startsWith("image/")) return showToast("Choose an image file.");
  photoFile = file;
  $("preview").src = URL.createObjectURL(file);
  $("dropzone").classList.add("hidden");
  $("previewWrap").classList.remove("hidden");
  $("extractButton").disabled = false;
}

$("photoInput").addEventListener("change", (e) => setPhoto(e.target.files[0]));
$("replacePhoto").addEventListener("click", () => $("photoInput").click());
for (const eventName of ["dragenter","dragover"]) $("dropzone").addEventListener(eventName, (e) => { e.preventDefault(); $("dropzone").classList.add("dragging"); });
for (const eventName of ["dragleave","drop"]) $("dropzone").addEventListener(eventName, (e) => { e.preventDefault(); $("dropzone").classList.remove("dragging"); });
$("dropzone").addEventListener("drop", (e) => setPhoto(e.dataTransfer.files[0]));

function amountAfter(label, text) {
  const rx = new RegExp(`(?:${label})\\s*[:\\-]?\\s*(?:R|ZAR|USD|EUR|GBP|\$|€|£)?\\s*([0-9][0-9 .,]*[.,][0-9]{2}|[0-9]+)`, "i");
  const match = text.match(rx);
  return match ? match[1].replace(/\s/g, "").replace(/,(?=\d{3}\b)/g, "").replace(",", ".") : "";
}

function parseText(text) {
  const lines = text.split(/\r?\n/).map((v) => v.trim()).filter(Boolean);
  const dateMatch = text.match(/\b(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})\b/) || text.match(/\b(\d{1,2})[-\/.](\d{1,2})[-\/.](20\d{2})\b/);
  let isoDate = "";
  if (dateMatch) {
    const yearFirst = dateMatch[1].length === 4;
    const year = yearFirst ? dateMatch[1] : dateMatch[3];
    const month = yearFirst ? dateMatch[2] : dateMatch[2];
    const day = yearFirst ? dateMatch[3] : dateMatch[1];
    isoDate = `${year}-${month.padStart(2,"0")}-${day.padStart(2,"0")}`;
  }
  const numberMatch = text.match(/(?:invoice|receipt|document|ref(?:erence)?)\s*(?:no\.?|number|#)?\s*[:#-]?\s*([A-Z0-9][A-Z0-9\-\/]{2,})/i);
  const currency = /\bZAR\b|R\s?\d/.test(text) ? "ZAR" : /\bUSD\b|\$/.test(text) ? "USD" : /\bEUR\b|€/.test(text) ? "EUR" : /\bGBP\b|£/.test(text) ? "GBP" : "";
  const card = text.match(/\b(visa|mastercard|amex|cash|eft|bank transfer)\b/i);
  const total = amountAfter("grand total|amount due|total", text);
  const tax = amountAfter("vat|tax", text);
  const subtotal = amountAfter("subtotal|sub total|net amount", text);
  return {
    documentDate: isoDate,
    supplier: lines.find((line) => /[A-Za-z]{3}/.test(line) && line.length < 55 && !/invoice|receipt|tax|date/i.test(line)) || "",
    documentNumber: numberMatch?.[1] || "",
    description: lines.find((line) => /description|service|item/i.test(line))?.replace(/description\s*:?/i, "").trim() || "",
    subtotal, tax, total, currency,
    paymentMethod: card ? card[1].toUpperCase() : "",
  };
}

$("extractButton").addEventListener("click", async () => {
  if (!photoFile || !window.Tesseract) return showToast("The text reader is still loading. Try again in a moment.");
  $("progress").classList.remove("hidden");
  $("extractButton").disabled = true;
  try {
    const result = await Tesseract.recognize(photoFile, "eng", { logger: ({ status, progress }) => {
      const pct = Math.round((progress || 0) * 100);
      $("progressBar").style.width = `${Math.max(5,pct)}%`;
      $("progressText").textContent = `${status || "Reading photograph"} ${pct ? `· ${pct}%` : ""}`;
    }});
    const text = result.data.text.trim();
    $("rawText").value = text;
    const parsed = parseText(text);
    for (const [key,value] of Object.entries(parsed)) if ($(key) && !$(key).value) $(key).value = value;
    showToast("Photograph read. Please review the fields.");
  } catch (error) {
    console.error(error);
    showToast("The photograph could not be read. You can still enter the values manually.");
  } finally {
    $("extractButton").disabled = false;
    $("progressText").textContent = "Complete";
    $("progressBar").style.width = "100%";
  }
});

$("workbookInput").addEventListener("change", (e) => {
  workbookFile = e.target.files[0] || null;
  $("workbookStatus").textContent = workbookFile ? `Selected: ${workbookFile.name}` : "No workbook selected. A new workbook will be created.";
});

function excelDate(value) {
  if (!value) return "";
  const [year,month,day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
}

$("appendButton").addEventListener("click", async () => {
  if (!window.XLSX) return showToast("The Excel writer is still loading. Try again in a moment.");
  try {
    let workbook;
    if (workbookFile) workbook = XLSX.read(await workbookFile.arrayBuffer(), { type:"array", cellDates:true });
    else workbook = XLSX.utils.book_new();
    const sheetName = workbook.SheetNames.includes("Photo records") ? "Photo records" : (workbook.SheetNames[0] || "Photo records");
    let sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      sheet = XLSX.utils.aoa_to_sheet([["Captured at","Document date","Supplier / person","Document number","Description","Category","Subtotal","Tax","Total","Currency","Payment method","Notes","Source file","Extracted text"]]);
      XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
    }
    const row = [new Date(), excelDate($("documentDate").value), $("supplier").value, $("documentNumber").value, $("description").value, $("category").value, Number($("subtotal").value)||null, Number($("tax").value)||null, Number($("total").value)||null, $("currency").value.toUpperCase(), $("paymentMethod").value, $("notes").value, photoFile?.name||"", $("rawText").value];
    XLSX.utils.sheet_add_aoa(sheet, [row], { origin:-1, cellDates:true });
    const filename = workbookFile?.name?.replace(/\.xls[x]?$/i, "") || "photo-records";
    XLSX.writeFile(workbook, `${filename}-updated.xlsx`, { compression:true, cellDates:true });
    showToast("Excel workbook downloaded with the new row.");
  } catch (error) {
    console.error(error);
    showToast("Could not update that workbook. Try the blank template.");
  }
});
