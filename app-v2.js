/* global pdfjsLib, XLSX, fflate */
const state = { pdfFiles: [], xlsxFile: null, pdfRecords: [], workbookRows: [], results: [] };
const $ = (id) => document.getElementById(id);
function setWorkflowStep(step) { document.querySelectorAll(".workflow-step").forEach((item, index) => item.classList.toggle("is-active", index === step - 1)); }
const elements = {
  pdfInput: $("pdfInput"), xlsxInput: $("xlsxInput"), pdfDrop: $("pdfDrop"), xlsxDrop: $("xlsxDrop"),
  pdfState: $("pdfState"), xlsxState: $("xlsxState"), verifyButton: $("verifyButton"), processing: $("processing"),
  processingText: $("processingText"), reviewPanel: $("reviewPanel"), reviewList: $("reviewList"),
  resultsPanel: $("resultsPanel"), reviewMessage: $("reviewMessage"), compareButton: $("compareButton"),
  resetButton: $("resetButton"), uploadHint: $("uploadHint"), downloadResultButton: $("downloadResultButton")
};

if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = "vendor/pdf.worker.min.js";

function bindDropZone(input, zone, kind) {
  input.addEventListener("change", () => setFiles(kind, [...(input.files || [])]));
  ["dragenter", "dragover"].forEach((name) => zone.addEventListener(name, (event) => { event.preventDefault(); zone.classList.add("dragging"); }));
  ["dragleave", "drop"].forEach((name) => zone.addEventListener(name, (event) => { event.preventDefault(); zone.classList.remove("dragging"); }));
  zone.addEventListener("drop", (event) => setFiles(kind, [...event.dataTransfer.files]));
}

async function setFiles(kind, files) {
  if (!files.length) return;
  const zone = kind === "pdf" ? elements.pdfDrop : elements.xlsxDrop;
  const label = kind === "pdf" ? elements.pdfState : elements.xlsxState;
  zone.classList.remove("ready", "error");
  if (kind === "pdf") {
    const pdfs = files.filter((file) => file.type === "application/pdf" || /\.pdf$/i.test(file.name));
    if (!pdfs.length || pdfs.length !== files.length) return showError("pdf", "PDF 파일만 여러 개 선택할 수 있어요.");
    state.pdfFiles = pdfs;
    label.innerHTML = `${pdfs.length}개 PDF 선택됨 <b>✓</b>`;
    zone.classList.add("ready");
  } else {
    const file = files[0];
    if (!/\.(xlsx|xls|xlsm|xlsb)$/i.test(file.name) && !/excel|spreadsheet/i.test(file.type)) return showError("xlsx", "배출인력 엑셀 파일을 선택해주세요.");
    label.innerHTML = `${escapeHtml(file.name)} <b>…</b>`;
    try {
      if (!window.XLSX && !window.fflate) throw new Error("엑셀 판독 모듈을 불러오지 못했습니다.");
      state.workbookRows = await parseWorkbook(file);
      state.xlsxFile = file;
      label.innerHTML = `${escapeHtml(file.name)} · ${state.workbookRows.length}건 인식 <b>✓</b>`;
      zone.classList.add("ready");
    } catch (error) { showError("xlsx", `인식 실패 · ${error.message}`); }
  }
  updateUploadState();
}

function showError(kind, message) {
  const zone = kind === "pdf" ? elements.pdfDrop : elements.xlsxDrop;
  const label = kind === "pdf" ? elements.pdfState : elements.xlsxState;
  if (kind === "pdf") state.pdfFiles = []; else { state.xlsxFile = null; state.workbookRows = []; }
  zone.classList.remove("ready"); zone.classList.add("error"); label.innerHTML = `${escapeHtml(message)} <b>!</b>`;
  updateUploadState();
}

function updateUploadState() {
  const ready = state.pdfFiles.length > 0 && state.xlsxFile;
  elements.verifyButton.setAttribute("aria-disabled", String(!ready));
  if (ready) elements.uploadHint.textContent = `PDF ${state.pdfFiles.length}개와 배출인력 ${state.workbookRows.length}건이 준비됐습니다.`;
  else if (state.xlsxFile) elements.uploadHint.textContent = "엑셀을 인식했습니다. PDF 파일을 하나 이상 추가해주세요.";
  else if (state.pdfFiles.length) elements.uploadHint.textContent = `PDF ${state.pdfFiles.length}개를 선택했습니다. 배출인력 엑셀을 추가해주세요.`;
  else elements.uploadHint.textContent = "PDF 여러 개와 엑셀 파일 하나를 선택해주세요.";
}

bindDropZone(elements.pdfInput, elements.pdfDrop, "pdf");
bindDropZone(elements.xlsxInput, elements.xlsxDrop, "xlsx");

elements.verifyButton.addEventListener("click", async () => {
  if (!state.pdfFiles.length || !state.xlsxFile) { updateUploadState(); elements.uploadHint.classList.add("attention"); setTimeout(() => elements.uploadHint.classList.remove("attention"), 1200); return; }
  if (!window.pdfjsLib || (!window.XLSX && !window.fflate)) return alert("문서 판독 라이브러리를 불러오지 못했습니다. 인터넷 연결 후 새로고침해주세요.");
  elements.processing.hidden = false; elements.reviewPanel.hidden = true; elements.resultsPanel.hidden = true; elements.verifyButton.setAttribute("aria-busy", "true"); setWorkflowStep(2);
  try {
    state.pdfRecords = [];
    for (let index = 0; index < state.pdfFiles.length; index += 1) {
      elements.processingText.textContent = `PDF ${index + 1}/${state.pdfFiles.length} 판독 중 · ${state.pdfFiles[index].name}`;
      const record = await parsePdf(state.pdfFiles[index]);
      state.pdfRecords.push({ ...record, fileName: state.pdfFiles[index].name });
    }
    const results = state.pdfRecords.map((record) => compare(record, state.workbookRows));
    renderBatch(results); elements.resultsPanel.hidden = false; elements.resultsPanel.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (error) { console.error(error); alert(`문서를 읽지 못했습니다.\n${error.message}`); }
  finally { elements.processing.hidden = true; elements.verifyButton.removeAttribute("aria-busy"); }
});

async function parseWorkbook(file) {
  let primaryError;
  if (window.XLSX) {
    try { return await parseWorkbookWithSheetJS(file); }
    catch (error) { primaryError = error; }
  }
  if (window.fflate) {
    try { return await parseWorkbookXml(file); }
    catch (error) { throw new Error(`기본 판독과 XML 판독이 모두 실패했습니다: ${error.message}`); }
  }
  throw primaryError || new Error("엑셀 판독 모듈을 사용할 수 없습니다.");
}

async function parseWorkbookWithSheetJS(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const normalizeHeader = (value) => String(value ?? "").normalize("NFKC").replace(/[\s\n\r]+/g, "").replace(/[()]/g, "");
  const candidates = workbook.SheetNames.map((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true });
    const headerIndex = rows.findIndex((row) => row.some((cell) => {
      const value = normalizeHeader(cell);
      return value === "성명" || value === "이름" || value.includes("참여인력성명");
    }));
    const dataScore = rows.slice(10).filter((row) => String(row[4] || "").trim() && (String(row[26] || "").trim() || row[28])).length;
    return { sheetName, rows, headerIndex, dataScore };
  });
  const selected = candidates.find((item) => item.headerIndex >= 0 && item.sheetName.includes("인력양성"))
    || candidates.find((item) => item.headerIndex >= 0)
    || [...candidates].sort((a, b) => b.dataScore - a.dataScore)[0];
  if (!selected || !selected.rows.length) throw new Error("엑셀에서 배출인력 표를 찾지 못했습니다.");
  const rows = selected.rows;
  const headerIndex = selected.headerIndex >= 0 ? selected.headerIndex : 9;
  const top = (rows[headerIndex] || []).map((cell) => String(cell).trim());
  const sub = (rows[headerIndex + 1] || []).map((cell) => String(cell).trim());
  const find = (headers, names, fallback) => {
    const normalizedNames = names.map(normalizeHeader);
    const index = headers.findIndex((value) => normalizedNames.some((name) => normalizeHeader(value) === name || normalizeHeader(value).includes(name)));
    return index >= 0 ? index : fallback;
  };
  const columns = { name: find(top, ["성명"], 4), career: find(sub, ["구분"], 23), org: find(sub, ["기관명"], 26), position: find(sub, ["직위"], 27), date: find(sub, ["(취업/창업)일자", "취업/창업일자"], 28), duty: find(sub, ["담당업무"], 29), employmentType: find(sub, ["고용형태"], 30) };
  const parsedRows = rows.slice(headerIndex + 2).map((row, index) => ({ rowNumber: headerIndex + index + 3, name: String(row[columns.name] || "").trim(), careerType: String(row[columns.career] || "").trim(), organization: String(row[columns.org] || "").trim(), position: String(row[columns.position] || "").trim(), date: normalizeDate(row[columns.date]), duty: String(row[columns.duty] || "").trim(), employmentType: String(row[columns.employmentType] || "").trim() })).filter((row) => row.name || row.organization || row.date || row.position || row.duty || row.employmentType);
  if (!parsedRows.length) throw new Error("배출인력 데이터가 없습니다. E열 성명, AA열 기관명, AC열 일자를 확인해주세요.");
  return parsedRows;
}

async function parseWorkbookXml(file) {
  const archive = fflate.unzipSync(new Uint8Array(await file.arrayBuffer()));
  const decode = (path) => {
    if (!archive[path]) throw new Error(`${path} 파일이 없습니다.`);
    return new TextDecoder("utf-8").decode(archive[path]);
  };
  const xml = (text) => new DOMParser().parseFromString(text, "application/xml");
  const shared = archive["xl/sharedStrings.xml"]
    ? [...xml(decode("xl/sharedStrings.xml")).getElementsByTagNameNS("*", "si")].map((node) => [...node.getElementsByTagNameNS("*", "t")].map((text) => text.textContent || "").join(""))
    : [];
  const worksheetPaths = Object.keys(archive).filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/.test(path));
  if (!worksheetPaths.length) throw new Error("워크시트 XML을 찾지 못했습니다.");
  const parsedSheets = worksheetPaths.map((path) => {
    const document = xml(decode(path));
    const rows = [];
    for (const rowNode of document.getElementsByTagNameNS("*", "row")) {
      const rowIndex = Number(rowNode.getAttribute("r") || rows.length + 1) - 1;
      const row = rows[rowIndex] || [];
      for (const cell of rowNode.getElementsByTagNameNS("*", "c")) {
        const reference = cell.getAttribute("r") || "A1";
        const letters = reference.match(/[A-Z]+/)?.[0] || "A";
        const column = [...letters].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1;
        const type = cell.getAttribute("t");
        const raw = cell.getElementsByTagNameNS("*", "v")[0]?.textContent || "";
        const inline = [...cell.getElementsByTagNameNS("*", "t")].map((node) => node.textContent || "").join("");
        row[column] = type === "s" ? (shared[Number(raw)] ?? "") : type === "inlineStr" ? inline : raw;
      }
      rows[rowIndex] = row;
    }
    return rows;
  });
  const sheet = parsedSheets.sort((a, b) => scoreIrtcRows(b) - scoreIrtcRows(a))[0];
  const headerIndex = sheet.findIndex((row = []) => row.some((cell) => String(cell || "").replace(/\s/g, "") === "성명"));
  const start = headerIndex >= 0 ? headerIndex + 2 : 11;
  const parsedRows = sheet.slice(start).map((row = [], index) => ({
    rowNumber: start + index + 1,
    name: String(row[4] || "").trim(),
    careerType: String(row[23] || "").trim(),
    organization: String(row[26] || "").trim(),
    position: String(row[27] || "").trim(),
    date: normalizeDate(row[28]),
    duty: String(row[29] || "").trim(),
    employmentType: String(row[30] || "").trim()
  })).filter((row) => row.name || row.organization || row.date || row.position || row.duty || row.employmentType);
  if (!parsedRows.length) throw new Error("E열·AA열·AC열에서 배출인력 데이터를 찾지 못했습니다.");
  return parsedRows;
}

function scoreIrtcRows(rows) {
  return rows.slice(10).filter((row = []) => String(row[4] || "").trim() && (String(row[26] || "").trim() || row[28])).length;
}

async function parsePdf(file) {
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const pageTexts = [];
  let detectedName = "";
  for (let number = 1; number <= pdf.numPages; number += 1) {
    const content = await (await pdf.getPage(number)).getTextContent();
    const items = content.items.map((item) => ({ text: item.str.trim(), x: item.transform[4], y: item.transform[5] })).filter((item) => item.text);
    const nameHeader = items.find((item) => item.text.replace(/\s/g, "") === "성명");
    if (nameHeader && !detectedName) {
      const names = items.filter((item) => /^[가-힣]{2,5}$/.test(item.text.replace(/\s/g, "")) && item.y < nameHeader.y && nameHeader.y - item.y < 55 && Math.abs(item.x - nameHeader.x) < 120);
      names.sort((a, b) => Math.hypot(a.x - nameHeader.x, a.y - nameHeader.y) - Math.hypot(b.x - nameHeader.x, b.y - nameHeader.y));
      detectedName = names[0]?.text.replace(/\s/g, "") || "";
    }
    items.sort((a, b) => Math.abs(b.y - a.y) > 3 ? b.y - a.y : a.x - b.x);
    pageTexts.push(items.map((item) => item.text).join(" "));
  }
  let text = pageTexts.join(" ").replace(/\s+/g, " ");
  let ocrUsed = false;
  let parsed = parseInsuranceText(text, detectedName);
  if (!parsed.row && window.Tesseract) {
    ocrUsed = true;
    const ocrPages = [];
    for (let number = 1; number <= pdf.numPages; number += 1) {
      const page = await pdf.getPage(number);
      const viewport = page.getViewport({ scale: 2.2 });
      const canvas = document.createElement("canvas");
      canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
      await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
      const result = await Tesseract.recognize(canvas, "kor+eng", {
        logger: (message) => {
          if (message.status === "recognizing text") elements.processingText.textContent = `OCR 판독 중 · ${file.name} · ${Math.round((message.progress || 0) * 100)}%`;
        },
        workerPath: "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/worker.min.js",
        langPath: "https://tessdata.projectnaptha.com/4.0.0"
      });
      ocrPages.push(result.data.text || "");
    }
    text = ocrPages.join(" ").replace(/\s+/g, " ");
    parsed = parseInsuranceText(text, detectedName);
  }
  return { name: parsed.name, organization: parsed.row?.[1]?.trim() || "", acquiredDate: parsed.row ? normalizeDate(parsed.row[2]) : "", subscriberType: "직장가입자", documentTypeValid: parsed.documentTypeValid, extractionReliable: Boolean(parsed.documentTypeValid && parsed.row && parsed.name), ocrUsed };
}

function parseInsuranceText(text, knownName = "") {
  const compact = text.replace(/\s+/g, "");
  const documentTypeValid = compact.includes("건강보험자격득실확인서") || compact.includes("건강보험자격득실내역") || compact.includes("건강보험자격득실");
  const name = knownName || compact.match(/성명(?:주민등록번호)?([가-힣]{2,5})(?=\d{6}|[*]{2,})/)?.[1] || compact.match(/성명([가-힣]{2,5})/)?.[1] || "";
  const row = compact.match(/직장가입자(.{2,100}?)(20\d{2}[.\/-]\d{1,2}[.\/-]\d{1,2})/);
  return { name, row, documentTypeValid };
}

function populateReviews() {
  elements.reviewList.innerHTML = state.pdfRecords.map((record, index) => `<article class="review-item ${record.extractionReliable ? "" : "needs-review"}" data-index="${index}"><div class="review-item-head"><strong>${escapeHtml(record.fileName)}</strong><span>${record.extractionReliable ? (record.ocrUsed ? "OCR 판독" : "자동 판독") : "확인 필요"}</span></div><div class="field-grid"><label>성명<input data-field="name" type="text" value="${escapeHtml(record.name)}" placeholder="예: 홍길동" /></label><label>사업장 명칭<input data-field="organization" type="text" value="${escapeHtml(record.organization)}" placeholder="예: 정보통신기획평가원" /></label><label>자격 취득일<input data-field="acquiredDate" type="date" value="${escapeHtml(record.acquiredDate)}" /></label><label>가입자 구분<select data-field="subscriberType"><option>직장가입자</option><option>지역가입자</option><option>피부양자</option></select></label></div></article>`).join("");
  const uncertain = state.pdfRecords.filter((record) => !record.extractionReliable).length;
  elements.reviewMessage.textContent = `PDF ${state.pdfRecords.length}개를 읽었습니다.${uncertain ? ` ${uncertain}개는 자동 판독값을 확인해주세요.` : " 모든 자동 판독값을 확인해주세요."}`;
}

elements.compareButton.addEventListener("click", () => {
  const records = [...elements.reviewList.querySelectorAll(".review-item")].map((item) => ({ fileName: state.pdfRecords[Number(item.dataset.index)].fileName, name: item.querySelector('[data-field="name"]').value.trim(), organization: item.querySelector('[data-field="organization"]').value.trim(), acquiredDate: item.querySelector('[data-field="acquiredDate"]').value, subscriberType: item.querySelector('[data-field="subscriberType"]').value }));
  const incomplete = records.filter((record) => !record.organization || !record.acquiredDate).length;
  if (incomplete) { elements.reviewMessage.textContent = `${incomplete}개 PDF의 사업장 명칭 또는 자격 취득일이 비어 있습니다.`; return; }
  renderBatch(records.map((record) => compare(record, state.workbookRows))); setWorkflowStep(3);
  elements.resultsPanel.hidden = false; elements.resultsPanel.scrollIntoView({ behavior: "smooth", block: "start" });
});

function compare(record, rows) {
  if (!record.documentTypeValid) return { status: "확인 필요", note: "건강보험자격득실확인서가 아님", record, row: null, orgMatch: null, dateMatch: null };
  if (record.subscriberType !== "직장가입자") return { status: "확인 필요", note: "직장가입자 이력이 아님", record, row: null, orgMatch: null, dateMatch: null };
  if (!record.organization || !record.acquiredDate) return { status: "확인 필요", note: "PDF 자동 판독 실패", record, row: null, orgMatch: null, dateMatch: null };
  let candidates = record.name ? rows.filter((row) => normalizeOrg(row.name) === normalizeOrg(record.name)) : rows;
  const nameFound = candidates.length > 0;
  if (record.name && !nameFound) {
    const exactEvidence = rows.filter((row) => normalizeOrg(row.organization) === normalizeOrg(record.organization) && normalizeDate(row.date) === normalizeDate(record.acquiredDate));
    if (exactEvidence.length === 1) return { status: "이름 확인 필요", note: "기관명·일자는 일치하지만 PDF 성명 판독을 확인해주세요", record, row: exactEvidence[0], orgMatch: true, dateMatch: true };
    return { status: "엑셀 미등록", note: "PDF에서 발견했지만 엑셀에 같은 성명이 없음", record, row: null, orgMatch: null, dateMatch: null };
  }
  if (!candidates.length) candidates = rows;
  const scored = candidates.map((row) => ({ row, orgMatch: normalizeOrg(row.organization) === normalizeOrg(record.organization), dateMatch: normalizeDate(row.date) === normalizeDate(record.acquiredDate) })).map((entry) => ({ ...entry, score: Number(entry.orgMatch) + Number(entry.dateMatch) })).sort((a, b) => b.score - a.score);
  if (!scored.length) return { status: "확인 필요", note: "엑셀 비교 대상 없음", record, row: null, orgMatch: null, dateMatch: null };
  const best = scored[0];
  if (record.name && !nameFound && best.score < 2) return { ...best, status: "확인 필요", note: "엑셀에서 같은 성명을 찾지 못함", record };
  if (best.row.careerType === "창업") return { ...best, status: "확인 필요", note: "창업은 건강보험 이력만으로 입증 불가", record };
  const status = best.orgMatch && best.dateMatch ? "일치" : best.orgMatch ? "날짜 불일치" : best.dateMatch ? "기관명 불일치" : "기관명·날짜 불일치";
  return { ...best, status, note: "", record };
}

function renderBatch(results) {
  state.results = results;
  const matched = results.filter((result) => result.status === "일치").length;
  const linkedRows = new Set(results.filter((result) => result.row).map((result) => result.row.rowNumber));
  const uncoveredRows = state.workbookRows.filter((row) => !linkedRows.has(row.rowNumber));
  const allMatched = matched === results.length;
  $("resultIcon").textContent = allMatched ? "✓" : "!";
  $("resultTitle").textContent = allMatched ? `${results.length}개 PDF가 모두 일치합니다` : `${matched}/${results.length}개 PDF가 일치합니다`;
  $("resultDescription").textContent = allMatched ? "모든 건강보험 이력이 배출인력 엑셀과 일치합니다." : "불일치 또는 확인 필요 항목을 아래 표에서 확인해주세요.";
  $("totalBadge").textContent = `${results.length}건`; $("matchedCount").textContent = `${matched}건`; $("mismatchCount").textContent = `${results.length - matched}건`;
  $("excelBadge").textContent = `${state.workbookRows.length}건`; $("coveredCount").textContent = `${linkedRows.size}건`; $("uncoveredCount").textContent = `${Math.max(0, state.workbookRows.length - linkedRows.size)}건`;
  const pdfRows = results.map((result) => { const row = result.row || {}; const orgStatus = result.orgMatch === true ? "일치" : result.orgMatch === false ? "불일치" : "확인 필요"; const dateStatus = result.dateMatch === true ? "일치" : result.dateMatch === false ? "불일치" : "확인 필요"; const pdfOnly = result.status === "엑셀 미등록"; return `<tr class="${pdfOnly ? "pdf-only-row" : ""}"><td>${escapeHtml(row.name || result.record.name || "-")}<small class="table-file">${escapeHtml(result.record.fileName)}${pdfOnly ? " · PDF에서 발견" : ""}</small></td><td>${escapeHtml(row.organization || "-")}</td><td>${escapeHtml(result.record.organization || "-")}</td><td class="status-${result.orgMatch ? "ok" : "check"}">${pdfOnly ? "엑셀 미등록" : orgStatus}</td><td>${escapeHtml(row.position || "-")}</td><td>${escapeHtml(row.date || "-")}</td><td>${escapeHtml(result.record.acquiredDate || "-")}</td><td class="status-${result.dateMatch ? "ok" : "check"}">${pdfOnly ? "엑셀 미등록" : dateStatus}</td><td>${escapeHtml(row.duty || "-")}</td><td>${escapeHtml(row.employmentType || "-")}</td><td class="status-${result.status === "일치" ? "ok" : "check"}">${escapeHtml(result.status)}</td></tr>`; }).join("");
  const missingRows = uncoveredRows.map((row) => `<tr><td>${escapeHtml(row.name || "-")}<small class="table-file">엑셀 ${row.rowNumber}행</small></td><td>${escapeHtml(row.organization || "-")}</td><td>-</td><td class="status-check">PDF 없음</td><td>${escapeHtml(row.position || "-")}</td><td>${escapeHtml(row.date || "-")}</td><td>-</td><td class="status-check">PDF 없음</td><td>${escapeHtml(row.duty || "-")}</td><td>${escapeHtml(row.employmentType || "-")}</td><td class="status-check">PDF 없음</td></tr>`).join("");
  $("resultBody").innerHTML = pdfRows + missingRows;
}

elements.downloadResultButton.addEventListener("click", () => {
  if (!window.XLSX || !state.results.length) return;
  const rows = state.results.map((result) => {
    const row = result.row || {};
    return { "성명": row.name || result.record.name || "", "AA 기관명": row.organization || "", "PDF 사업장명": result.record.organization || "", "기관 판정": result.orgMatch === true ? "일치" : result.orgMatch === false ? "불일치" : result.status === "엑셀 미등록" ? "엑셀 미등록" : "확인 필요", "AB 직위": row.position || "", "AC 취업/창업일자": row.date || "", "PDF 취득일": result.record.acquiredDate || "", "일자 판정": result.dateMatch === true ? "일치" : result.dateMatch === false ? "불일치" : result.status === "엑셀 미등록" ? "엑셀 미등록" : "확인 필요", "AD 담당업무": row.duty || "", "AE 고용형태": row.employmentType || "", "종합 판정": result.status, "비고": result.note || "", "PDF 파일": result.record.fileName || "" };
  });
  state.workbookRows.filter((row) => !new Set(state.results.filter((result) => result.row).map((result) => result.row.rowNumber)).has(row.rowNumber)).forEach((row) => rows.push({ "성명": row.name || "", "AA 기관명": row.organization || "", "PDF 사업장명": "", "기관 판정": "PDF 없음", "AB 직위": row.position || "", "AC 취업/창업일자": row.date || "", "PDF 취득일": "", "일자 판정": "PDF 없음", "AD 담당업무": row.duty || "", "AE 고용형태": row.employmentType || "", "종합 판정": "PDF 없음", "비고": "엑셀에는 있으나 연결된 PDF가 없습니다.", "PDF 파일": "" }));
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, "검증 결과");
  XLSX.writeFile(workbook, `건강보험_취업일자_검증결과_${new Date().toISOString().slice(0, 10)}.xlsx`);
});

function normalizeOrg(value) { return String(value || "").normalize("NFKC").toLowerCase().trim().replace(/^(?:주식회사|\(주\)|㈜)|(?:주식회사|\(주\)|㈜)$/g, "").replace(/[\s\-_.·,()]/g, ""); }
function normalizeDate(value) { if (!value) return ""; if (value instanceof Date && !Number.isNaN(value.valueOf())) return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`; if (typeof value === "number" || /^\d+(\.\d+)?$/.test(String(value))) { const serial = Number(value); if (serial > 20000 && serial < 100000) { const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000); return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`; } } const raw = String(value).trim().replace(/[./]/g, "-").replace(/년/g, "-").replace(/월/g, "-").replace(/일/g, "").replace(/\s/g, ""); const match = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/) || raw.match(/^(\d{4})(\d{2})(\d{2})$/); return match ? `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}` : ""; }
function escapeHtml(value) { return String(value).replace(/[&<>"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[char])); }

elements.resetButton.addEventListener("click", () => {
  state.pdfFiles = []; state.xlsxFile = null; state.pdfRecords = []; state.workbookRows = []; state.results = []; setWorkflowStep(1);
  elements.pdfInput.value = ""; elements.xlsxInput.value = ""; elements.pdfDrop.classList.remove("ready", "error"); elements.xlsxDrop.classList.remove("ready", "error");
  elements.pdfState.innerHTML = "파일 선택 <b>+</b>"; elements.xlsxState.innerHTML = "파일 선택 <b>+</b>"; elements.reviewPanel.hidden = true; elements.resultsPanel.hidden = true; updateUploadState(); $("checker").scrollIntoView({ behavior: "smooth" });
});
