/* global Office, Word */

function setStatus(message, type = "") {
  const status = document.getElementById("status");
  status.textContent = message;
  status.className = "status-msg show " + type;
  if (status.timeoutId) clearTimeout(status.timeoutId);
  if (message && type !== "loading") {
    status.timeoutId = setTimeout(() => { status.className = "status-msg " + type; }, 5000);
  }
}

function getOptions() {
  return {
    spaces: document.getElementById("optSpaces").checked,
    blankLines: document.getElementById("optBlankLines").checked,
    clearBg: document.getElementById("optClearBg").checked,
    standardFont: document.getElementById("optStandardFont").checked,
    tableFix: document.getElementById("optTableFix").checked,
    imageFix: document.getElementById("optImageFix").checked
  };
}

function getScope() {
  const checked = document.querySelector("input[name='scope']:checked");
  return checked ? checked.value : "selection";
}

async function getRange(context, allowEmpty = false) {
  let scope = getScope();
  if (scope === "document") return context.document.body.getRange();
  let range = context.document.getSelection();
  if (!allowEmpty) {
    range.load("text"); await context.sync();
    if (!range.text || range.text.trim().length === 0) throw new Error("Vui lòng bôi đen nội dung cần thao tác.");
  }
  return range;
}

// Xử lý báo lỗi API qua tiếng Việt để dễ hiểu
function handleGeminiError(errorMsg) {
  if (errorMsg.includes("exceeded your current quota") || errorMsg.includes("Rate limit") || errorMsg.includes("429")) {
    return "API Key đã hết hạn mức (Quota Exceeded) hoặc bạn gọi quá nhanh trong thời gian ngắn. Vui lòng chờ vài phút rồi thử lại.";
  }
  if (errorMsg.includes("API key not valid")) {
    return "API Key không hợp lệ. Hãy kiểm tra lại Key đã nhập!";
  }
  return errorMsg;
}

// Nén ảnh trước khi gửi cho Gemini (tránh lỗi Payload Too Large hoặc hao hụt Quota API của các thao tác ảnh)
function resizeImage(base64Str, maxWidth = 1200) {
  return new Promise((resolve, reject) => {
    let img = new Image();
    img.onload = () => {
      let canvas = document.createElement("canvas");
      let width = img.width, height = img.height;
      if (width > maxWidth) {
        height = Math.round((height * maxWidth) / width);
        width = maxWidth;
      }
      canvas.width = width; canvas.height = height;
      let ctx = canvas.getContext("2d");
      // Cố định nền trắng cho ảnh trong suốt (hỗ trợ ảnh PNG)
      ctx.fillStyle = "#ffffff"; ctx.fillRect(0, 0, width, height);
      ctx.drawImage(img, 0, 0, width, height);
      // Nén 75% chất lượng để tiết kiệm tối đa dung lượng Base64
      resolve(canvas.toDataURL("image/jpeg", 0.75));
    };
    img.onerror = () => reject(new Error("Lỗi đọc dữ liệu ảnh."));
    img.src = base64Str.startsWith("data:image") ? base64Str : "data:image/jpeg;base64," + base64Str;
  });
}

function cleanLatexForMathType(text) {
  if (!text) return text;
  text = text.replace(/\\triangle(?![a-zA-Z])/g, "\\Delta");
  text = text.replace(/\\angle\s*([A-Z0-9]+)/g, "\\widehat{$1}");
  text = text.replace(/\\cong/g, "=");
  text = text.replace(/(\${1,2})([^\$]+)\1/g, function(match, d, formula) {
    formula = formula.replace(/(\d)\.(\d)/g, "$1,$2");
    let newF = formula.replace(/\\(?:text|mbox|mathrm|textbf)\s*\{([^}]+)\}/g, ` ${d} $1 ${d} `);
    let escapedD = d.replace(/\$/g, "\\$");
    let emptyBlockRegex = new RegExp(escapedD + "\\s*" + escapedD, "g");
    return `${d}${newF}${d}`.replace(emptyBlockRegex, "");
  });
  return text;
}

function formatTextToHtml(text) {
  if (!text) return "";
  return text.replace(/>\s+</g, "><").replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>").replace(/\n/g, "<br>");
}

function extractJsonFromText(text) {
  let start = text.indexOf('{');
  let end = text.lastIndexOf('}');
  return (start !== -1 && end !== -1 && end > start) ? text.substring(start, end + 1) : text;
}

/* =========================================================
   CÁC HÀM CÔNG CỤ DỌN DẸP / FORMAT CHUNG
   ========================================================= */

async function toolFixLatex() {
  setStatus("Đang dọn dẹp lỗi LaTeX...", "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context);
      let oldText = range.text;
      if (!oldText) return;
      let newText = cleanLatexForMathType(oldText);
      if (newText !== oldText) range.insertText(newText, Word.InsertLocation.replace);
      await context.sync();
    });
    setStatus("Đã khắc phục lỗi Toán học thành công.", "success");
  } catch(e) { setStatus(e.message, "error"); }
}

async function toolRemoveLinks() {
  setStatus("Đang xóa Link...", "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context);
      range.font.underline = "None"; range.font.color = "#000000";
      await context.sync();
    });
    setStatus("Đã xóa giao diện Link thành công.", "success");
  } catch(e) { setStatus(e.message, "error"); }
}

async function toolSwapNumbers() {
  setStatus("Đang đổi chuẩn số...", "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context);
      let results = range.search("[0-9,\.]{3,}", { matchWildcards: true });
      results.load("items, text"); await context.sync();
      for (let i = 0; i < results.items.length; i++) {
        let txt = results.items[i].text;
        if (txt.includes(',') || txt.includes('.')) {
          let newTxt = txt, oldTxt;
          do {
            oldTxt = newTxt;
            newTxt = newTxt.replace(/(\d),(\d)/g, "$1TEMP$2").replace(/(\d)\.(\d)/g, "$1COMMA$2").replace(/TEMP/g, ".").replace(/COMMA/g, ",");
          } while (newTxt !== oldTxt);
          if (newTxt !== txt) results.items[i].insertText(newTxt, Word.InsertLocation.replace);
        }
      }
      await context.sync();
    });
    setStatus("Đã đổi định dạng số thành công.", "success");
  } catch(e) { setStatus(e.message, "error"); }
}

async function toolChangeCase(toUpper) {
  setStatus("Đang chuyển đổi hoa/thường...", "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context);
      if (toUpper) range.font.allCaps = true;
      else {
        range.font.allCaps = false; range.load("text"); await context.sync();
        range.insertText(range.text.toLowerCase(), Word.InsertLocation.replace);
      }
      await context.sync();
    });
    setStatus("Đã chuyển đổi chữ thành công.", "success");
  } catch(e) { setStatus(e.message, "error"); }
}

async function toolAddTestLines() {
  const numLines = parseInt(document.getElementById("numTestLines").value) || 3;
  setStatus(`Đang tạo ${numLines} dòng kẻ bài tập...`, "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context, true);
      let rowsHtml = "";
      for (let i = 0; i < numLines; i++) rowsHtml += `<tr><td style="border-bottom: 1.5px dotted black; height: 35px;"></td></tr>\n`;
      range.insertHtml(`<table style="width:100%; border-collapse:collapse; margin-top: 10px; margin-bottom: 10px;">${rowsHtml}</table><br>`, Word.InsertLocation.after);
      await context.sync();
    });
    setStatus(`Đã thêm ${numLines} dòng chấm để làm bài tập.`, "success");
  } catch(e) { setStatus(e.message, "error"); }
}

async function processAutoClean(context, range, options) {
  if (options.clearBg) {
    range.font.highlightColor = null; let paras = range.paragraphs; paras.load("items"); await context.sync();
    for (let i = 0; i < paras.items.length; i++) paras.items[i].font.highlightColor = null; 
  }
  if (options.standardFont) { range.font.name = "Times New Roman"; range.font.size = 14; }
  if (options.tableFix) {
    let tables = range.tables; tables.load("items"); await context.sync();
    for (let i = 0; i < tables.items.length; i++) {
        let table = tables.items[i]; table.autoFitWindow();
        let borders = table.borders; borders.load("outsideBorderColor, insideBorderColor"); await context.sync();
        borders.outsideBorderColor = "#000000"; borders.outsideBorderStyle = "Single"; borders.outsideBorderWidth = "pt05";
        borders.insideBorderColor = "#000000"; borders.insideBorderStyle = "Single"; borders.insideBorderWidth = "pt05";
    }
  }
  if (options.imageFix) {
    let pics = range.inlinePictures; pics.load("items"); await context.sync();
    for (let i = 0; i < pics.items.length; i++) pics.items[i].lockAspectRatio = true;
  }
  if (options.blankLines) {
    let soft = range.search("^l", { matchWildcards: false }); soft.load("items"); await context.sync();
    for (let i = 0; i < soft.items.length; i++) soft.items[i].insertText(" ", Word.InsertLocation.replace); await context.sync();
    let paras = range.paragraphs; paras.load("items, text"); await context.sync();
    for (let i = paras.items.length - 1; i >= 0; i--) if (paras.items[i].text.trim() === "") paras.items[i].delete();
    await context.sync();
  }
  if (options.spaces) {
    let spaces = range.search(" {2,}", { matchWildcards: true }); spaces.load("items"); await context.sync();
    for (let i = 0; i < spaces.items.length; i++) spaces.items[i].insertText(" ", Word.InsertLocation.replace); await context.sync();
    const puncts = [",", ".", ";", ":", "!", "?"]; let punctTasks = [];
    for (const p of puncts) { let punctSearch = range.search(" " + p, { matchWildcards: false }); punctSearch.load("items"); punctTasks.push({ search: punctSearch, replaceText: p }); }
    await context.sync(); 
    for (let t of punctTasks) for (let i = 0; i < t.search.items.length; i++) t.search.items[i].insertText(t.replaceText, Word.InsertLocation.replace);
  }
}

async function runAutoClean() {
  const options = getOptions(); setStatus("Đang xử lý dọn dẹp toàn diện...", "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context);
      await processAutoClean(context, range, options);
      await context.sync();
    });
    setStatus("Đã chuẩn hoá thành công.", "success");
    setTimeout(() => { if(document.getElementById("status").classList.contains("success")) setStatus("", ""); }, 4000);
  } catch (error) { setStatus(error.message, "error"); }
}

async function toolFormatMCQ() {
  setStatus("Đang dồn đáp án...", "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context, true);
      const targets = ["B", "C", "D", "b", "c", "d"]; const suffixes = [".", ")"]; const lineBreaks = ["^p", "^l"]; const spaceVars = ["", " ", "  "];
      let count = 0, searchTasks = [];
      for (const lb of lineBreaks) for (const t of targets) for (const s of suffixes) for (const sp of spaceVars) {
        let results = range.search(`${lb}${sp}${t}${s}`, { matchCase: true, matchWildcards: false });
        results.load("items"); searchTasks.push({ results: results, replaceStr: `\t\t${t}${s}` });
      }
      await context.sync();
      for (let task of searchTasks) {
        for (let i = 0; i < task.results.items.length; i++) {
          task.results.items[i].insertText(task.replaceStr, Word.InsertLocation.replace); count++;
        }
      }
      if(count > 0) setStatus(`Đã dồn ngang ${count} khối đáp án thành công.`, "success");
      else setStatus(`Không tìm thấy cấu trúc đáp án B, C, D nằm rời rạc.`, "error");
    });
  } catch(e) { setStatus(e.message, "error"); }
}

async function toolFindReplace() {
  const findText = document.getElementById("findText").value;
  const replaceText = document.getElementById("replaceText").value;
  if (!findText) return setStatus("Vui lòng nhập từ cần tìm.", "error");
  setStatus("Đang tìm và thay thế...", "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context, true);
      let searchResults = range.search(findText, { matchCase: false, matchWholeWord: false });
      searchResults.load("items"); await context.sync(); let count = searchResults.items.length;
      for (let i = 0; i < count; i++) searchResults.items[i].insertText(replaceText, Word.InsertLocation.replace);
      await context.sync();
      const summaryDiv = document.getElementById("findReplaceSummary");
      if (count > 0) {
        setStatus(`Đã thay thế ${count} vị trí thành công.`, "success");
        summaryDiv.style.display = "block"; summaryDiv.style.borderLeftColor = "var(--success)"; summaryDiv.style.backgroundColor = "rgba(46, 204, 113, 0.1)";
        summaryDiv.innerHTML = `Đã tìm và thay thế thành công <strong>${count}</strong> từ "<em>${findText}</em>".`;
      } else {
        setStatus("Không tìm thấy.", "error");
        summaryDiv.style.display = "block"; summaryDiv.style.borderLeftColor = "var(--error)"; summaryDiv.style.backgroundColor = "rgba(239, 68, 68, 0.1)";
        summaryDiv.innerHTML = `Không tìm thấy từ "<em>${findText}</em>".`;
      }
    });
  } catch(e) { setStatus(e.message, "error"); }
}

/* =========================================================
   CÁC HÀM GỌI GEMINI AI
   ========================================================= */

async function callGeminiApi() {
  const btn = document.getElementById("btnAiWrite");
  const prompt = document.getElementById("aiPrompt").value.trim();
  const apiKey = document.getElementById("geminiApiKey").value.trim();
  
  if (!prompt) return setStatus("Vui lòng nhập yêu cầu.", "error");
  if (!apiKey) return setStatus("Vui lòng nhập API Key.", "error");
  localStorage.setItem("geminiApiKey", apiKey);

  setStatus("AI đang suy nghĩ...", "loading");
  const summaryDiv = document.getElementById("aiSummary");
  summaryDiv.style.display = "block"; summaryDiv.style.borderLeftColor = "var(--primary)"; summaryDiv.style.backgroundColor = "rgba(99, 102, 241, 0.1)";
  summaryDiv.innerHTML = "<strong>Đang xử lý:</strong> AI đang suy nghĩ... ⏳";
  btn.disabled = true; const originalBtnText = btn.innerHTML; btn.innerHTML = "⏳ Đang xử lý..."; btn.style.opacity = "0.7"; btn.style.cursor = "not-allowed";
  
  try {
    let isEditing = false;
    let finalPrompt = `Yêu cầu: ${prompt}\n\n[QUY TẮC]\nƯu tiên số nguyên, phân số tối giản. Cấm Tiếng Việt trong LaTeX. Ký hiệu chuẩn: \\Delta, \\widehat{}, \\sim, =, \\cdot, \\parallel, \\perp. Dùng thẻ HTML <table>.\nTRẢ VỀ JSON:\n{"summary": "Một câu đánh giá ngắn.", "text": "VĂN BẢN KẾT QUẢ ĐỊNH DẠNG HTML"}`;
    
    await Word.run(async (context) => {
      let range = context.document.getSelection(); range.load("text"); await context.sync();
      if (range.text && range.text.trim().length > 0) {
        isEditing = true;
        finalPrompt = `Văn bản đang chọn:\n"""\n${range.text}\n"""\nYêu cầu: ${prompt}\nTRẢ VỀ JSON:\n{"summary": "...", "text": "...", "edits": [{"old": "gốc", "new": "sửa"}]}`;
      }
    });

    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ contents: [{ parts: [{ text: finalPrompt }] }], generationConfig: { responseMimeType: "application/json" } })
    });
    
    const data = await res.json();
    if (data.error) throw new Error(handleGeminiError(data.error.message));
    if (!data.candidates) throw new Error("AI không phản hồi.");
    
    let responseText = data.candidates[0].content.parts[0].text;
    let resultObj;
    try { resultObj = JSON.parse(extractJsonFromText(responseText)); } catch(e) { resultObj = { text: responseText, summary: "Đã xử lý xong." }; }
    let summary = resultObj.summary || "Hoàn thành!";
    
    await Word.run(async (context) => {
      let range = context.document.getSelection();
      if (isEditing && resultObj.edits && resultObj.edits.length > 0 && (!resultObj.text || resultObj.text.trim() === '')) {
        for (const edit of resultObj.edits) {
          if (edit.old && edit.new) {
             let searchResults = range.search(edit.old, { matchCase: true, matchWholeWord: false });
             searchResults.load("items"); await context.sync();
             for (let i = 0; i < searchResults.items.length; i++) searchResults.items[i].insertText(edit.new, Word.InsertLocation.replace);
          }
        }
      } else {
        let textToUse = resultObj.text || responseText; textToUse = cleanLatexForMathType(textToUse);
        if (textToUse && textToUse.trim() !== '') {
            let html = textToUse.trim();
            if (!html.includes('<p>') && !html.includes('<table>') && !html.includes('<ul>')) {
              html = html.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>").replace(/\*(.*?)\*/g, "<em>$1</em>").split(/\n{1,}/).map(p => p.trim()).filter(p => p !== '')
                .map(p => `<p style="margin-top: 6pt; margin-bottom: 6pt; font-family: 'Times New Roman', serif; font-size: 14pt; line-height: 1.15;">${p}</p>`).join('');
            }
            range.insertHtml(html, Word.InsertLocation.replace);
        }
      }
      await context.sync();
    });
    summaryDiv.innerHTML = "<strong>Báo cáo từ AI:</strong><br>" + summary;
    setStatus("Tuyệt vời! AI đã xử lý xong.", "success");
  } catch (e) {
    summaryDiv.innerHTML = `<strong>Lỗi AI:</strong> ${e.message}`;
    summaryDiv.style.borderLeftColor = "var(--error)"; summaryDiv.style.backgroundColor = "rgba(239, 68, 68, 0.1)";
    setStatus(e.message, "error");
  } finally { btn.disabled = false; btn.innerHTML = originalBtnText; btn.style.opacity = "1"; btn.style.cursor = "pointer"; }
}

async function callGeminiDuplicate() {
  const btn = document.getElementById("btnAiDuplicate");
  const extraPrompt = document.getElementById("aiPrompt").value.trim() || "";
  const numExercises = document.getElementById("numExercises").value || "1";
  const apiKey = document.getElementById("geminiApiKey").value.trim();
  if (!apiKey) return setStatus("Vui lòng nhập API Key.", "error");

  localStorage.setItem("geminiApiKey", apiKey);
  setStatus(`Đang nhân bản ${numExercises} bài tập...`, "loading");
  const summaryDiv = document.getElementById("aiSummary");
  summaryDiv.style.display = "block"; summaryDiv.style.borderLeftColor = "#e11d48"; summaryDiv.style.backgroundColor = "rgba(225, 29, 72, 0.1)";
  summaryDiv.innerHTML = `<strong>Đang xử lý:</strong> AI đang nhân bản... ⏳`;
  btn.disabled = true; const originalBtnText = btn.innerHTML; btn.innerHTML = "⏳ Đang tạo..."; btn.style.opacity = "0.7"; btn.style.cursor = "not-allowed";
  
  try {
    let sourceText = "", base64Image = null;
    await Word.run(async (context) => {
      let range = context.document.getSelection(); range.load("text");
      let pics = range.inlinePictures; pics.load("items"); await context.sync();
      sourceText = range.text;
      if (pics.items.length > 0) {
        let pic = pics.items[0]; let base64 = pic.getBase64ImageSrc(); await context.sync();
        base64Image = await resizeImage(base64.value);
      }
    });

    if ((!sourceText || sourceText.trim().length === 0) && !base64Image) throw new Error("Vui lòng bôi đen tài liệu hoặc chọn ảnh.");

    const finalPrompt = `NHIỆM VỤ: Phân tích bài tập gửi kèm và tạo ra ${numExercises} bài tập tương tự nhưng KHÁC SỐ. Yêu cầu phụ: ${extraPrompt}\nQUY TẮC: Ưu tiên nghiệm và số chẵn/tối giản. CẤM Tiếng Việt có dấu vào block LaTeX $...$. Kẻ bảng bằng thẻ <table> HTML.\n${base64Image ? "" : `--- ĐỀ BÀI (VĂN BẢN TRÍCH XUẤT TỪ WORD) ---\n${sourceText}`}\nOUTPUT JSON: {"exercises": [{"problemMarkdown": "...", "solutionMarkdown": "..."}]}`;
    
    let parts = [{ text: finalPrompt }];
    if (base64Image) {
      let mime = base64Image.split(";")[0].split(":")[1] || "image/jpeg";
      parts.push({ inlineData: { mimeType: mime, data: base64Image.split(",")[1] || base64Image } });
    }

    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ contents: [{ parts: parts }], generationConfig: { responseMimeType: "application/json" } })
    });
    
    const data = await res.json();
    if (data.error) throw new Error(handleGeminiError(data.error.message));
    if (!data.candidates) throw new Error("AI không phản hồi.");
    
    let resultObj = JSON.parse(extractJsonFromText(data.candidates[0].content.parts[0].text));

    await Word.run(async (context) => {
      let range = context.document.getSelection(); let htmlOutput = "";
      resultObj.exercises.forEach((ex) => {
        let cleanProblem = cleanLatexForMathType(ex.problemMarkdown || ""), cleanSolution = cleanLatexForMathType(ex.solutionMarkdown || "");
        htmlOutput += `<div style="margin-top: 15px; margin-bottom: 15px; font-family: 'Times New Roman', serif; font-size: 14pt;"><p><strong>Đề bài:</strong><br>${formatTextToHtml(cleanProblem)}</p><p><strong>Lời giải:</strong><br>${formatTextToHtml(cleanSolution)}</p></div>`;
      });
      range.insertHtml(htmlOutput, Word.InsertLocation.after); await context.sync();
    });
    summaryDiv.innerHTML = `<strong>Thành công:</strong> Đã tạo ${resultObj.exercises.length} bài tập!`;
    summaryDiv.style.borderLeftColor = "var(--success)"; summaryDiv.style.backgroundColor = "rgba(46, 204, 113, 0.1)";
    setStatus("Đã tạo đề thi thành công!", "success");
  } catch (error) {
    summaryDiv.innerHTML = `<strong>Lỗi:</strong> ${error.message}`;
    summaryDiv.style.borderLeftColor = "var(--error)"; summaryDiv.style.backgroundColor = "rgba(239, 68, 68, 0.1)";
    setStatus(error.message, "error");
  } finally { btn.disabled = false; btn.innerHTML = originalBtnText; btn.style.opacity = "1"; btn.style.cursor = "pointer"; }
}

async function callGeminiDigitize() {
  const btn = document.getElementById("btnAiDigitize");
  const extraPrompt = document.getElementById("aiPrompt").value.trim() || "";
  const apiKey = document.getElementById("geminiApiKey").value.trim();
  if (!apiKey) return setStatus("Vui lòng nhập API Key.", "error");

  localStorage.setItem("geminiApiKey", apiKey);
  setStatus("AI đang số hoá ảnh...", "loading");
  const summaryDiv = document.getElementById("aiSummary");
  summaryDiv.style.display = "block"; summaryDiv.style.borderLeftColor = "#10b981"; summaryDiv.style.backgroundColor = "rgba(16, 185, 129, 0.1)";
  summaryDiv.innerHTML = "<strong>Đang xử lý:</strong> Đang quét và số hóa ảnh, vui lòng chờ... ⏳";
  btn.disabled = true; const originalBtnText = btn.innerHTML; btn.innerHTML = "⏳ Đang quét..."; btn.style.opacity = "0.7"; btn.style.cursor = "not-allowed";
  
  try {
    let base64Image = null;
    await Word.run(async (context) => {
      let range = context.document.getSelection();
      let pics = range.inlinePictures; pics.load("items"); await context.sync();
      if (pics.items.length > 0) {
        let pic = pics.items[0]; let base64 = pic.getBase64ImageSrc(); await context.sync();
        // GỌI HÀM RESIZE ĐỂ TRÁNH LỖI OVER QUOTA CỦA BẠN LÚC TRƯỚC
        base64Image = await resizeImage(base64.value);
      }
    });

    if (!base64Image) throw new Error("Vui lòng BẤM CHỌN VÀO MỘT BỨC ẢNH trong Word.");

    const finalPrompt = `NHIỆM VỤ: Phân tích tài liệu hình ảnh này và chuyển chính xác sang Text.
${extraPrompt}
QUY TẮC MATHTYPE TRONG WORD BẮT BUỘC: 
1. Không chèn từ ngữ Tiếng Việt có âm tiết vào cặp dấu $.
2. Bảng biểu (nếu có biểu đồ hoặc bảng đề bài) TRÌNH BÀY BẰNG CẤU TRÚC HTML <table>. KHÔNG dùng markdown.
BẠN HÃY TRẢ VỀ TRỰC TIẾP TOÀN BỘ KẾT QUẢ ĐÃ SỐ HÓA SAU ĐÂY (TRẢ VỀ TEXT THỔ, KHÔNG BỌC VÀO BLOCK MARKDOWN \`\`\`html):`;
    
    let mime = base64Image.split(";")[0].split(":")[1] || "image/jpeg";
    let parts = [{ text: finalPrompt }, { inlineData: { mimeType: mime, data: base64Image.split(",")[1] || base64Image } }];

    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ contents: [{ parts: parts }] })
    });
    
    const data = await res.json();
    if (data.error) throw new Error(handleGeminiError(data.error.message));
    if (!data.candidates) throw new Error("AI không thể xem được hình ảnh trong tài liệu này.");
    
    let cleanText = formatTextToHtml(cleanLatexForMathType(data.candidates[0].content.parts[0].text.replace(/```html/gi, '').replace(/```/g, '').trim()));

    await Word.run(async (context) => {
      let range = context.document.getSelection();
      let htmlOutput = `
        <div style="margin-top: 10px; margin-bottom: 10px; border: 1px dashed #10b981; padding: 10px; font-family: 'Times New Roman', serif; font-size: 14pt;">
          <h3 style="color: #10b981; text-align: center; font-size: 16pt;">VĂN BẢN ĐÃ SỐ HOÁ</h3>
          <div>${cleanText}</div>
        </div><br>`;
      range.insertHtml(htmlOutput, Word.InsertLocation.after);
      await context.sync();
    });
    
    summaryDiv.innerHTML = `<strong>Thành công:</strong> Đã quét và số hóa ảnh thành văn bản!`;
    summaryDiv.style.borderLeftColor = "var(--success)"; summaryDiv.style.backgroundColor = "rgba(46, 204, 113, 0.1)";
    setStatus("Đã số hoá ảnh thành công!", "success");
  } catch (error) {
    summaryDiv.innerHTML = `<strong>Lỗi:</strong> ${error.message}`;
    summaryDiv.style.borderLeftColor = "var(--error)"; summaryDiv.style.backgroundColor = "rgba(239, 68, 68, 0.1)";
    setStatus(error.message, "error");
  } finally { btn.disabled = false; btn.innerHTML = originalBtnText; btn.style.opacity = "1"; btn.style.cursor = "pointer"; }
}

async function toolLatexToWord() {
  const btn = document.getElementById("btnLatexToWord");
  const apiKey = document.getElementById("geminiApiKey").value.trim();
  if (!apiKey) return setStatus("Vui lòng nhập API Key.", "error");

  setStatus("Đang dịch MathML...", "loading");
  const summaryDiv = document.getElementById("mathSummary");
  summaryDiv.style.display = "block"; summaryDiv.style.borderLeftColor = "var(--primary)"; summaryDiv.style.backgroundColor = "rgba(99, 102, 241, 0.1)";
  summaryDiv.innerHTML = "<strong>Đang xử lý:</strong> Đang phân tích mã LaTeX sang MathML... ⏳";
  btn.disabled = true; btn.style.opacity = "0.7"; btn.style.cursor = "not-allowed";

  try {
    let sourceText = "";
    await Word.run(async (context) => {
      let range = context.document.getSelection();
      range.load("text"); await context.sync(); sourceText = range.text;
    });

    if (!sourceText || sourceText.trim() === "") throw new Error("Vui lòng bôi đen mã LaTeX ($...$).");

    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ contents: [{ parts: [{ text: `Convert all LaTeX to MathML. Output EXACT original text but replace $...$ with <math xmlns="http://www.w3.org/1998/Math/MathML">...</math>. No markdown blocks. Text:\n${sourceText}` }] }] })
    });
    
    const data = await res.json();
    if (data.error) throw new Error(handleGeminiError(data.error.message));
    let cleanText = data.candidates[0].content.parts[0].text.replace(/\`\`\`html/gi, '').replace(/\`\`\`/g, '').trim();

    await Word.run(async (context) => {
      let range = context.document.getSelection(); range.insertHtml(cleanText, Word.InsertLocation.replace); await context.sync();
    });
    summaryDiv.innerHTML = "<strong>Thành công:</strong> Chuyển hệ thống LaTeX sang Word Equation!";
    summaryDiv.style.borderLeftColor = "var(--success)"; summaryDiv.style.backgroundColor = "rgba(46, 204, 113, 0.1)";
    setStatus("Hoàn tất!", "success");
  } catch (error) {
    summaryDiv.innerHTML = "<strong>Lỗi:</strong> " + error.message; summaryDiv.style.borderLeftColor = "var(--error)"; summaryDiv.style.backgroundColor = "rgba(239, 68, 68, 0.1)";
    setStatus(error.message, "error");
  } finally { btn.disabled = false; btn.style.opacity = "1"; btn.style.cursor = "pointer"; }
}

async function toolWordToLatex() {
  const btn = document.getElementById("btnWordToLatex");
  const apiKey = document.getElementById("geminiApiKey").value.trim();
  if (!apiKey) return setStatus("Vui lòng nhập API Key.", "error");

  setStatus("Đang dịch công thức...", "loading");
  const summaryDiv = document.getElementById("mathSummary");
  summaryDiv.style.display = "block"; summaryDiv.style.borderLeftColor = "#10b981"; summaryDiv.style.backgroundColor = "rgba(16, 185, 129, 0.1)";
  summaryDiv.innerHTML = "<strong>Đang xử lý:</strong> Đọc dữ liệu OOXML để convert sang LaTeX... ⏳";
  btn.disabled = true; btn.style.opacity = "0.7"; btn.style.cursor = "not-allowed";

  try {
    let sourceHtml = "";
    await Word.run(async (context) => {
      let range = context.document.getSelection();
      let html = range.getHtml(); await context.sync(); sourceHtml = html.value;
    });

    if (!sourceHtml) throw new Error("Vui lòng bôi đen tài liệu có chứa Word Equation.");

    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ contents: [{ parts: [{ text: `Extract all MathML/OMML formulas inside this HTML into LaTeX format. Return plain text replacing equations with $...$. Output NO markdown wrappers like \`\`\`. HTML:\n${sourceHtml}` }] }] })
    });
    
    const data = await res.json();
    if (data.error) throw new Error(handleGeminiError(data.error.message));
    let cleanText = data.candidates[0].content.parts[0].text.replace(/\`\`\`html/gi, '').replace(/\`\`\`/g, '').trim();

    await Word.run(async (context) => {
      let range = context.document.getSelection(); range.insertText(cleanText, Word.InsertLocation.replace); await context.sync();
    });

    summaryDiv.innerHTML = "<strong>Thành công:</strong> Chuyển công thức sang dạng LaTeX";
    summaryDiv.style.borderLeftColor = "var(--success)"; summaryDiv.style.backgroundColor = "rgba(46, 204, 113, 0.1)";
    setStatus("Hoàn tất!", "success");
  } catch (error) {
    summaryDiv.innerHTML = "<strong>Lỗi:</strong> " + error.message; summaryDiv.style.borderLeftColor = "var(--error)"; summaryDiv.style.backgroundColor = "rgba(239, 68, 68, 0.1)";
    setStatus(error.message, "error");
  } finally { btn.disabled = false; btn.style.opacity = "1"; btn.style.cursor = "pointer"; }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("btnClean").addEventListener("click", runAutoClean);
    document.getElementById("btnRemoveLinks").addEventListener("click", toolRemoveLinks);
    document.getElementById("btnSwapNumbers").addEventListener("click", toolSwapNumbers);
    document.getElementById("btnUpperCase").addEventListener("click", () => toolChangeCase(true));
    document.getElementById("btnLowerCase").addEventListener("click", () => toolChangeCase(false));
    document.getElementById("btnTestLines").addEventListener("click", toolAddTestLines);
    document.getElementById("btnFormatMCQ").addEventListener("click", toolFormatMCQ);
    document.getElementById("btnFixLatex").addEventListener("click", toolFixLatex);
    document.getElementById("btnAiWrite").addEventListener("click", callGeminiApi);
    document.getElementById("btnAiDuplicate").addEventListener("click", callGeminiDuplicate);
    document.getElementById("btnAiDigitize").addEventListener("click", callGeminiDigitize);
    document.getElementById("btnFindReplace").addEventListener("click", toolFindReplace);
    document.getElementById("btnLatexToWord").addEventListener("click", toolLatexToWord);
    document.getElementById("btnWordToLatex").addEventListener("click", toolWordToLatex);
    document.getElementById("geminiApiKey").value = localStorage.getItem("geminiApiKey") || "";
  }
});
