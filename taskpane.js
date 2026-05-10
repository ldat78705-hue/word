/* global Office, Word */

// Xóa các Map cũ vì đã gỡ tính năng dịch font.
function setStatus(message, type = "") {
  const status = document.getElementById("status");
  status.textContent = message;
  status.className = "status-msg show " + type;
  
  if (status.timeoutId) clearTimeout(status.timeoutId);
  
  if (message && type !== "loading") {
    status.timeoutId = setTimeout(() => {
      status.className = "status-msg " + type;
    }, 4000);
  }
}

function getFriendlyErrorMessage(e) {
  let msg = (e.message || "").toLowerCase();
  let detail = e.message || String(e);
  if (msg.includes("quota") || msg.includes("exceeded") || msg.includes("429")) {
    return "API Key đã quá tải (Quota exceeded). Vui lòng đợi 1 phút và thử lại, hoặc dùng API Key mới!";
  }
  if (msg.includes("api key not valid") || msg.includes("api_key_invalid")) {
    return "API Key không hợp lệ hoặc đã bị khóa. Vui lòng kiểm tra lại!";
  }
  if (msg.includes("400")) {
    return "Yêu cầu không hợp lệ (Lỗi 400). Có thể dữ liệu đầu vào không đúng hoặc kích thước quá lớn.";
  }
  if (msg.includes("network") || msg.includes("fetch") || msg.includes("failed to fetch")) {
    return "Lỗi đường truyền hoặc bị chặn kết nối tới máy chủ Google.";
  }
  return detail;
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
    range.load("text");
    await context.sync();
    if (!range.text || range.text.trim().length === 0) throw new Error("Vui lòng bôi đen nội dung cần thao tác.");
  }
  return range;
}

/* =========================================================
   XỬ LÝ LATEX & MATHTYPE
   ========================================================= */
function cleanLatexForMathType(text) {
  if (!text) return text;
  
  // 1. Chuẩn hóa tag quen thuộc
  text = text.replace(/\\triangle(?![a-zA-Z])/g, "\\Delta");
  text = text.replace(/\\angle\s*([A-Z0-9]+)/g, "\\widehat{$1}");
  text = text.replace(/\\cong/g, "=");
  
  // 2. Tìm và bóc tách \text{...}, \mbox{...}, \mathrm{...} ra khỏi khối $...$ hoặc $$...$$
  text = text.replace(/(\${1,2})([^\$]+)\1/g, function(match, d, formula) {
    // Ép dấu thập phân trong Toán học về dấu phẩy chuẩn Việt Nam (vd: 3.14 -> 3,14)
    formula = formula.replace(/(\d)\.(\d)/g, "$1,$2");
    
    let newF = formula.replace(/\\(?:text|mbox|mathrm|textbf)\s*\{([^}]+)\}/g, ` ${d} $1 ${d} `);
    
    // Tái cấu trúc lại
    let rebuilt = `${d}${newF}${d}`;
    
    // Dọn dẹp khối rỗng sinh ra kiểu $$ $$ hoặc $ $
    let escapedD = d.replace(/\$/g, "\\$");
    let emptyBlockRegex = new RegExp(escapedD + "\\s*" + escapedD, "g");
    rebuilt = rebuilt.replace(emptyBlockRegex, "");
    
    return rebuilt;
  });

  return text;
}

function formatTextToHtml(text) {
  if (!text) return "";
  // Xóa khoảng trắng và \n giữa các thẻ HTML để tránh sinh <br> làm vỡ layout bảng
  let safeText = text.replace(/>\s+</g, "><");
  // Chuyển đổi Markdown Bold sang HTML
  safeText = safeText.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
  // Thay \n bằng <br> cho các đoạn text thông thường
  return safeText.replace(/\n/g, "<br>");
}

async function getGeminiModel(apiKey) {
  let cachedModel = localStorage.getItem("geminiModelName");
  if (cachedModel) return cachedModel;
  
  let modelName = "gemini-1.5-flash"; // Fallback an toàn
  try {
    const modelRes = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
    const modelData = await modelRes.json();
    if (modelData.models) {
      const validModel = modelData.models.find(m => m.name.includes("gemini-1.5-flash")) || 
                         modelData.models.find(m => m.name.includes("gemini-1.5-pro")) ||
                         modelData.models.find(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent"));
      if (validModel) {
        modelName = validModel.name.replace("models/", "");
        localStorage.setItem("geminiModelName", modelName);
      }
    }
  } catch (e) {
    console.warn("Lỗi khi dò tìm model, dùng mặc định.", e);
  }
  return modelName;
}

async function toolFixLatex() {
  setStatus("Đang dọn dẹp lỗi LaTeX...", "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context);
      let oldText = range.text;
      if (!oldText) return;
      
      let newText = cleanLatexForMathType(oldText);
      if (newText !== oldText) {
        range.insertText(newText, Word.InsertLocation.replace);
      }
      await context.sync();
    });
    setStatus("Đã khắc phục lỗi Toán học thành công.", "success");
  } catch(e) { setStatus(e.message, "error"); }
}

/* =========================================================
   CÔNG CỤ ĐỘC LẬP
   ========================================================= */
async function toolRemoveLinks() {
  setStatus("Đang xóa Link...", "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context);
      range.font.underline = "None";
      range.font.color = "#000000";
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
      results.load("items, text");
      await context.sync();
      for (let i = 0; i < results.items.length; i++) {
        let txt = results.items[i].text;
        if (txt.includes(',') || txt.includes('.')) {
          let newTxt = txt;
          let oldTxt;
          do {
            oldTxt = newTxt;
            // Chỉ tráo đổi nếu dấu , hoặc . nằm giữa 2 chữ số (bảo toàn dấu chấm câu ở cuối)
            newTxt = newTxt.replace(/(\d),(\d)/g, "$1TEMP$2");
            newTxt = newTxt.replace(/(\d)\.(\d)/g, "$1COMMA$2");
            newTxt = newTxt.replace(/TEMP/g, ".");
            newTxt = newTxt.replace(/COMMA/g, ",");
          } while (newTxt !== oldTxt);
          
          if (newTxt !== txt) {
            results.items[i].insertText(newTxt, Word.InsertLocation.replace);
          }
        }
      }
      await context.sync();
    });
    setStatus("Đã đổi định dạng số thành công.", "success");
  } catch(e) { setStatus(e.message, "error"); }
}

// Hàm dịch font cũ đã bị gỡ.

async function toolChangeCase(toUpper) {
  setStatus("Đang chuyển đổi hoa/thường...", "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context);
      if (toUpper) {
        range.font.allCaps = true;
      } else {
        range.font.allCaps = false;
        range.load("text");
        await context.sync();
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
      for (let i = 0; i < numLines; i++) {
        rowsHtml += `<tr><td style="border-bottom: 1.5px dotted black; height: 35px;"></td></tr>\n`;
      }
      
      // Sử dụng HTML để tạo bảng thay vì API Word cũ để tránh lỗi undefined
      const html = `
        <table style="width:100%; border-collapse:collapse; margin-top: 10px; margin-bottom: 10px;">
          ${rowsHtml}
        </table>
        <br>
      `;
      range.insertHtml(html, Word.InsertLocation.after);
      
      await context.sync();
    });
    setStatus(`Đã thêm ${numLines} dòng chấm để làm bài tập.`, "success");
  } catch(e) { setStatus(e.message, "error"); }
}

function extractJsonFromText(text) {
  let start = text.indexOf('{');
  let end = text.lastIndexOf('}');
  if (start !== -1 && end !== -1 && end > start) {
    return text.substring(start, end + 1);
  }
  return text;
}

async function callGeminiApi() {
  const btn = document.getElementById("btnAiWrite");
  const prompt = document.getElementById("aiPrompt").value.trim();
  const apiKey = document.getElementById("geminiApiKey").value.trim();
  
  if (!prompt) return setStatus("Vui lòng nhập yêu cầu (Ví dụ: Dịch sang tiếng Anh, hoặc Viết đơn...).", "error");
  if (!apiKey) return setStatus("Vui lòng nhập API Key để sử dụng Gemini.", "error");

  localStorage.setItem("geminiApiKey", apiKey);
  setStatus("AI đang suy nghĩ và xử lý...", "loading");
  
  // Hiển thị trạng thái ngay tại khu vực nút
  const summaryDiv = document.getElementById("aiSummary");
  summaryDiv.style.display = "block";
  summaryDiv.style.borderLeftColor = "var(--primary)";
  summaryDiv.style.backgroundColor = "rgba(99, 102, 241, 0.1)";
  summaryDiv.innerHTML = "<strong>Đang xử lý:</strong> AI đang suy nghĩ, vui lòng chờ trong giây lát... ⏳";

  // Khóa nút trong quá trình xử lý
  btn.disabled = true;
  const originalBtnText = btn.innerHTML;
  btn.innerHTML = "⏳ Đang xử lý...";
  btn.style.opacity = "0.7";
  btn.style.cursor = "not-allowed";
  
  try {
    const mathRules = `
[CÁC QUY TẮC TOÁN HỌC & GIÁO DỤC BẮT BUỘC NẾU CÓ]
1. QUY TẮC SỐ LIỆU ĐẸP (MATH BEAUTY RULES):
- Kết quả & Phương trình: Ưu tiên số nguyên, phân số tối giản (mẫu <100), hạn chế vô tỉ dài.
- Căn thức: Ưu tiên số chính phương hoặc các căn quen thuộc.
- Hình học & Logic: Số liệu phải tạo thành hình hợp logic (Pytago 3-4-5, tổng 2 cạnh...).
- Thống kê: Số liệu chẵn hoặc làm tròn 1-2 chữ số thập phân hợp lý.
2. ĐỊNH DẠNG MATHTYPE MS WORD (VIETNAMESE MATH RULES):
- TUYỆT ĐỐI không chèn Tiếng Việt có dấu vào khối LaTeX ($...$). Các từ như (thỏa mãn) phải để ngoài $$.
- Ký hiệu: Tam giác dùng \\Delta, Góc dùng \\widehat{ABC}, Đồng dạng dùng \\sim, Bằng nhau dùng = (CẤM dùng \\cong, \\angle, \\triangle).
- Phép nhân dùng \\cdot, Số thập phân dùng phẩy (,).
- Đo góc dùng ^\\circ. Song song // hoặc \\parallel, vuông góc \\perp.
- Nếu phát hiện chữ lỗi font TCVN3/VNI, hãy tự động dịch sang Unicode chuẩn.
- Nếu cần kẻ bảng, HÃY DÙNG MÃ HTML (<table>, <tr>, <td> với style viền đen) để Word có thể hiển thị bảng trực tiếp!
`;

    let isEditing = false;
    let finalPrompt = `Yêu cầu của người dùng: ${prompt}\n\n${mathRules}\n\nBạn BẮT BUỘC phải trả về kết quả dưới dạng JSON (không có markdown code block bao quanh) với đúng 2 trường sau:\n- "summary": Một câu ngắn thông báo kết quả. Nếu viết mới thì tóm tắt.\n- "text": ĐOẠN VĂN BẢN ĐÃ XỬ LÝ HOÀN CHỈNH. ĐỂ TRÌNH BÀY ĐẸP VÀ HỖ TRỢ KẺ BẢNG, VUI LÒNG ĐỊNH DẠNG TRƯỜNG TEXT BẰNG HTML (dùng <p>, <strong>, <em>, <table>, <ul>, <li>...).`;
    
    await Word.run(async (context) => {
      let range = context.document.getSelection();
      range.load("text");
      await context.sync();
      if (range.text && range.text.trim().length > 0) {
        isEditing = true;
        finalPrompt = `Bạn là một biên tập viên chuyên nghiệp. Dưới đây là văn bản đang chọn:\n"""\n${range.text}\n"""\n\nYêu cầu: ${prompt}\n\n${mathRules}\n\nĐỂ GIỮ NGUYÊN ĐỊNH DẠNG GỐC CỦA VĂN BẢN, đối với tác vụ sửa lỗi chính tả, thay từ, KHÔNG ĐƯỢC viết lại toàn bộ mà CHỈ LIỆT KÊ các cụm từ ngắn cần thay thế.\nCHỈ KHI yêu cầu là dịch thuật, định dạng bảng, hoặc viết lại hoàn toàn mới dùng trường "text" (phải định dạng bằng mã HTML).\nBẠN BẮT BUỘC TRẢ VỀ JSON SAU:\n{\n  "summary": "Tóm tắt...",\n  "text": "Văn bản mới (dùng HTML). Để rỗng nếu chỉ sửa lỗi.",\n  "edits": [\n    {"old": "từ lỗi gốc", "new": "từ đã sửa"}\n  ]\n}`;
      }
    });

    let modelName = await getGeminiModel(apiKey);

    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: [{ text: finalPrompt }] }],
        generationConfig: { responseMimeType: "application/json" }
      })
    });
    
    const data = await res.json();
    if (data.error) throw new Error(data.error.message || JSON.stringify(data.error));
    if (!data.candidates || data.candidates.length === 0) throw new Error("AI không phản hồi.");
    
    let responseText = data.candidates[0].content?.parts?.[0]?.text;
    if (!responseText) throw new Error("AI trả về kết quả rỗng hoặc bị lỗi.");
    let resultObj;
    try {
      let cleanText = extractJsonFromText(responseText);
      resultObj = JSON.parse(cleanText);
    } catch(e) {
      resultObj = { text: responseText, summary: "Đã xử lý xong (Không thể trích xuất chi tiết lỗi)." };
    }
    
    let summary = resultObj.summary || "Hoàn thành!";
    
    await Word.run(async (context) => {
      let range = context.document.getSelection();
      
      // Nếu có chỉnh sửa nhỏ (edits) và text rỗng -> Cập nhật inline giữ nguyên format
      if (isEditing && resultObj.edits && resultObj.edits.length > 0 && (!resultObj.text || resultObj.text.trim() === '')) {
        for (const edit of resultObj.edits) {
          if (edit.old && edit.new) {
             let searchResults = range.search(edit.old, { matchCase: true, matchWholeWord: false });
             searchResults.load("items");
             await context.sync();
             for (let i = 0; i < searchResults.items.length; i++) {
               searchResults.items[i].insertText(edit.new, Word.InsertLocation.replace);
             }
          }
        }
      } else {
        // Viết lại toàn bộ hoặc dịch thuật -> Replace toàn bộ range
        let textToUse = resultObj.text || "";
        if (!textToUse && responseText) textToUse = responseText; // fallback an toàn
        
        textToUse = cleanLatexForMathType(textToUse);
        
        if (textToUse && textToUse.trim() !== '') {
            let html = textToUse.trim();
            if (!html.includes('<p>') && !html.includes('<table>') && !html.includes('<ul>')) {
              html = html.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
                .replace(/\*(.*?)\*/g, "<em>$1</em>")
                .replace(/^### (.*$)/gim, "<strong>$1</strong>")
                .replace(/^## (.*$)/gim, "<strong style='font-size:15pt;'>$1</strong>")
                .replace(/^# (.*$)/gim, "<strong style='font-size:16pt;'>$1</strong>")
                .split(/\n{1,}/)
                .map(p => p.trim())
                .filter(p => p !== '')
                .map(p => `<p style="margin-top: 6pt; margin-bottom: 6pt; font-family: 'Times New Roman', serif; font-size: 14pt; line-height: 1.15;">${p}</p>`)
                .join('');
            }
            range.insertHtml(html, Word.InsertLocation.replace);
        }
      }
      await context.sync();
    });
    
    document.getElementById("aiSummary").style.display = "block";
    document.getElementById("aiSummary").innerHTML = "<strong>Báo cáo từ AI:</strong><br>" + summary;
    
    setStatus("Tuyệt vời! AI đã xử lý xong.", "success");
  } catch (e) {
    const summaryDiv = document.getElementById("aiSummary");
    let friendlyError = getFriendlyErrorMessage(e);
    summaryDiv.innerHTML = `<strong>Lỗi AI:</strong> ${friendlyError}`;
    summaryDiv.style.display = "block";
    summaryDiv.style.borderLeftColor = "var(--error)";
    summaryDiv.style.backgroundColor = "rgba(239, 68, 68, 0.1)";
    setStatus("Lỗi AI: " + friendlyError, "error");
  } finally {
    // Mở khóa nút lại sau khi xong
    btn.disabled = false;
    btn.innerHTML = originalBtnText;
    btn.style.opacity = "1";
    btn.style.cursor = "pointer";
  }
}

/* =========================================================
   HÀM CHẠY TỰ ĐỘNG CHUẨN HOÁ
   ========================================================= */
async function processAutoClean(context, range, options) {
  // 2. Định dạng Web
  if (options.clearBg) {
    range.font.highlightColor = null;
    let paras = range.paragraphs;
    paras.load("items");
    await context.sync();
    for (let i = 0; i < paras.items.length; i++) {
      paras.items[i].font.highlightColor = null; 
    }
  }
  if (options.standardFont) {
    range.font.name = "Times New Roman";
    range.font.size = 14; 
  }

  // 3. Bảng biểu & Ảnh
  if (options.tableFix) {
    let tables = range.tables;
    tables.load("items");
    await context.sync();
    for (let i = 0; i < tables.items.length; i++) {
      let table = tables.items[i];
      table.autoFitWindow();
      let borders = table.borders;
      borders.load("outsideBorderColor, insideBorderColor");
      await context.sync();
      borders.outsideBorderColor = "#000000";
      borders.outsideBorderStyle = "Single";
      borders.outsideBorderWidth = "pt05"; // pt05 = 0.5pt
      borders.insideBorderColor = "#000000";
      borders.insideBorderStyle = "Single";
      borders.insideBorderWidth = "pt05";
    }
  }
  if (options.imageFix) {
    let pics = range.inlinePictures;
    pics.load("items");
    await context.sync();
    for (let i = 0; i < pics.items.length; i++) {
      pics.items[i].lockAspectRatio = true;
    }
  }

  // 4. Ngắt dòng mềm và Blank lines
  if (options.blankLines) {
    // 4.1. Sửa ngắt dòng mềm (Shift+Enter -> Dấu cách)
    let soft = range.search("^l", { matchWildcards: false });
    soft.load("items");
    await context.sync();
    for (let i = 0; i < soft.items.length; i++) {
      soft.items[i].insertText(" ", Word.InsertLocation.replace);
    }
    await context.sync();
    
    // 4.2. Xoá dòng trống thừa (Dùng vòng lặp paragraphs cho triệt để)
    let paras = range.paragraphs;
    paras.load("items, text");
    await context.sync();
    // Duyệt ngược từ dưới lên để không bị lỗi index khi xoá
    for (let i = paras.items.length - 1; i >= 0; i--) {
      if (paras.items[i].text.trim() === "") {
        paras.items[i].delete();
      }
    }
    await context.sync();
  }

  // 5. Khoảng trắng thừa
  if (options.spaces) {
    let spaces = range.search(" {2,}", { matchWildcards: true });
    spaces.load("items");
    await context.sync();
    for (let i = 0; i < spaces.items.length; i++) spaces.items[i].insertText(" ", Word.InsertLocation.replace);
    await context.sync();

    const puncts = [",", ".", ";", ":", "!", "?"];
    let punctTasks = [];
    for (const p of puncts) {
      let punctSearch = range.search(" " + p, { matchWildcards: false });
      punctSearch.load("items");
      punctTasks.push({ search: punctSearch, replaceText: p });
    }
    await context.sync(); // Đồng bộ một lần cho tất cả dấu câu
    
    for (let t of punctTasks) {
      for (let i = 0; i < t.search.items.length; i++) {
        t.search.items[i].insertText(t.replaceText, Word.InsertLocation.replace);
      }
    }
  }
}

async function runAutoClean() {
  const options = getOptions();
  setStatus("Đang xử lý dọn dẹp toàn diện...", "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context);
      await processAutoClean(context, range, options);
      await context.sync();
    });
    setStatus("Tuyệt vời! Đã chuẩn hoá thành công.", "success");
    setTimeout(() => { if(document.getElementById("status").classList.contains("success")) setStatus("", ""); }, 4000);
  } catch (error) {
    setStatus(error.message, "error");
  }
}

async function toolFormatMCQ() {
  setStatus("Đang dồn các đáp án trắc nghiệm...", "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context, true);
      
      const targets = ["B", "C", "D", "b", "c", "d"];
      const suffixes = [".", ")"];
      const lineBreaks = ["^p", "^l"];
      const spaceVars = ["", " ", "  "];
      let count = 0;
      let searchTasks = [];
      
      // Đưa toàn bộ các lệnh tìm kiếm vào hàng chờ để chạy đồng thời
      for (const lb of lineBreaks) {
        for (const t of targets) {
          for (const s of suffixes) {
            for (const sp of spaceVars) {
              let findStr = `${lb}${sp}${t}${s}`; 
              let replaceStr = `\t\t${t}${s}`;
              let results = range.search(findStr, { matchCase: true, matchWildcards: false });
              results.load("items");
              searchTasks.push({ results: results, replaceStr: replaceStr });
            }
          }
        }
      }
      
      // Gọi một lệnh sync duy nhất thay vì 72 lần (tối ưu tốc độ x30 lần)
      await context.sync();
      
      for (let task of searchTasks) {
        for (let i = 0; i < task.results.items.length; i++) {
          task.results.items[i].insertText(task.replaceStr, Word.InsertLocation.replace);
          count++;
        }
      }
      
      if(count > 0) {
        setStatus(`Đã dồn ngang ${count} khối đáp án thành công.`, "success");
      } else {
        setStatus(`Không tìm thấy cấu trúc đáp án B, C, D nằm rời rạc.`, "error");
      }
    });
  } catch(e) {
    setStatus(e.message, "error");
  }
}

async function toolFindReplace() {
  const findText = document.getElementById("findText").value;
  const replaceText = document.getElementById("replaceText").value;
  
  if (!findText) {
    setStatus("Vui lòng nhập từ cần tìm.", "error");
    return;
  }
  
  setStatus("Đang tìm và thay thế...", "loading");
  try {
    await Word.run(async (context) => {
      let range = await getRange(context, true);
      // Nếu là document (toàn bộ tài liệu), getRange trả về body.getRange()
      let searchResults = range.search(findText, { matchCase: false, matchWholeWord: false });
      searchResults.load("items");
      await context.sync();
      
      let count = searchResults.items.length;
      for (let i = 0; i < count; i++) {
        searchResults.items[i].insertText(replaceText, Word.InsertLocation.replace);
      }
      await context.sync();
      
      const summaryDiv = document.getElementById("findReplaceSummary");
      
      if (count > 0) {
        setStatus(`Đã thay thế ${count} vị trí thành công.`, "success");
        summaryDiv.style.display = "block";
        summaryDiv.style.borderLeftColor = "var(--success)";
        summaryDiv.style.backgroundColor = "rgba(46, 204, 113, 0.1)";
        summaryDiv.innerHTML = `<strong>Thống kê:</strong><br>Đã tìm và thay thế thành công <strong>${count}</strong> từ "<em>${findText}</em>".`;
      } else {
        setStatus("Không tìm thấy từ khoá.", "error");
        summaryDiv.style.display = "block";
        summaryDiv.style.borderLeftColor = "var(--error)";
        summaryDiv.style.backgroundColor = "rgba(239, 68, 68, 0.1)";
        summaryDiv.innerHTML = `<strong>Thống kê:</strong><br>Không tìm thấy từ "<em>${findText}</em>" nào trong phạm vi được chọn.`;
      }
    });
  } catch(e) { 
    setStatus(e.message, "error"); 
  }
}

async function callGeminiDuplicate() {
  const btn = document.getElementById("btnAiDuplicate");
  const extraPrompt = document.getElementById("aiPrompt").value.trim() || "Không có yêu cầu phụ đặc biệt.";
  const numExercises = document.getElementById("numExercises").value || "1";
  const apiKey = document.getElementById("geminiApiKey").value.trim();
  
  if (!apiKey) return setStatus("Vui lòng nhập API Key để sử dụng Gemini.", "error");

  localStorage.setItem("geminiApiKey", apiKey);
  setStatus(`AI đang nhân bản ${numExercises} bài tập...`, "loading");
  
  // Hiển thị trạng thái ngay tại khu vực nút
  const summaryDiv = document.getElementById("aiSummary");
  summaryDiv.style.display = "block";
  summaryDiv.style.borderLeftColor = "#e11d48";
  summaryDiv.style.backgroundColor = "rgba(225, 29, 72, 0.1)";
  summaryDiv.innerHTML = `<strong>Đang xử lý:</strong> AI đang nhân bản ${numExercises} bài tập, vui lòng chờ... ⏳`;

  // Khóa nút trong quá trình xử lý
  btn.disabled = true;
  const originalBtnText = btn.innerHTML;
  btn.innerHTML = "⏳ Đang tạo...";
  btn.style.opacity = "0.7";
  btn.style.cursor = "not-allowed";
  
  try {
    let sourceText = "";
    let base64Image = null;

    await Word.run(async (context) => {
      let range = context.document.getSelection();
      range.load("text");
      
      // Load hình ảnh trong vùng chọn
      let pics = range.inlinePictures;
      pics.load("items");
      
      await context.sync();
      sourceText = range.text;

      if (pics.items.length > 0) {
        let pic = pics.items[0];
        let base64 = pic.getBase64ImageSrc();
        await context.sync();
        base64Image = base64.value;
      }
    });

    if ((!sourceText || sourceText.trim().length === 0) && !base64Image) {
      throw new Error("Vui lòng bôi đen một bài tập gốc (văn bản hoặc hình ảnh) trong Word để AI làm mẫu.");
    }

    const finalPrompt = `ĐÓNG VAI: Chuyên gia ra đề thi cấp Quốc gia.
NHIỆM VỤ: Phân tích bài tập trong ${base64Image ? "bức ảnh" : "đoạn văn bản sau đây"} và tạo ra ${numExercises} bài tập mới TƯƠNG TỰ (về dạng bài, độ khó) nhưng khác số liệu.

YÊU CẦU PHỤ ĐẶC BIỆT TỪ GIÁO VIÊN: ${extraPrompt}

QUAN TRỌNG NHẤT:
1. **Giữ nguyên dạng toán**: Bài tập mới phải dùng cùng phương pháp giải.
2. **THAY SỐ LIỆU THÔNG MINH**: Áp dụng triệt để quy tắc **MATH_BEAUTY_RULES** dưới đây. Không random số bừa bãi.
3. **QUY TẮC CỨNG VỚI HÌNH VẼ**: BỎ QUA HOÀN TOÀN các câu hỏi/bài tập có chứa đồ thị, hình học, hoặc các placeholder liên quan đến hình vẽ. Không tạo bài tập tương tự cho các câu này vì hệ thống không thể chèn ảnh gốc sang bài tập mới. Chỉ nhân bản các bài toán không cần hình vẽ minh hoạ.

**QUY TẮC "SỐ LIỆU ĐẸP & CHÍNH XÁC" (MATH BEAUTY RULES):**
Để đảm bảo tính chuẩn xác và sự logic trong toán học, khoa học:
1. **Kết quả cuối cùng**: Ưu tiên số nguyên hoặc phân số tối giản (mẫu số nhỏ < 100).
2. **Căn thức**: Nếu có căn, số trong căn phải là số chính phương (để khai căn ra số nguyên) hoặc căn thức quen thuộc (căn 2, căn 3).
3. **Hình học và Logic Không Gian**: 
    - Các số liệu đo đạc (độ dài, góc) phải tạo thành một hình có thật (VD: Tổng hai hình tam giác > cạnh thứ ba, tam giác vuông phải tuân đúng Pytago 3-4-5, 5-12-13...).
    - Dữ liệu trong đề thi VÀ sự tương ứng trên hình vẽ (nếu có biểu đồ, hình vẽ AI) phải trùng khớp, logic 100%. Không tự mẫu thuẫn dữ liệu.
4. **Phương trình**: Nghiệm phải đẹp (số nguyên, phân số đơn giản). Không chấp nhận nghiệm vô tỉ dài dòng trừ khi đề bài yêu cầu làm tròn.
5. **Thống kê/Xác suất**: Tỉ lệ phần trăm phải chẵn hoặc làm tròn 1-2 chữ số thập phân hợp lý.

**QUY TẮC VỀ ĐỊNH DẠNG TOÁN HỌC (CỰC KỲ QUAN TRỌNG ĐỂ KHÔNG BỊ LỖI MATHTYPE TRONG WORD):**
Để đảm bảo tính sư phạm và tương thích hoàn toàn với phần mềm MathType khi giáo viên dán vào MS Word, bạn **BẮT BUỘC** phải tuân thủ các quy tắc sau:

1. **Tuyệt đối KHÔNG VIẾT TIẾNG VIỆT CÓ DẤU HOẶC VĂN BẢN TRONG BLOCK LATEX**: 
    - Việc để chữ tiếng Việt có dấu (như "cùng phụ", "đồng", "điều kiện", "thỏa mãn", "loại", "mà", "nên", "hay") hoặc văn bản thường vào trong cặp dấu $ ... $ hoặc $$ ... $$ (kể cả khi dùng \\text{} hay \\mbox{}) sẽ **GÂY LỖI NGHIÊM TRỌNG DẪN ĐẾN HỎNG FONT MATHTYPE**.
    - Bạn PHẢI đóng khối hệ thức Toán lại, viết chữ tiếng Việt ở ngoài, rồi mới mở khối Toán khác.
    - ❌ **Ví dụ SAI 1**: $\\widehat{MBA} (\\text{cùng phụ } \\widehat{ABO}) \\text{ nên } \\widehat{DCB} = \\widehat{MBA}$
    - ✅ **Ví dụ ĐÚNG 1**: $\\widehat{MBA}$ (cùng phụ $\\widehat{ABO}$) nên $\\widehat{DCB} = \\widehat{MBA}$
    - ❌ **Ví dụ SAI 2**: $x = 5 \\text{ (thỏa mãn điều kiện)}$
    - ✅ **Ví dụ ĐÚNG 2**: $x = 5$ (thỏa mãn điều kiện)
    - ❌ **Ví dụ SAI 3**: $\\Delta ABC \\sim \\Delta DEF (c.g.c)$ hoặc $A B C (g-c-g)$
    - ✅ **Ví dụ ĐÚNG 3**: $\\Delta ABC \\sim \\Delta DEF$ (c.g.c)

2. **Ký hiệu Tam giác**: CHỈ dùng \\Delta (tam giác to). Ví dụ: $\\Delta ABC$. **TUYỆT ĐỐI KHÔNG** dùng lệnh \\triangle.
3. **Ký hiệu Góc**: CHỈ dùng \\widehat{ABC} (có mũ ở trên). Ví dụ: $\\widehat{ABC} = 60^\\circ$. **TUYỆT ĐỐI KHÔNG** dùng \\angle ABC.
4. **Cú pháp chuẩn**: Phải viết thường (\\frac, \\sqrt). TUYỆT ĐỐI KHÔNG viết hoa lệnh (\\FRAC).
5. **Tam giác bằng nhau & Đồng dạng**: Dùng = cho bằng nhau ($\\Delta ABC = \\Delta A'B'C'$). Dùng \\sim cho đồng dạng. KHÔNG dùng \\cong.
6. **Độ & Dấu phẩy thập phân**: Đo góc phải có độ (^\\circ). Dấu thập phân của Việt Nam bắt buộc là dấu phẩy , ($3,14$ thay vì $3.14$).
7. **Song song & Vuông góc**: Dùng \\parallel (hoặc //) và \\perp.
8. **Hệ phương trình / Hệ điều kiện**: Khi dùng \\begin{cases} ... \\end{cases}, nếu cần ghi chú (thỏa mãn/loại), CHỈ dùng chữ viết tắt tiếng Việt không dấu: (\\text{TM}) hoặc (\\text{L}) để tránh lỗi font MathType. 

---
**QUY TRÌNH KIỂM DUYỆT CHẤT LƯỢNG 3 LỚP (3-LAYER QA PROTOCOL):**
Trước khi xuất ra kết quả JSON cuối cùng, bạn **PHẢI** thực hiện quy trình tự kiểm tra và sửa lỗi ngầm (Internal Self-Correction) sau đây:

1. **VÒNG 1: KIỂM TRA SỐ LIỆU & LOGIC KHOA HỌC**
    - Tự giải lại bài toán/câu hỏi. Kết quả phải chính xác và "đẹp" (số nguyên, phân số đơn giản).
    - Logic bài dạy phải trơn tru: Hoạt động trước là tiền đề cho hoạt động sau.
    - Thời lượng phân bổ hợp lý, tổng thời gian phải khớp với quy định.

2. **VÒNG 2: BIÊN TẬP VIÊN TOÁN HỌC & NGÔN NGỮ**
    - **LaTeX:** Rà soát từng mã, dùng \\widehat{...} cho góc, \\Delta cho tam giác, dấu phẩy , cho số thập phân.
    - **QUÉT LỖI MATHTYPE (CỰC KỲ QUAN TRỌNG):** BẮT BUỘC rà soát lại toàn bộ công thức. TUYỆT ĐỐI KHÔNG để chữ tiếng Việt có dấu, hoặc các chữ như (c.g.c), chữ viết tắt lọt vào trong khối $ ... $ hoặc $$ ... $$. Hãy kiểm tra xem có bất kỳ lệnh \\text{} chứa tiếng Việt nào không. Nếu có, PHẢI ngắt biểu thức ra ngoài. Ví dụ đúng: $\\widehat{MBA}$ (cùng phụ $\\widehat{ABO}$) nên...
    - **Ngôn ngữ:** Dùng từ ngữ sư phạm chuẩn mực (Ví dụ: "Yêu cầu HS...", "Hướng dẫn HS...", không dùng văn nói).

3. **VÒNG 3: KỸ THUẬT VIÊN MÃ HOÁ BẢNG (QUAN TRỌNG)**
    - **NẾU ĐỀ BÀI CÓ BẢNG BIỂU**, BẠN BẮT BUỘC PHẢI DÙNG THẺ HTML (ví dụ: `<table style="border-collapse: collapse;" border="1">`, `<tr>`, `<td>`) ĐỂ VẼ BẢNG.
    - **TUYỆT ĐỐI KHÔNG** dùng bảng Markdown (`|---|`) vì Microsoft Word không tự động chuyển đổi được.
    - Bạn vẫn có thể dùng Markdown để in đậm (`**`), in nghiêng (`*`), nhưng RIÊNG BẢNG BIỂU thì phải dùng HTML. Dùng thẻ `<br>` nếu cần xuống dòng.

**CHỈ XUẤT RA KẾT QUẢ ĐÃ QUA 3 VÒNG KIỂM TRA VÀ ĐÃ ĐƯỢC SỬA SẠCH LỖI.**
---
${base64Image ? "" : `\n--- ĐỀ BÀI (VĂN BẢN TRÍCH XUẤT TỪ FILE WORD) ---\n${sourceText}`}

OUTPUT JSON:
- exercises: Mảng các bài tập mới. Mỗi bài gồm:
    - problemMarkdown: Đề bài (dùng LaTeX chuẩn cho công thức theo quy tắc ở trên).
    - solutionMarkdown: Lời giải chi tiết (bắt buộc phải giải ra số đẹp).
`;

    let modelName = await getGeminiModel(apiKey);
    
    // Xây dựng parts cho Gemini
    let parts = [{ text: finalPrompt }];
    if (base64Image) {
      let dataImg = base64Image;
      let mime = "image/jpeg";
      if (dataImg.startsWith("data:")) {
        mime = dataImg.split(";")[0].split(":")[1];
        dataImg = dataImg.split(",")[1];
      }
      dataImg = dataImg.replace(/\s+/g, "");
      parts.push({
        inlineData: {
          mimeType: mime,
          data: dataImg
        }
      });
    }

    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: parts }],
        generationConfig: { responseMimeType: "application/json" }
      })
    });
    
    const data = await res.json();
    if (data.error) throw new Error(data.error.message || JSON.stringify(data.error));
    if (!data.candidates || data.candidates.length === 0) throw new Error("AI không phản hồi.");
    
    let responseText = data.candidates[0].content?.parts?.[0]?.text;
    if (!responseText) throw new Error("AI trả về kết quả rỗng hoặc bị lỗi.");
    let resultObj;
    try {
      let cleanText = extractJsonFromText(responseText);
      resultObj = JSON.parse(cleanText);
    } catch(e) {
      throw new Error("AI trả về sai định dạng JSON. Vui lòng thử lại!");
    }
    
    if (!resultObj.exercises || !Array.isArray(resultObj.exercises)) {
      throw new Error("Dữ liệu bài tập trả về bị lỗi cấu trúc.");
    }

    await Word.run(async (context) => {
      let range = context.document.getSelection();
      
      let htmlOutput = "";
      resultObj.exercises.forEach((ex, idx) => {
        let cleanProblem = cleanLatexForMathType(ex.problemMarkdown || "");
        let cleanSolution = cleanLatexForMathType(ex.solutionMarkdown || "");
        let problemHtml = formatTextToHtml(cleanProblem);
        let solutionHtml = formatTextToHtml(cleanSolution);
        htmlOutput += `
        <div style="margin-top: 15px; margin-bottom: 15px; font-family: 'Times New Roman', serif; font-size: 14pt;">
          <p style="margin-bottom: 6pt;"><strong>Đề bài:</strong><br>${problemHtml}</p>
          <p style="margin-top: 6pt;"><strong>Lời giải:</strong><br>${solutionHtml}</p>
        </div>`;
      });

      // Insert at the end of selection
      range.insertHtml(htmlOutput, Word.InsertLocation.after);
      await context.sync();
    });
    
    const summaryDiv = document.getElementById("aiSummary");
    summaryDiv.innerHTML = `<strong>Thành công:</strong> Đã tạo ${resultObj.exercises.length} bài tập!`;
    summaryDiv.style.display = "block";
    summaryDiv.style.borderLeftColor = "var(--success)";
    summaryDiv.style.backgroundColor = "rgba(46, 204, 113, 0.1)";
    setStatus("Đã tạo đề thi thành công!", "success");

  } catch (error) {
    const summaryDiv = document.getElementById("aiSummary");
    let friendlyError = getFriendlyErrorMessage(error);
    summaryDiv.innerHTML = `<strong>Lỗi:</strong> ${friendlyError}`;
    summaryDiv.style.display = "block";
    summaryDiv.style.borderLeftColor = "var(--error)";
    summaryDiv.style.backgroundColor = "rgba(239, 68, 68, 0.1)";
    setStatus("Quá trình nhân bản bị lỗi.", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = originalBtnText;
    btn.style.opacity = "1";
    btn.style.cursor = "pointer";
  }
}

async function callGeminiDigitize() {
  const btn = document.getElementById("btnAiDigitize");
  const extraPrompt = document.getElementById("aiPrompt").value.trim() || "Không có yêu cầu phụ đặc biệt.";
  const apiKey = document.getElementById("geminiApiKey").value.trim();
  
  if (!apiKey) return setStatus("Vui lòng nhập API Key để sử dụng Gemini.", "error");

  localStorage.setItem("geminiApiKey", apiKey);
  setStatus("AI đang số hoá ảnh thành văn bản...", "loading");
  
  // Hiển thị trạng thái ngay tại khu vực nút
  const summaryDiv = document.getElementById("aiSummary");
  summaryDiv.style.display = "block";
  summaryDiv.style.borderLeftColor = "#10b981";
  summaryDiv.style.backgroundColor = "rgba(16, 185, 129, 0.1)";
  summaryDiv.innerHTML = "<strong>Đang xử lý:</strong> Đang quét và số hoá ảnh, vui lòng chờ... ⏳";

  btn.disabled = true;
  const originalBtnText = btn.innerHTML;
  btn.innerHTML = "⏳ Đang quét...";
  btn.style.opacity = "0.7";
  btn.style.cursor = "not-allowed";
  
  try {
    let base64Image = null;

    await Word.run(async (context) => {
      let range = context.document.getSelection();
      let pics = range.inlinePictures;
      pics.load("items");
      await context.sync();

      if (pics.items.length > 0) {
        let pic = pics.items[0];
        let base64 = pic.getBase64ImageSrc();
        await context.sync();
        base64Image = base64.value;
      }
    });

    if (!base64Image) {
      throw new Error("Vui lòng bấm CHỌN VÀO MỘT BỨC ẢNH trong Word để AI có thể nhìn thấy!");
    }

    const finalPrompt = `ĐÓNG VAI: Chuyên gia số hóa tài liệu giáo dục.
NHIỆM VỤ: Phân tích hình ảnh (hoặc PDF file) đề thi và chuyển thể sang định dạng văn bản chính xác tuyệt đối.

YÊU CẦU PHỤ ĐẶC BIỆT TỪ GIÁO VIÊN: ${extraPrompt}

QUY TẮC QUAN TRỌNG:
1. **Giữ nguyên nội dung**: Không được tự ý sửa đề, trừ khi đó là lỗi chính tả rõ ràng.
2. **Đầu trang (Header)**: Nếu đây là trang đầu tiên của đề thi, hãy trích xuất thông tin tiêu đề ở góc trái và góc phải.
3. **Nội dung đề**: Chuyển thành văn bản. Giữ nguyên cấu trúc (Trắc nghiệm/Tự luận). Bắt đầu từ sau phần Header (nếu có).
4. **Hình ảnh**: Thay thế các hình vẽ/đồ thị bằng cách ghi chú [HÌNH_X] (X là số thứ tự 1, 2...).
5. **Bảng biểu & Đáp án**: BẮT BUỘC phải giữ nguyên tối đa cấu trúc bảng. Trình bày bằng cấu trúc BẢNG HTML (Dùng <table>, <tr>, <td>). Không dùng bảng Markdown vì Word không tự chuyển đổi được.

**QUY TẮC VỀ ĐỊNH DẠNG TOÁN HỌC (CỰC KỲ QUAN TRỌNG ĐỂ KHÔNG BỊ LỖI MATHTYPE TRONG WORD):**
Để đảm bảo tính sư phạm và tương thích hoàn toàn với phần mềm MathType khi giáo viên dán vào MS Word, bạn **BẮT BUỘC** phải tuân thủ các quy tắc sau:

1. **Tuyệt đối KHÔNG VIẾT TIẾNG VIỆT CÓ DẤU HOẶC VĂN BẢN TRONG BLOCK LATEX**: 
    - Việc để chữ tiếng Việt có dấu (như "cùng phụ", "đồng", "điều kiện", "thỏa mãn", "loại", "mà", "nên", "hay") hoặc văn bản thường vào trong cặp dấu $ ... $ hoặc $$ ... $$ (kể cả khi dùng \\text{} hay \\mbox{}) sẽ **GÂY LỖI NGHIÊM TRỌNG DẪN ĐẾN HỎNG FONT MATHTYPE**.
    - Bạn PHẢI đóng khối hệ thức Toán lại, viết chữ tiếng Việt ở ngoài, rồi mới mở khối Toán khác.
    - ❌ **Ví dụ SAI 1**: $\\widehat{MBA} (\\text{cùng phụ } \\widehat{ABO}) \\text{ nên } \\widehat{DCB} = \\widehat{MBA}$
    - ✅ **Ví dụ ĐÚNG 1**: $\\widehat{MBA}$ (cùng phụ $\\widehat{ABO}$) nên $\\widehat{DCB} = \\widehat{MBA}$
    - ❌ **Ví dụ SAI 2**: $x = 5 \\text{ (thỏa mãn điều kiện)}$
    - ✅ **Ví dụ ĐÚNG 2**: $x = 5$ (thỏa mãn điều kiện)
    - ❌ **Ví dụ SAI 3**: $\\Delta ABC \\sim \\Delta DEF (c.g.c)$ hoặc $A B C (g-c-g)$
    - ✅ **Ví dụ ĐÚNG 3**: $\\Delta ABC \\sim \\Delta DEF$ (c.g.c)

2. **Ký hiệu Tam giác**: CHỈ dùng \\Delta (tam giác to). Ví dụ: $\\Delta ABC$. **TUYỆT ĐỐI KHÔNG** dùng lệnh \\triangle.
3. **Ký hiệu Góc**: CHỈ dùng \\widehat{ABC} (có mũ ở trên). Ví dụ: $\\widehat{ABC} = 60^\\circ$. **TUYỆT ĐỐI KHÔNG** dùng \\angle ABC.
4. **Cú pháp chuẩn**: Phải viết thường (\\frac, \\sqrt). TUYỆT ĐỐI KHÔNG viết hoa lệnh (\\FRAC).
5. **Tam giác bằng nhau & Đồng dạng**: Dùng = cho bằng nhau ($\\Delta ABC = \\Delta A'B'C'$). Dùng \\sim cho đồng dạng. KHÔNG dùng \\cong.
6. **Độ & Dấu phẩy thập phân**: Đo góc phải có độ (^\\circ). Dấu thập phân của Việt Nam bắt buộc là dấu phẩy , ($3,14$ thay vì $3.14$).
7. **Song song & Vuông góc**: Dùng \\parallel (hoặc //) và \\perp.
8. **Hệ phương trình / Hệ điều kiện**: Khi dùng \\begin{cases} ... \\end{cases}, nếu cần ghi chú (thỏa mãn/loại), CHỈ dùng chữ viết tắt tiếng Việt không dấu: (\\text{TM}) hoặc (\\text{L}) để tránh lỗi font MathType. 

---
**QUY TRÌNH KIỂM DUYỆT CHẤT LƯỢNG 3 LỚP (3-LAYER QA PROTOCOL):**
Trước khi xuất ra kết quả cuối cùng, bạn **PHẢI** thực hiện quy trình tự kiểm tra và sửa lỗi ngầm (Internal Self-Correction) sau đây:

1. **VÒNG 1: KIỂM TRA SỐ LIỆU & LOGIC KHOA HỌC**
    - Đảm bảo không bỏ sót bất kỳ dòng nội dung hay câu hỏi nào từ ảnh truyền vào. Nội dung số hóa phải liền mạch, khớp 100% với ảnh.
2. **VÒNG 2: BIÊN TẬP VIÊN TOÁN HỌC & NGÔN NGỮ**
    - **LaTeX:** Rà soát từng mã, dùng \\widehat{...} cho góc, \\Delta cho tam giác, dấu phẩy , cho số thập phân.
    - **QUÉT LỖI MATHTYPE (CỰC KỲ QUAN TRỌNG):** BẮT BUỘC rà soát lại toàn bộ công thức. TUYỆT ĐỐI KHÔNG để chữ tiếng Việt có dấu, hoặc các chữ như (c.g.c), chữ viết tắt lọt vào trong khối $ ... $ hoặc $$ ... $$.
3. **VÒNG 3: KỸ THUẬT VIÊN MÃ HOÁ (QUAN TRỌNG CHO WORD)**
    - Đảm bảo trả về VĂN BẢN TRỰC TIẾP.
    - Dùng thẻ HTML (<h1>, <p>, <br>, <strong>, <table>) để định dạng. Không dùng Markdown block code (\`\`\`). Word không đọc được bảng Markdown! Mọi bảng biểu phải chuyển thành <table> HTML chuẩn.

**BẠN HÃY TRẢ VỀ TOÀN BỘ KẾT QUẢ SỐ HOÁ BẰNG HTML NGAY SAU ĐÂY:**
`;

    let modelName = await getGeminiModel(apiKey);
    let dataImg = base64Image;
    let mime = "image/jpeg";
    if (dataImg.startsWith("data:")) {
      mime = dataImg.split(";")[0].split(":")[1];
      dataImg = dataImg.split(",")[1];
    }
    dataImg = dataImg.replace(/\s+/g, "");
    
    let parts = [
      { text: finalPrompt },
      {
        inlineData: {
          mimeType: mime,
          data: dataImg
        }
      }
    ];

    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: parts }]
      })
    });
    
    const data = await res.json();
    if (data.error) throw new Error(data.error.message || JSON.stringify(data.error));
    if (!data.candidates || data.candidates.length === 0) throw new Error("AI không thể đọc được ảnh này.");
    
    let responseText = data.candidates[0].content?.parts?.[0]?.text;
    if (!responseText) throw new Error("Kết quả từ AI bị rỗng.");
    
    // Áp dụng bộ lọc LaTeX
    responseText = cleanLatexForMathType(responseText);
    
    // Xóa markdown wrappers nếu AI cố tình dùng
    let cleanText = responseText.replace(/```html/gi, '').replace(/```/g, '').trim();
    cleanText = formatTextToHtml(cleanText);

    await Word.run(async (context) => {
      let range = context.document.getSelection();
      
      let htmlOutput = `
        <div style="margin-top: 10px; margin-bottom: 10px; border: 1px dashed #10b981; padding: 10px; font-family: 'Times New Roman', serif; font-size: 14pt;">
          <h3 style="color: #10b981; text-align: center; font-size: 16pt;">VĂN BẢN ĐÃ SỐ HOÁ</h3>
          <div>${cleanText}</div>
        </div><br>`;

      // Insert at the end of selection
      range.insertHtml(htmlOutput, Word.InsertLocation.after);
      await context.sync();
    });
    
    const summaryDiv = document.getElementById("aiSummary");
    summaryDiv.innerHTML = `<strong>Thành công:</strong> Quá trình quét và số hoá ảnh đã hoàn tất.`;
    summaryDiv.style.display = "block";
    summaryDiv.style.borderLeftColor = "var(--success)";
    summaryDiv.style.backgroundColor = "rgba(46, 204, 113, 0.1)";
    setStatus("Đã số hoá ảnh thành công!", "success");

  } catch (error) {
    const summaryDiv = document.getElementById("aiSummary");
    let friendlyError = getFriendlyErrorMessage(error);
    summaryDiv.innerHTML = `<strong>Lỗi:</strong> ${friendlyError}`;
    summaryDiv.style.display = "block";
    summaryDiv.style.borderLeftColor = "var(--error)";
    summaryDiv.style.backgroundColor = "rgba(239, 68, 68, 0.1)";
    setStatus("Quá trình quét ảnh bị lỗi.", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = originalBtnText;
    btn.style.opacity = "1";
    btn.style.cursor = "pointer";
  }
}

async function toolLatexToWord() {
  const btn = document.getElementById("btnLatexToWord");
  const apiKey = document.getElementById("geminiApiKey").value.trim();
  if (!apiKey) return setStatus("Vui lòng nhập API Key ở mục Trợ lý AI trước.", "error");

  setStatus("Đang chuyển đổi LaTeX sang Word Equation...", "loading");
  const summaryDiv = document.getElementById("mathSummary");
  summaryDiv.style.display = "block";
  summaryDiv.style.borderLeftColor = "var(--primary)";
  summaryDiv.style.backgroundColor = "rgba(99, 102, 241, 0.1)";
  summaryDiv.innerHTML = "<strong>Đang xử lý:</strong> AI đang dịch mã LaTeX sang MathML, vui lòng chờ... ⏳";

  btn.disabled = true; btn.style.opacity = "0.7"; btn.style.cursor = "not-allowed";

  try {
    let sourceText = "";
    await Word.run(async (context) => {
      let range = context.document.getSelection();
      range.load("text");
      await context.sync();
      sourceText = range.text;
    });

    if (!sourceText || sourceText.trim().length === 0) throw new Error("Vui lòng bôi đen văn bản chứa mã LaTeX.");

    const prompt = `Convert all LaTeX math expressions in the following text into MathML format. 
You must output the exact same text, but replace every LaTeX formula (like $...$ or $$...$$) with its corresponding <math xmlns="http://www.w3.org/1998/Math/MathML">...</math> tag.
Return ONLY the final HTML string. Do not include markdown code blocks, do not explain.
Text:
${sourceText}`;

    let modelName = await getGeminiModel(apiKey);
    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
    });
    
    const data = await res.json();
    if (data.error) throw new Error(data.error.message);
    
    let cleanText = data.candidates[0].content.parts[0].text.replace(/\`\`\`html/gi, '').replace(/\`\`\`/g, '').trim();

    await Word.run(async (context) => {
      let range = context.document.getSelection();
      range.insertHtml(cleanText, Word.InsertLocation.replace);
      await context.sync();
    });

    summaryDiv.innerHTML = "<strong>Thành công:</strong> Đã chuyển mã LaTeX sang Word Equation!";
    summaryDiv.style.borderLeftColor = "var(--success)"; summaryDiv.style.backgroundColor = "rgba(46, 204, 113, 0.1)";
    setStatus("Chuyển đổi hoàn tất!", "success");
  } catch (error) {
    let friendlyError = getFriendlyErrorMessage(error);
    summaryDiv.innerHTML = "<strong>Lỗi:</strong> " + friendlyError;
    summaryDiv.style.borderLeftColor = "var(--error)"; summaryDiv.style.backgroundColor = "rgba(239, 68, 68, 0.1)";
    setStatus("Lỗi chuyển đổi.", "error");
  } finally {
    btn.disabled = false; btn.style.opacity = "1"; btn.style.cursor = "pointer";
  }
}

async function toolWordToLatex() {
  const btn = document.getElementById("btnWordToLatex");
  const apiKey = document.getElementById("geminiApiKey").value.trim();
  if (!apiKey) return setStatus("Vui lòng nhập API Key ở mục Trợ lý AI trước.", "error");

  setStatus("Đang dịch công thức Word sang LaTeX...", "loading");
  const summaryDiv = document.getElementById("mathSummary");
  summaryDiv.style.display = "block";
  summaryDiv.style.borderLeftColor = "#10b981";
  summaryDiv.style.backgroundColor = "rgba(16, 185, 129, 0.1)";
  summaryDiv.innerHTML = "<strong>Đang xử lý:</strong> AI đang phân tích OOXML để chuyển sang LaTeX... ⏳";

  btn.disabled = true; btn.style.opacity = "0.7"; btn.style.cursor = "not-allowed";

  try {
    let sourceHtml = "";
    await Word.run(async (context) => {
      let range = context.document.getSelection();
      let html = range.getHtml();
      await context.sync();
      sourceHtml = html.value;
    });

    if (!sourceHtml) throw new Error("Vui lòng bôi đen đoạn văn bản chứa công thức.");

    const prompt = `Extract and convert all MathML/OMML equations from the following HTML into LaTeX format.
You must output the plain text, replacing every math equation with its LaTeX equivalent wrapped in $...$.
Return ONLY the final plain string. Do not use markdown blocks.
HTML:
${sourceHtml}`;

    let modelName = await getGeminiModel(apiKey);
    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
    });
    
    const data = await res.json();
    if (data.error) throw new Error(data.error.message);
    
    let cleanText = data.candidates[0].content.parts[0].text.replace(/\`\`\`html/gi, '').replace(/\`\`\`/g, '').trim();

    await Word.run(async (context) => {
      let range = context.document.getSelection();
      range.insertText(cleanText, Word.InsertLocation.replace);
      await context.sync();
    });

    summaryDiv.innerHTML = "<strong>Thành công:</strong> Đã chuyển công thức thành mã LaTeX!";
    summaryDiv.style.borderLeftColor = "var(--success)"; summaryDiv.style.backgroundColor = "rgba(46, 204, 113, 0.1)";
    setStatus("Chuyển đổi hoàn tất!", "success");
  } catch (error) {
    let friendlyError = getFriendlyErrorMessage(error);
    summaryDiv.innerHTML = "<strong>Lỗi:</strong> " + friendlyError;
    summaryDiv.style.borderLeftColor = "var(--error)"; summaryDiv.style.backgroundColor = "rgba(239, 68, 68, 0.1)";
    setStatus("Lỗi chuyển đổi.", "error");
  } finally {
    btn.disabled = false; btn.style.opacity = "1"; btn.style.cursor = "pointer";
  }
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
