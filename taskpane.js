/* global Office, Word */

// Maps for Encoding
const vniMap = {"aù":"á", "aà":"à", "aû":"ả", "aõ":"ã", "aï":"ạ", "aâ":"â", "aá":"ấ", "aà":"ầ", "aå":"ẩ", "aã":"ẫ", "aä":"ậ", "aê":"ă", "aé":"ắ", "aè":"ằ", "aú":"ẳ", "aø":"ẵ", "aë":"ặ", "eù":"é", "eà":"è", "eû":"ẻ", "eõ":"ẽ", "eï":"ẹ", "eâ":"ê", "eá":"ế", "eà":"ề", "eå":"ể", "eã":"ễ", "eä":"ệ", "où":"ó", "oà":"ò", "oû":"ỏ", "oõ":"õ", "oï":"ọ", "oâ":"ô", "oá":"ố", "oà":"ồ", "oå":"ổ", "oã":"ỗ", "oä":"ộ", "oê":"ơ", "oé":"ớ", "oè":"ờ", "oú":"ở", "oø":"ỡ", "oë":"ợ", "uù":"ú", "uà":"ù", "uû":"ủ", "uõ":"ũ", "uï":"ụ", "uê":"ư", "ué":"ứ", "uè":"ừ", "uú":"ử", "uø":"ữ", "uë":"ự", "iù":"í", "ià":"ì", "iû":"ỉ", "iõ":"ĩ", "iï":"ị", "yù":"ý", "yà":"ỳ", "yû":"ỷ", "yõ":"ỹ", "yï":"ỵ", "dñ":"đ", "DÑ":"Đ"};
const tcvn3Map = {"¸":"á", "µ":"à", "¶":"ả", "·":"ã", "¹":"ạ", "©":"â", "Ê":"ấ", "Ç":"ầ", "È":"ẩ", "É":"ẫ", "Ë":"ậ", "¨":"ă", "¾":"ắ", "»":"ằ", "¼":"ẳ", "½":"ẵ", "Æ":"ặ", "Ð":"é", "Ì":"è", "Î":"ẻ", "Ï":"ẽ", "Ñ":"ẹ", "ª":"ê", "Õ":"ế", "Ò":"ề", "Ó":"ể", "Ô":"ễ", "Ö":"ệ", "ã":"ó", "ß":"ò", "á":"ỏ", "â":"õ", "ä":"ọ", "«":"ô", "è":"ố", "å":"ồ", "æ":"ổ", "ç":"ỗ", "é":"ộ", "¬":"ơ", "í":"ớ", "ê":"ờ", "ë":"ở", "ì":"ỡ", "î":"ợ", "ó":"ú", "ï":"ù", "ñ":"ủ", "ò":"ũ", "ô":"ụ", "­":"ư", "ø":"ứ", "õ":"ừ", "ö":"ử", "÷":"ữ", "ù":"ự", "Ý":"í", "×":"ì", "Ø":"ỉ", "Ü":"ĩ", "Þ":"ị", "ý":"ý", "ú":"ỳ", "û":"ỷ", "ü":"ỹ", "þ":"ỵ", "®":"đ"};

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
        // Chỉ đổi nếu có chứa phẩy/chấm
        if (txt.includes(',') || txt.includes('.')) {
          let newTxt = txt.replace(/,/g, 'TEMP').replace(/\./g, ',').replace(/TEMP/g, '.');
          results.items[i].insertText(newTxt, Word.InsertLocation.replace);
        }
      }
      await context.sync();
    });
    setStatus("Đã đổi định dạng số thành công.", "success");
  } catch(e) { setStatus(e.message, "error"); }
}

async function toolConvertEncoding(type) {
  setStatus(`Đang dịch lỗi phông ${type.toUpperCase()}...`, "loading");
  const map = type === 'vni' ? vniMap : tcvn3Map;
  try {
    await Word.run(async (context) => {
      let range = await getRange(context);
      for (const [bad, good] of Object.entries(map)) {
        let results = range.search(bad, { matchCase: true });
        results.load("items");
        await context.sync();
        for (let i = 0; i < results.items.length; i++) {
          results.items[i].insertText(good, Word.InsertLocation.replace);
        }
      }
      await context.sync();
    });
    setStatus(`Đã dịch ${type.toUpperCase()} thành công.`, "success");
  } catch(e) { setStatus(e.message, "error"); }
}

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

async function callGeminiApi() {
  const btn = document.getElementById("btnAiWrite");
  const prompt = document.getElementById("aiPrompt").value.trim();
  const apiKey = document.getElementById("geminiApiKey").value.trim();
  
  if (!prompt) return setStatus("Vui lòng nhập yêu cầu (Ví dụ: Dịch sang tiếng Anh, hoặc Viết đơn...).", "error");
  if (!apiKey) return setStatus("Vui lòng nhập API Key để sử dụng Gemini.", "error");

  localStorage.setItem("geminiApiKey", apiKey);
  setStatus("AI đang suy nghĩ và xử lý...", "loading");
  
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

    let modelName = "gemini-1.5-flash";
    try {
      const modelRes = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
      const modelData = await modelRes.json();
      if (modelData.models) {
        const validModel = modelData.models.find(m => 
          m.supportedGenerationMethods && 
          m.supportedGenerationMethods.includes("generateContent") &&
          m.name.includes("gemini")
        );
        if (validModel) {
          modelName = validModel.name.replace("models/", "");
        }
      }
    } catch (e) {
      console.warn("Lỗi khi dò tìm model, dùng mặc định.", e);
    }

    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: [{ text: finalPrompt }] }],
        generationConfig: { responseMimeType: "application/json" }
      })
    });
    
    const data = await res.json();
    if (data.error) throw new Error(data.error.message);
    if (!data.candidates || data.candidates.length === 0) throw new Error("AI không phản hồi.");
    
    let responseText = data.candidates[0].content.parts[0].text;
    let resultObj;
    try {
      let cleanText = responseText.replace(/^```json\s*/i, '').replace(/\s*```$/i, '').trim();
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
    setStatus("Lỗi AI: " + e.message, "error");
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
    
    // 4.2. Xoá dòng trống thừa (Thay nhiều Enter bằng 1 Enter)
    // Dùng vòng lặp tìm ^p^p vì tìm kiếm wildcards ^13 đôi khi không ổn định trên mọi bản Word
    let iters = 10; 
    let found = true;
    while(found && iters > 0) {
      let blank = range.search("^p^p", { matchWildcards: false });
      blank.load("items");
      await context.sync();
      if (blank.items.length === 0) { 
        found = false; 
      } else {
        for (let i = 0; i < blank.items.length; i++) {
          blank.items[i].insertText("\r", Word.InsertLocation.replace);
        }
        await context.sync();
      }
      iters--;
    }
  }

  // 5. Khoảng trắng thừa
  if (options.spaces) {
    let spaces = range.search(" {2,}", { matchWildcards: true });
    spaces.load("items");
    await context.sync();
    for (let i = 0; i < spaces.items.length; i++) spaces.items[i].insertText(" ", Word.InsertLocation.replace);

    const puncts = [",", ".", ";", ":", "!", "?"];
    for (const p of puncts) {
      let punctSearch = range.search(" " + p, { matchWildcards: false });
      punctSearch.load("items");
      await context.sync();
      for (let i = 0; i < punctSearch.items.length; i++) punctSearch.items[i].insertText(p, Word.InsertLocation.replace);
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
      let count = 0;
      
      for (const t of targets) {
        for (const s of suffixes) {
            // Thay thế trường hợp không có dấu cách: "^pB." -> "\t\tB."
            let findStr1 = `^p${t}${s}`; 
            let replaceStr1 = `\t\t${t}${s}`;
            let results1 = range.search(findStr1, { matchCase: true, matchWildcards: false });
            results1.load("items");
            await context.sync();
            for (let i = 0; i < results1.items.length; i++) {
              results1.items[i].insertText(replaceStr1, Word.InsertLocation.replace);
              count++;
            }
            
            // Thay thế trường hợp có khoảng trắng thừa: "^p B." -> "\t\tB."
            let findStr2 = `^p ${t}${s}`; 
            let results2 = range.search(findStr2, { matchCase: true, matchWildcards: false });
            results2.load("items");
            await context.sync();
            for (let i = 0; i < results2.items.length; i++) {
              results2.items[i].insertText(replaceStr1, Word.InsertLocation.replace);
              count++;
            }
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

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("btnClean").addEventListener("click", runAutoClean);
    document.getElementById("btnRemoveLinks").addEventListener("click", toolRemoveLinks);
    document.getElementById("btnSwapNumbers").addEventListener("click", toolSwapNumbers);
    document.getElementById("btnConvertVNI").addEventListener("click", () => toolConvertEncoding('vni'));
    document.getElementById("btnConvertTCVN3").addEventListener("click", () => toolConvertEncoding('tcvn3'));
    document.getElementById("btnUpperCase").addEventListener("click", () => toolChangeCase(true));
    document.getElementById("btnLowerCase").addEventListener("click", () => toolChangeCase(false));
    document.getElementById("btnTestLines").addEventListener("click", toolAddTestLines);
    document.getElementById("btnFormatMCQ").addEventListener("click", toolFormatMCQ);
    document.getElementById("btnAiWrite").addEventListener("click", callGeminiApi);
    document.getElementById("btnFindReplace").addEventListener("click", toolFindReplace);
    document.getElementById("geminiApiKey").value = localStorage.getItem("geminiApiKey") || "";
  }
});
