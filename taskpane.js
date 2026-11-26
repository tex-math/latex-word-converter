/* ==========================================================================
 * Chuyển đổi LaTeX Word — Microsoft Word Web Add-in (Vanilla ES6+)
 * Host tĩnh trên GitHub Pages: https://tex-math.github.io/latex-word-converter/
 * Tính năng:
 *  1. 2 chế độ phạm vi: "Vùng đang bôi đen" (tự động chuyển sang Toàn bộ tài
 *     liệu nếu người dùng không bôi đen) và "Toàn bộ tài liệu".
 *  2. Nhận diện chính xác công thức cùng dòng ($...$, \(...\)) và công thức
 *     đứng riêng dòng ($$...$$, \[...\]), tự động bỏ qua ký hiệu tiền tệ.
 *  3. Dùng Temml (CDN) chuyển LaTeX -> MathML -> OMML, đóng gói Flat OPC
 *     OOXML (<pkg:package>) và chèn bằng insertOoxml(..., Word.InsertLocation.replace).
 *  4. Thanh tiến trình & bộ đếm thời gian thực ("Đang xử lý X/Y công thức...")
 *     có nhường luồng UI và nút Dừng để không treo Word khi xử lý đề dài.
 *  5. Xử lý ngoại lệ từng công thức: lỗi cú pháp LaTeX sẽ được bỏ qua, tô màu
 *     đỏ cảnh báo ngay tại chỗ trong Word và tiếp tục chạy các công thức sau.
 * ========================================================================== */

(() => {
  "use strict";

  const HISTORY_STORAGE_KEY = "latex_word_converter_history_v1";
  const MAX_HISTORY_ITEMS = 15;
  const WORD_SEARCH_MAX_LEN = 250;

  const MATH_TEMPLATES = [
    {
      title: "Nghiệm phương trình bậc hai",
      display: true,
      latex: "x_{1,2} = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}"
    },
    {
      title: "Hệ phương trình bậc nhất",
      display: true,
      latex: "\\begin{cases} 2x + y = 5 \\\\ x - y = 1 \\end{cases}"
    },
    {
      title: "Tích phân xác định",
      display: true,
      latex: "I = \\int_{0}^{\\frac{\\pi}{2}} \\sin x \\, dx = 1"
    },
    {
      title: "Giới hạn lượng giác",
      display: true,
      latex: "\\lim_{x \\to 0} \\frac{\\sin x}{x} = 1"
    },
    {
      title: "Tổng cấp số nhân",
      display: true,
      latex: "S_n = \\sum_{k=0}^{n-1} u_1 q^k = u_1 \\frac{1 - q^n}{1 - q}"
    },
    {
      title: "Tổ hợp & Nhị thức Newton",
      display: false,
      latex: "(a+b)^n = \\sum_{k=0}^{n} C_n^k a^{n-k} b^k"
    },
    {
      title: "Tích vô hướng hai vectơ",
      display: false,
      latex: "\\vec{a} \\cdot \\vec{b} = |\\vec{a}| \\cdot |\\vec{b}| \\cdot \\cos(\\vec{a}, \\vec{b})"
    },
    {
      title: "Ma trận vuông 2×2",
      display: true,
      latex: "A = \\begin{pmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{pmatrix}"
    }
  ];

  // Toán tử N-ary lớn chuẩn Word OMML
  const NARY_CHARS = new Set(["∫", "∬", "∭", "∮", "∑", "∏", "⋃", "⋂", "⋁", "⋀", "∐"]);

  // Ánh xạ dấu phụ phía trên cho thẻ <m:acc>
  const ACCENT_MAP = {
    "→": "\u20D7",
    "⃗": "\u20D7",
    "←": "\u20D6",
    "^": "\u0302",
    "̂": "\u0302",
    "~": "\u0303",
    "̃": "\u0303",
    "˙": "\u0307",
    "¨": "\u0308",
    "ˇ": "\u030C",
    "˘": "\u0306"
  };

  // Ánh xạ ký tự Unicode trong OMML sang lệnh LaTeX khi chuyển ngược
  const UNICODE_TO_LATEX = {
    "×": "\\times ",
    "÷": "\\div ",
    "±": "\\pm ",
    "∓": "\\mp ",
    "·": "\\cdot ",
    "⋅": "\\cdot ",
    "≤": "\\le ",
    "≥": "\\ge ",
    "≠": "\\ne ",
    "≈": "\\approx ",
    "≡": "\\equiv ",
    "∞": "\\infty",
    "→": "\\to ",
    "⇒": "\\Rightarrow ",
    "⇔": "\\Leftrightarrow ",
    "∈": "\\in ",
    "∉": "\\notin ",
    "⊂": "\\subset ",
    "⊆": "\\subseteq ",
    "∪": "\\cup ",
    "∩": "\\cap ",
    "∀": "\\forall ",
    "∃": "\\exists ",
    "∂": "\\partial ",
    "∇": "\\nabla ",
    "∠": "\\angle ",
    "⊥": "\\perp ",
    "∥": "\\parallel ",
    "△": "\\triangle ",
    "°": "^\\circ",
    "−": "-",
    "ℝ": "\\mathbb{R}",
    "ℤ": "\\mathbb{Z}",
    "ℕ": "\\mathbb{N}",
    "ℚ": "\\mathbb{Q}",
    "ℂ": "\\mathbb{C}",
    "α": "\\alpha",
    "β": "\\beta",
    "γ": "\\gamma",
    "δ": "\\delta",
    "ε": "\\varepsilon",
    "θ": "\\theta",
    "λ": "\\lambda",
    "μ": "\\mu",
    "π": "\\pi",
    "ρ": "\\rho",
    "σ": "\\sigma",
    "φ": "\\varphi",
    "ω": "\\omega",
    "Δ": "\\Delta",
    "Ω": "\\Omega",
    "Σ": "\\Sigma",
    "Π": "\\Pi",
    "∫": "\\int",
    "∬": "\\iint",
    "∭": "\\iiint",
    "∮": "\\oint",
    "∑": "\\sum",
    "∏": "\\prod"
  };

  const state = {
    isWordHost: false,
    scanScope: "selection", // "selection" (tự động fallback sang "body" nếu rỗng) | "body"
    displayMode: false,
    isBusy: false,
    cancelRequested: false,
    lastMathML: "",
    lastScanItems: []
  };

  const dom = {};

  /* ------------------------------------------------------------------------
   * 1. KHỞI TẠO OFFICE.JS & SỰ KIỆN GIAO DIỆN
   * ------------------------------------------------------------------------ */
  function initApp(isWord) {
    state.isWordHost = Boolean(isWord);
    cacheDom();
    bindEvents();
    renderTemplates();
    renderHistory();
    updateHostIndicator();
    updateScopeHint();
    updateLivePreview();
  }

  const PROJECT_HOME_URL = "https://tex-math.github.io/latex-word-converter/";

  function cacheDom() {
    dom.hostIndicator = document.getElementById("hostIndicator");
    dom.tabBtns = Array.from(document.querySelectorAll(".tab-btn"));
    dom.tabPanels = {
      convert: document.getElementById("panel-convert"),
      about: document.getElementById("panel-about")
    };
    dom.subTabBtns = Array.from(document.querySelectorAll(".sub-tab-btn"));
    dom.subPanels = {
      batch: document.getElementById("subpanel-batch"),
      editor: document.getElementById("subpanel-editor"),
      library: document.getElementById("subpanel-library")
    };
    dom.btnOpenProjectHome = document.getElementById("btnOpenProjectHome");

    // Tab 1: Chuyển đổi hàng loạt
    dom.scopeBtns = Array.from(document.querySelectorAll("[data-scope]"));
    dom.scopeHint = document.getElementById("scopeHint");
    dom.optInlineDollar = document.getElementById("optInlineDollar");
    dom.optDisplayBlock = document.getElementById("optDisplayBlock");
    dom.optSkipCurrency = document.getElementById("optSkipCurrency");
    dom.optHighlightError = document.getElementById("optHighlightError");
    dom.btnBatchConvert = document.getElementById("btnBatchConvert");
    dom.btnScanPreview = document.getElementById("btnScanPreview");
    dom.btnBatchReverse = document.getElementById("btnBatchReverse");
    dom.optMathTypeFontSize = document.getElementById("optMathTypeFontSize");
    dom.btnLatexToMathType = document.getElementById("btnLatexToMathType");
    dom.btnConvertMathType = document.getElementById("btnConvertMathType");
    dom.mtPreviewCanvas = document.getElementById("mtPreviewCanvas");
    dom.btnCancelBatch = document.getElementById("btnCancelBatch");

    dom.batchProgressCard = document.getElementById("batchProgressCard");
    dom.batchProgressText = document.getElementById("batchProgressText");
    dom.batchProgressPct = document.getElementById("batchProgressPct");
    dom.batchProgressFill = document.getElementById("batchProgressFill");
    dom.statSuccessCount = document.getElementById("statSuccessCount");
    dom.statErrorCount = document.getElementById("statErrorCount");

    dom.errorReportCard = document.getElementById("errorReportCard");
    dom.errorCountLabel = document.getElementById("errorCountLabel");
    dom.errorReportList = document.getElementById("errorReportList");
    dom.scanCountLabel = document.getElementById("scanCountLabel");
    dom.scanResultsList = document.getElementById("scanResultsList");

    // Mục con: Soạn công thức
    dom.modeBtns = Array.from(document.querySelectorAll("[data-mode]"));
    dom.symbolToolbar = document.getElementById("symbolToolbar");
    dom.latexInput = document.getElementById("latexInput");
    dom.btnClearInput = document.getElementById("btnClearInput");
    dom.btnCopyMathml = document.getElementById("btnCopyMathml");
    dom.mathPreview = document.getElementById("mathPreview");
    dom.btnInsertWord = document.getElementById("btnInsertWord");
    dom.btnReadFromWord = document.getElementById("btnReadFromWord");
    dom.btnWrapInsertLatex = document.getElementById("btnWrapInsertLatex");

    // Mục con: Mẫu & Lịch sử
    dom.templateList = document.getElementById("templateList");
    dom.historyList = document.getElementById("historyList");
    dom.btnClearHistory = document.getElementById("btnClearHistory");
    dom.statusBar = document.getElementById("statusBar");
  }

  function updateHostIndicator() {
    if (!dom.hostIndicator) return;
    dom.hostIndicator.textContent = state.isWordHost
      ? "Đã kết nối Word"
      : "Chế độ xem thử";
  }

  function updateScopeHint() {
    if (!dom.scopeHint) return;
    dom.scopeHint.textContent =
      state.scanScope === "selection"
        ? "Tự động chuyển sang toàn bộ tài liệu nếu bạn không bôi đen đoạn văn nào."
        : "Sẽ quét và chuyển đổi toàn bộ công thức trong tài liệu hiện tại.";
  }

  function bindEvents() {
    dom.tabBtns.forEach((btn) => {
      btn.addEventListener("click", () => switchTab(btn.dataset.tab));
    });

    dom.subTabBtns.forEach((btn) => {
      btn.addEventListener("click", () => switchSubTab(btn.dataset.subtab));
    });

    if (dom.btnOpenProjectHome) {
      dom.btnOpenProjectHome.addEventListener("click", openProjectHomepage);
    }

    dom.scopeBtns.forEach((btn) => {
      btn.addEventListener("click", () => {
        state.scanScope = btn.dataset.scope === "body" ? "body" : "selection";
        dom.scopeBtns.forEach((b) => b.classList.toggle("active", b === btn));
        updateScopeHint();
      });
    });

    dom.modeBtns.forEach((btn) => {
      btn.addEventListener("click", () => {
        state.displayMode = btn.dataset.mode === "display";
        dom.modeBtns.forEach((b) => b.classList.toggle("active", b === btn));
        updateLivePreview();
      });
    });

    dom.symbolToolbar.addEventListener("click", (event) => {
      const target = event.target.closest(".sym-btn");
      if (!target) return;
      insertSnippetAtCursor(target.dataset.snippet || "");
    });

    dom.latexInput.addEventListener("input", updateLivePreview);
    dom.latexInput.addEventListener("keydown", (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
        event.preventDefault();
        handleInsertEquation();
      }
    });

    dom.btnClearInput.addEventListener("click", () => {
      dom.latexInput.value = "";
      dom.latexInput.focus();
      updateLivePreview();
    });

    dom.btnCopyMathml.addEventListener("click", handleCopyMathML);
    dom.btnInsertWord.addEventListener("click", () => handleInsertEquation());
    dom.btnReadFromWord.addEventListener("click", handleReadEquationFromWord);
    dom.btnWrapInsertLatex.addEventListener("click", handleInsertRawLatexText);

    dom.btnScanPreview.addEventListener("click", handleScanPreview);
    dom.btnBatchConvert.addEventListener("click", handleBatchConvert);
    dom.btnBatchReverse.addEventListener("click", handleBatchReverse);
    if (dom.btnLatexToMathType) {
      dom.btnLatexToMathType.addEventListener("click", () => convertLatexToMathType());
    }
    if (dom.btnConvertMathType) {
      dom.btnConvertMathType.addEventListener("click", () => convertMathTypeToLatex());
    }
    dom.btnCancelBatch.addEventListener("click", () => {
      state.cancelRequested = true;
      dom.btnCancelBatch.disabled = true;
      dom.btnCancelBatch.textContent = "Đang dừng...";
    });

    dom.btnClearHistory.addEventListener("click", () => {
      saveHistory([]);
      renderHistory();
      showStatus("Đã xóa toàn bộ lịch sử công thức.", "success");
    });
  }

  /**
   * Chuyển đổi mượt mà giữa 2 Tab chính ("Chuyển đổi" và "Giới thiệu")
   * bằng cách ẩn/hiện DOM tại chỗ, không tải lại trang hay làm mất dữ liệu đang nhập.
   */
  function switchTab(tabId) {
    // Tương thích ngược khi các hàm nội bộ gọi switchTab("editor" | "batch" | "library")
    if (tabId === "batch" || tabId === "editor" || tabId === "library") {
      switchSubTab(tabId);
      tabId = "convert";
    }
    const targetMain = tabId === "about" ? "about" : "convert";
    dom.tabBtns.forEach((btn) => {
      const isActive = btn.dataset.tab === targetMain;
      btn.classList.toggle("active", isActive);
      btn.setAttribute("aria-selected", String(isActive));
    });
    Object.keys(dom.tabPanels).forEach((key) => {
      if (dom.tabPanels[key]) {
        dom.tabPanels[key].classList.toggle("active", key === targetMain);
      }
    });
  }

  function switchSubTab(subId) {
    const targetSub = dom.subPanels[subId] ? subId : "batch";
    dom.subTabBtns.forEach((btn) => {
      const isActive = btn.dataset.subtab === targetSub;
      btn.classList.toggle("active", isActive);
      btn.setAttribute("aria-selected", String(isActive));
    });
    Object.keys(dom.subPanels).forEach((key) => {
      if (dom.subPanels[key]) {
        dom.subPanels[key].classList.toggle("active", key === targetSub);
      }
    });
  }

  /**
   * Mở trang chủ / kho mã nguồn dự án bằng trình duyệt ngoài trong Microsoft Word
   * (qua Office.context.ui.openBrowserWindow), hoặc dùng thẻ <a target="_blank"> khi chạy web.
   */
  function openProjectHomepage(event) {
    try {
      if (
        typeof window.Office !== "undefined" &&
        window.Office.context &&
        window.Office.context.ui &&
        typeof window.Office.context.ui.openBrowserWindow === "function"
      ) {
        if (event && typeof event.preventDefault === "function") {
          event.preventDefault();
        }
        window.Office.context.ui.openBrowserWindow(PROJECT_HOME_URL);
      }
    } catch (_err) {
      // Dự phòng tự động bằng liên kết target="_blank" của thẻ <a>
    }
  }

  function insertSnippetAtCursor(snippet) {
    const input = dom.latexInput;
    const start = input.selectionStart ?? input.value.length;
    const end = input.selectionEnd ?? input.value.length;
    input.value = input.value.slice(0, start) + snippet + input.value.slice(end);
    const cursorPos = start + snippet.length;
    input.setSelectionRange(cursorPos, cursorPos);
    input.focus();
    updateLivePreview();
  }

  function showStatus(message, type = "info") {
    dom.statusBar.textContent = message;
    dom.statusBar.className = "status-bar visible" + (type !== "info" ? " " + type : "");
  }

  function setBusy(busy, showCancel = false) {
    state.isBusy = busy;
    [
      dom.btnInsertWord,
      dom.btnReadFromWord,
      dom.btnWrapInsertLatex,
      dom.btnScanPreview,
      dom.btnBatchConvert,
      dom.btnBatchReverse,
      dom.btnLatexToMathType,
      dom.optMathTypeFontSize,
      dom.btnConvertMathType
    ].forEach((btn) => {
      if (btn) btn.disabled = busy;
    });

    if (dom.btnCancelBatch) {
      dom.btnCancelBatch.classList.toggle("hidden", !(busy && showCancel));
      dom.btnCancelBatch.disabled = false;
      dom.btnCancelBatch.textContent = "Dừng chuyển đổi";
    }
  }

  function updateProgressUI(processed, total, successCount, errorCount, customLabel) {
    if (!total) {
      dom.batchProgressCard.classList.remove("visible");
      return;
    }
    dom.batchProgressCard.classList.add("visible");
    const pct = Math.min(100, Math.round((processed / total) * 100));
    dom.batchProgressFill.style.width = pct + "%";
    dom.batchProgressPct.textContent = pct + "%";
    dom.batchProgressText.textContent =
      customLabel || ("Đang xử lý " + processed + "/" + total + " công thức...");
    dom.statSuccessCount.textContent = String(successCount);
    dom.statErrorCount.textContent = String(errorCount);
  }

  // Nhường luồng sự kiện để trình duyệt vẽ lại progress bar, tránh treo UI
  function yieldToUi() {
    return new Promise((resolve) => {
      if (typeof window.requestAnimationFrame === "function") {
        window.requestAnimationFrame(() => setTimeout(resolve, 0));
      } else {
        setTimeout(resolve, 0);
      }
    });
  }

  /* ------------------------------------------------------------------------
   * 2. TEMML CỤC BỘ (temml.render / temml.renderToString) & CHUẨN HÓA MATHML
   * ------------------------------------------------------------------------ */
  const MATHML_XMLNS = "http://www.w3.org/1998/Math/MathML";
  const OMML_CTRL_PR =
    '<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></m:ctrlPr>';

  const HTML_ENTITY_TO_NUMERIC = {
    "&nbsp;": "&#160;",
    "&ThinSpace;": "&#8201;",
    "&MediumSpace;": "&#8287;",
    "&ThickSpace;": "&#8197;",
    "&ApplyFunction;": "&#8289;",
    "&InvisibleTimes;": "&#8290;",
    "&InvisibleComma;": "&#8291;",
    "&minus;": "&#8722;",
    "&times;": "&#215;",
    "&middot;": "&#183;",
    "&plusmn;": "&#177;",
    "&le;": "&#8804;",
    "&ge;": "&#8805;",
    "&ne;": "&#8800;",
    "&infin;": "&#8734;",
    "&pi;": "&#960;"
  };

  /**
   * Chuẩn hóa chuỗi MathML trước khi phân tích XML và đóng gói OOXML:
   * - Đảm bảo thẻ <math> luôn có thuộc tính xmlns="http://www.w3.org/1998/Math/MathML".
   * - Chuyển các HTML named entities (nếu có) sang numeric XML entities để DOMParser không lỗi.
   * - Giữ nguyên vẹn cấu trúc các thẻ <msup>, <msub>, <msubsup> để Word dịch sang <m:sSup>
   *   không bị rớt số mũ hoặc chỉ số.
   */
  function normalizeMathML(mathmlString) {
    let xml = String(mathmlString || "").replace(/^\uFEFF/, "").trim();
    if (!xml) {
      throw new Error("Chuỗi MathML trả về từ Temml bị trống.");
    }

    // Thay thế các HTML named entity không thuộc 5 thực thể chuẩn XML
    Object.keys(HTML_ENTITY_TO_NUMERIC).forEach((entity) => {
      if (xml.indexOf(entity) !== -1) {
        xml = xml.split(entity).join(HTML_ENTITY_TO_NUMERIC[entity]);
      }
    });

    // Đảm bảo thẻ <math ...> luôn khai báo đúng xmlns="http://www.w3.org/1998/Math/MathML"
    if (/<math\b/i.test(xml)) {
      xml = xml.replace(/<math\b([^>]*)>/i, (fullMatch, attrs) => {
        if (/\bxmlns\s*=\s*["'][^"']*["']/i.test(attrs)) {
          return "<math" + attrs.replace(/\bxmlns\s*=\s*["'][^"']*["']/i, 'xmlns="' + MATHML_XMLNS + '"') + ">";
        }
        return '<math xmlns="' + MATHML_XMLNS + '"' + attrs + ">";
      });
    } else {
      xml = '<math xmlns="' + MATHML_XMLNS + '">' + xml + "</math>";
    }

    return xml;
  }

  window.normalizeMathML = normalizeMathML;

  function renderLatexToMathML(latex, displayMode) {
    const clean = stripOuterDelimiters(latex);
    if (!clean) {
      throw new Error("Nội dung công thức LaTeX đang trống.");
    }
    const rawMathML = temml.renderToString(clean, {
      displayMode: Boolean(displayMode),
      xml: true,
      throwOnError: true,
      annotate: false
    });
    return normalizeMathML(rawMathML);
  }

  function stripOuterDelimiters(raw) {
    const s = (raw || "").trim();
    if (s.startsWith("$$") && s.endsWith("$$") && s.length > 4) {
      return s.slice(2, -2).trim();
    }
    if (s.startsWith("$") && s.endsWith("$") && s.length > 2) {
      return s.slice(1, -1).trim();
    }
    if (s.startsWith("\\[") && s.endsWith("\\]") && s.length > 4) {
      return s.slice(2, -2).trim();
    }
    if (s.startsWith("\\(") && s.endsWith("\\)") && s.length > 4) {
      return s.slice(2, -2).trim();
    }
    return s;
  }

  function updateLivePreview() {
    const raw = dom.latexInput.value.trim();
    if (!raw) {
      state.lastMathML = "";
      dom.mathPreview.innerHTML =
        '<span class="preview-placeholder">Nhập mã LaTeX để xem trước công thức...</span>';
      return;
    }

    const clean = stripOuterDelimiters(raw);
    try {
      dom.mathPreview.innerHTML = "";
      temml.render(clean, dom.mathPreview, {
        displayMode: Boolean(state.displayMode),
        throwOnError: true
      });
      state.lastMathML = renderLatexToMathML(clean, state.displayMode);
    } catch (err) {
      state.lastMathML = "";
      dom.mathPreview.innerHTML = "";
      const errSpan = document.createElement("span");
      errSpan.className = "preview-error";
      errSpan.textContent =
        "Lỗi cú pháp LaTeX: " + (err && err.message ? err.message : String(err));
      dom.mathPreview.appendChild(errSpan);
    }
  }

  /* ------------------------------------------------------------------------
   * 3. CHUYỂN MATHML -> WORD OMML & ĐÓNG GÓI FLAT OPC OOXML (<pkg:package>)
   * ------------------------------------------------------------------------ */
  function escapeXml(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function createOmmlRun(text, styleMode = false) {
    if (!text) return "";
    let rPr = "";
    if (styleMode === "normal") {
      rPr = "<m:rPr><m:nor/></m:rPr>";
    } else if (styleMode === true || styleMode === "plain") {
      rPr = '<m:rPr><m:sty m:val="p"/></m:rPr>';
    }
    const wRPr =
      '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
    const needsPreserve = /^\s|\s$|\s{2,}/.test(text);
    const spaceAttr = needsPreserve ? ' xml:space="preserve"' : "";
    return "<m:r>" + rPr + wRPr + "<m:t" + spaceAttr + ">" + escapeXml(text) + "</m:t></m:r>";
  }

  function getDirectText(node) {
    return (node && node.textContent ? node.textContent : "").trim();
  }

  function isNaryOperatorNode(node) {
    if (!node || node.nodeType !== 1) return false;
    const tag = node.localName;
    if (tag === "mo") return NARY_CHARS.has(getDirectText(node));
    if (tag === "mrow" && node.children.length === 1) {
      return isNaryOperatorNode(node.children[0]);
    }
    return false;
  }

  function isBoundaryRelation(node) {
    if (!node || node.nodeType !== 1 || node.localName !== "mo") return false;
    const t = getDirectText(node);
    return ["=", "<", ">", "≤", "≥", "≠", "≈", "⇒", "⇔", ",", ";"].includes(t);
  }

  function convertChildrenList(children) {
    let out = "";
    let i = 0;

    while (i < children.length) {
      const node = children[i];
      const tag = node.localName;

      const isNaryWrapper =
        (tag === "msubsup" || tag === "munderover" || tag === "msub" || tag === "munder") &&
        node.children.length >= 2 &&
        isNaryOperatorNode(node.children[0]);

      if (isNaryWrapper) {
        const chr = getDirectText(node.children[0]);
        const isUndOvr = tag === "munderover" || tag === "munder";
        const hasSup = tag === "msubsup" || tag === "munderover";
        const subXml = convertMathNode(node.children[1]);
        const supXml = hasSup && node.children[2] ? convertMathNode(node.children[2]) : "";

        const exprNodes = [];
        let j = i + 1;
        while (j < children.length && !isBoundaryRelation(children[j])) {
          exprNodes.push(children[j]);
          j++;
        }
        const exprXml =
          exprNodes.length > 0 ? convertChildrenList(exprNodes) : createOmmlRun(" ", true);

        const limLoc = isUndOvr ? "undOvr" : "subSup";
        const supHide = hasSup ? "0" : "1";
        out +=
          "<m:nary>" +
            "<m:naryPr>" +
              '<m:chr m:val="' + escapeXml(chr) + '"/>' +
              '<m:limLoc m:val="' + limLoc + '"/>' +
              '<m:subHide m:val="0"/>' +
              '<m:supHide m:val="' + supHide + '"/>' +
            "</m:naryPr>" +
            "<m:sub>" + subXml + "</m:sub>" +
            "<m:sup>" + supXml + "</m:sup>" +
            "<m:e>" + exprXml + "</m:e>" +
          "</m:nary>";

        i = j;
        continue;
      }

      out += convertMathNode(node);
      i++;
    }

    return out;
  }

  function convertMathNode(node) {
    if (!node) return "";
    if (node.nodeType === 3) {
      const txt = node.nodeValue || "";
      return txt.trim() ? createOmmlRun(txt, true) : "";
    }
    if (node.nodeType !== 1) return "";

    const tag = (node.localName || node.nodeName || "").toLowerCase();
    const kids = Array.from(node.children);

    switch (tag) {
      case "math":
      case "mstyle":
      case "mpadded":
      case "mphantom":
      case "menclose":
        return convertChildrenList(kids);

      case "semantics":
        return kids.length > 0 ? convertMathNode(kids[0]) : "";

      case "mrow": {
        if (kids.length >= 2 && kids[0].localName === "mo" && kids[kids.length - 1].localName === "mo") {
          const firstOp = getDirectText(kids[0]);
          const lastOp = getDirectText(kids[kids.length - 1]);
          const openSet = new Set(["(", "[", "{", "|", "‖", "⟨", "⌊", "⌈"]);
          const closeSet = new Set([")", "]", "}", "|", "‖", "⟩", "⌋", "⌉", ""]);
          const isStretchy =
            kids[0].getAttribute("stretchy") === "true" ||
            kids[kids.length - 1].getAttribute("stretchy") === "true" ||
            kids.some((k) => k.localName === "mtable");

          if (isStretchy && openSet.has(firstOp) && closeSet.has(lastOp)) {
            const innerXml = convertChildrenList(kids.slice(1, -1));
            return (
              "<m:d>" +
                "<m:dPr>" +
                  '<m:begChr m:val="' + escapeXml(firstOp) + '"/>' +
                  '<m:endChr m:val="' + escapeXml(lastOp) + '"/>' +
                "</m:dPr>" +
                "<m:e>" + innerXml + "</m:e>" +
              "</m:d>"
            );
          }
        }
        return convertChildrenList(kids);
      }

      case "mi": {
        const text = node.textContent || "";
        const variant = node.getAttribute("mathvariant");
        const isPlain = variant === "normal" || text.length > 1;
        return createOmmlRun(text, isPlain ? "plain" : false);
      }

      case "mn":
        return createOmmlRun(node.textContent || "", "plain");

      case "mtext":
        return createOmmlRun(node.textContent || "", "normal");

      case "mo": {
        const text = node.textContent || "";
        return text ? createOmmlRun(text, "plain") : "";
      }

      case "mspace":
        return createOmmlRun(" ", "plain");

      case "mfrac": {
        if (kids.length < 2) return convertChildrenList(kids);
        const lineThickness = node.getAttribute("linethickness");
        const fPr =
          lineThickness === "0" || lineThickness === "0px"
            ? '<m:fPr><m:type m:val="noBar"/></m:fPr>'
            : "";
        return (
          "<m:f>" +
            fPr +
            "<m:num>" + convertMathNode(kids[0]) + "</m:num>" +
            "<m:den>" + convertMathNode(kids[1]) + "</m:den>" +
          "</m:f>"
        );
      }

      case "msqrt":
        return (
          "<m:rad>" +
            '<m:radPr><m:degHide m:val="1"/></m:radPr>' +
            "<m:deg/>" +
            "<m:e>" + convertChildrenList(kids) + "</m:e>" +
          "</m:rad>"
        );

      case "mroot": {
        if (kids.length < 2) return convertChildrenList(kids);
        return (
          "<m:rad>" +
            "<m:deg>" + convertMathNode(kids[1]) + "</m:deg>" +
            "<m:e>" + convertMathNode(kids[0]) + "</m:e>" +
          "</m:rad>"
        );
      }

      case "msup": {
        if (kids.length < 2) return convertChildrenList(kids);
        const baseXml = convertMathNode(kids[0]) || createOmmlRun(" ", "plain");
        const supXml = convertMathNode(kids[1]);
        return (
          "<m:sSup>" +
            "<m:sSupPr>" + OMML_CTRL_PR + "</m:sSupPr>" +
            "<m:e>" + baseXml + "</m:e>" +
            "<m:sup>" + supXml + "</m:sup>" +
          "</m:sSup>"
        );
      }

      case "msub": {
        if (kids.length < 2) return convertChildrenList(kids);
        if (getDirectText(kids[0]) === "lim") {
          return (
            "<m:limLow>" +
              "<m:e>" + convertMathNode(kids[0]) + "</m:e>" +
              "<m:lim>" + convertMathNode(kids[1]) + "</m:lim>" +
            "</m:limLow>"
          );
        }
        const baseXml = convertMathNode(kids[0]) || createOmmlRun(" ", "plain");
        const subXml = convertMathNode(kids[1]);
        return (
          "<m:sSub>" +
            "<m:sSubPr>" + OMML_CTRL_PR + "</m:sSubPr>" +
            "<m:e>" + baseXml + "</m:e>" +
            "<m:sub>" + subXml + "</m:sub>" +
          "</m:sSub>"
        );
      }

      case "msubsup": {
        if (kids.length < 3) return convertChildrenList(kids);
        const baseXml = convertMathNode(kids[0]) || createOmmlRun(" ", "plain");
        const subXml = convertMathNode(kids[1]);
        const supXml = convertMathNode(kids[2]);
        return (
          "<m:sSubSup>" +
            "<m:sSubSupPr>" + OMML_CTRL_PR + "</m:sSubSupPr>" +
            "<m:e>" + baseXml + "</m:e>" +
            "<m:sub>" + subXml + "</m:sub>" +
            "<m:sup>" + supXml + "</m:sup>" +
          "</m:sSubSup>"
        );
      }

      case "mmultiscripts": {
        if (kids.length === 0) return "";
        const baseXml = convertMathNode(kids[0]) || createOmmlRun(" ", "plain");
        const postSub = kids[1] && kids[1].localName !== "none" && kids[1].localName !== "mprescripts"
          ? convertMathNode(kids[1])
          : "";
        const postSup = kids[2] && kids[2].localName !== "none" && kids[2].localName !== "mprescripts"
          ? convertMathNode(kids[2])
          : "";
        if (postSub && postSup) {
          return (
            "<m:sSubSup>" +
              "<m:sSubSupPr>" + OMML_CTRL_PR + "</m:sSubSupPr>" +
              "<m:e>" + baseXml + "</m:e>" +
              "<m:sub>" + postSub + "</m:sub>" +
              "<m:sup>" + postSup + "</m:sup>" +
            "</m:sSubSup>"
          );
        }
        if (postSup) {
          return (
            "<m:sSup>" +
              "<m:sSupPr>" + OMML_CTRL_PR + "</m:sSupPr>" +
              "<m:e>" + baseXml + "</m:e>" +
              "<m:sup>" + postSup + "</m:sup>" +
            "</m:sSup>"
          );
        }
        if (postSub) {
          return (
            "<m:sSub>" +
              "<m:sSubPr>" + OMML_CTRL_PR + "</m:sSubPr>" +
              "<m:e>" + baseXml + "</m:e>" +
              "<m:sub>" + postSub + "</m:sub>" +
            "</m:sSub>"
          );
        }
        return baseXml;
      }

      case "munder": {
        if (kids.length < 2) return convertChildrenList(kids);
        const underText = getDirectText(kids[1]);
        if (underText === "_" || underText === "̲" || underText === "‾") {
          return (
            "<m:bar>" +
              '<m:barPr><m:pos m:val="bot"/></m:barPr>' +
              "<m:e>" + convertMathNode(kids[0]) + "</m:e>" +
            "</m:bar>"
          );
        }
        if (underText === "⏟") {
          return (
            "<m:groupChr>" +
              '<m:groupChrPr><m:chr m:val="⏟"/><m:pos m:val="bot"/></m:groupChrPr>' +
              "<m:e>" + convertMathNode(kids[0]) + "</m:e>" +
            "</m:groupChr>"
          );
        }
        return (
          "<m:limLow>" +
            "<m:e>" + convertMathNode(kids[0]) + "</m:e>" +
            "<m:lim>" + convertMathNode(kids[1]) + "</m:lim>" +
          "</m:limLow>"
        );
      }

      case "mover": {
        if (kids.length < 2) return convertChildrenList(kids);
        const overText = getDirectText(kids[1]);
        if (overText === "¯" || overText === "‾" || overText === "―") {
          return (
            "<m:bar>" +
              '<m:barPr><m:pos m:val="top"/></m:barPr>' +
              "<m:e>" + convertMathNode(kids[0]) + "</m:e>" +
            "</m:bar>"
          );
        }
        if (overText === "⏞") {
          return (
            "<m:groupChr>" +
              '<m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/></m:groupChrPr>' +
              "<m:e>" + convertMathNode(kids[0]) + "</m:e>" +
            "</m:groupChr>"
          );
        }
        if (ACCENT_MAP[overText]) {
          return (
            "<m:acc>" +
              '<m:accPr><m:chr m:val="' + escapeXml(ACCENT_MAP[overText]) + '"/></m:accPr>' +
              "<m:e>" + convertMathNode(kids[0]) + "</m:e>" +
            "</m:acc>"
          );
        }
        return (
          "<m:limUpp>" +
            "<m:e>" + convertMathNode(kids[0]) + "</m:e>" +
            "<m:lim>" + convertMathNode(kids[1]) + "</m:lim>" +
          "</m:limUpp>"
        );
      }

      case "munderover": {
        if (kids.length < 3) return convertChildrenList(kids);
        const lowXml =
          "<m:limLow>" +
            "<m:e>" + convertMathNode(kids[0]) + "</m:e>" +
            "<m:lim>" + convertMathNode(kids[1]) + "</m:lim>" +
          "</m:limLow>";
        return (
          "<m:limUpp>" +
            "<m:e>" + lowXml + "</m:e>" +
            "<m:lim>" + convertMathNode(kids[2]) + "</m:lim>" +
          "</m:limUpp>"
        );
      }

      case "mtable": {
        const rows = kids.filter((k) => k.localName === "mtr");
        const maxCols = rows.reduce((max, row) => {
          const cells = Array.from(row.children).filter((c) => c.localName === "mtd");
          return Math.max(max, cells.length);
        }, 1);

        const mrXml = rows
          .map((row) => {
            const cells = Array.from(row.children).filter((c) => c.localName === "mtd");
            const cellsXml = cells
              .map((cell) => "<m:e>" + convertChildrenList(Array.from(cell.children)) + "</m:e>")
              .join("");
            return "<m:mr>" + cellsXml + "</m:mr>";
          })
          .join("");

        return (
          "<m:m>" +
            "<m:mPr>" +
              "<m:mcs><m:mc><m:mcPr>" +
                '<m:count m:val="' + maxCols + '"/>' +
                '<m:mcJc m:val="center"/>' +
              "</m:mcPr></m:mc></m:mcs>" +
            "</m:mPr>" +
            mrXml +
          "</m:m>"
        );
      }

      case "mfenced": {
        const open = node.getAttribute("open") ?? "(";
        const close = node.getAttribute("close") ?? ")";
        return (
          "<m:d>" +
            "<m:dPr>" +
              '<m:begChr m:val="' + escapeXml(open) + '"/>' +
              '<m:endChr m:val="' + escapeXml(close) + '"/>' +
            "</m:dPr>" +
            "<m:e>" + convertChildrenList(kids) + "</m:e>" +
          "</m:d>"
        );
      }

      default:
        return convertChildrenList(kids);
    }
  }

  function mathmlToOmml(mathmlString) {
    const normalized = normalizeMathML(mathmlString);
    const parser = new DOMParser();
    const doc = parser.parseFromString(normalized, "application/xml");
    if (doc.querySelector("parsererror")) {
      throw new Error("Không thể phân tích cú pháp MathML trả về từ Temml.");
    }
    return "<m:oMath>" + convertMathNode(doc.documentElement) + "</m:oMath>";
  }

  function buildWordOoxml(oMathXml, displayMode) {
    const mathBody = displayMode
      ? "<m:oMathPara>" + oMathXml + "</m:oMathPara>"
      : oMathXml;

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
        '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">' +
          "<pkg:xmlData>" +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
            "</Relationships>" +
          "</pkg:xmlData>" +
        "</pkg:part>" +
        '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
          "<pkg:xmlData>" +
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">' +
              "<w:body><w:p>" + mathBody + "</w:p></w:body>" +
            "</w:document>" +
          "</pkg:xmlData>" +
        "</pkg:part>" +
      "</pkg:package>"
    );
  }

  /**
   * Kiểm tra ngoặc nhọn cân bằng trước khi chạy bộ dịch dự phòng để vẫn
   * phát hiện chính xác công thức LaTeX sai cú pháp thật sự.
   */
  function hasBalancedLatexBraces(str) {
    let depth = 0;
    for (let i = 0; i < str.length; i++) {
      if (str[i] === "\\" && i + 1 < str.length) {
        i += 1;
        continue;
      }
      if (str[i] === "{") depth += 1;
      else if (str[i] === "}") {
        depth -= 1;
        if (depth < 0) return false;
      }
    }
    return depth === 0;
  }

  /**
   * Bộ dịch dự phòng (Fallback LaTeX-to-OMML) viết bằng JavaScript thuần.
   * Xử lý trơn tru \frac, \sqrt (căn bậc 2 & căn bậc n), chỉ số trên (^),
   * chỉ số dưới (_), ngoặc, tích phân, tổng, giới hạn, vectơ, ma trận và ký hiệu
   * khi gặp cú pháp đặc biệt mà Temml không hỗ trợ.
   */
  function astNodesToOmml(nodes) {
    if (!nodes || nodes.length === 0) return "";
    let out = "";

    function convertSingleNode(node) {
      if (!node) return "";
      switch (node.kind) {
        case "char": {
          const isNormal =
            node.tf === MT_FN_NUMBER ||
            node.tf === MT_FN_TEXT ||
            node.tf === MT_FN_FUNCTION ||
            node.tf === MT_FN_UCGREEK ||
            node.tf === MT_FN_SYMBOL;
          return createOmmlRun(node.ch || "", isNormal);
        }

        case "space":
          return createOmmlRun(" ", true);

        case "text":
        case "func":
          return createOmmlRun(node.name || node.text || "", true);

        case "group":
          return astNodesToOmml(node.items || []);

        case "frac":
          return (
            "<m:f>" +
              "<m:num>" + astNodesToOmml(node.num || []) + "</m:num>" +
              "<m:den>" + astNodesToOmml(node.den || []) + "</m:den>" +
            "</m:f>"
          );

        case "root": {
          const bodyXml = astNodesToOmml(node.body || []);
          if (node.deg && node.deg.length > 0) {
            return (
              "<m:rad>" +
                "<m:deg>" + astNodesToOmml(node.deg) + "</m:deg>" +
                "<m:e>" + bodyXml + "</m:e>" +
              "</m:rad>"
            );
          }
          return (
            "<m:rad>" +
              '<m:radPr><m:degHide m:val="1"/></m:radPr>' +
              "<m:deg/>" +
              "<m:e>" + bodyXml + "</m:e>" +
            "</m:rad>"
          );
        }

        case "script": {
          const subXml = node.sub ? astNodesToOmml(node.sub) : "";
          const supXml = node.sup ? astNodesToOmml(node.sup) : "";
          if (subXml && supXml) {
            return (
              "<m:sSubSup>" +
                "<m:sSubSupPr>" + OMML_CTRL_PR + "</m:sSubSupPr>" +
                "<m:e>" + createOmmlRun(" ", "plain") + "</m:e>" +
                "<m:sub>" + subXml + "</m:sub>" +
                "<m:sup>" + supXml + "</m:sup>" +
              "</m:sSubSup>"
            );
          }
          if (supXml) {
            return (
              "<m:sSup>" +
                "<m:sSupPr>" + OMML_CTRL_PR + "</m:sSupPr>" +
                "<m:e>" + createOmmlRun(" ", "plain") + "</m:e>" +
                "<m:sup>" + supXml + "</m:sup>" +
              "</m:sSup>"
            );
          }
          if (subXml) {
            return (
              "<m:sSub>" +
                "<m:sSubPr>" + OMML_CTRL_PR + "</m:sSubPr>" +
                "<m:e>" + createOmmlRun(" ", "plain") + "</m:e>" +
                "<m:sub>" + subXml + "</m:sub>" +
              "</m:sSub>"
            );
          }
          return "";
        }

        case "fence": {
          const beg = node.left && node.left !== "." ? node.left : "";
          const end = node.right && node.right !== "." ? node.right : "";
          return (
            "<m:d>" +
              "<m:dPr>" +
                '<m:begChr m:val="' + escapeXml(beg) + '"/>' +
                '<m:endChr m:val="' + escapeXml(end) + '"/>' +
              "</m:dPr>" +
              "<m:e>" + astNodesToOmml(node.body || []) + "</m:e>" +
            "</m:d>"
          );
        }

        case "nary": {
          const opChars = {
            int: "∫", iint: "∬", iiint: "∭", oint: "∮",
            sum: "∑", prod: "∏", coprod: "∐", bigcup: "⋃", bigcap: "⋂"
          };
          const chr = opChars[node.op] || "∫";
          const limLoc = node.op && node.op.indexOf("int") !== -1 ? "subSup" : "undOvr";
          const subXml = node.sub ? astNodesToOmml(node.sub) : "";
          const supXml = node.sup ? astNodesToOmml(node.sup) : "";
          return (
            "<m:nary>" +
              "<m:naryPr>" +
                '<m:chr m:val="' + escapeXml(chr) + '"/>' +
                '<m:limLoc m:val="' + limLoc + '"/>' +
                (subXml ? "" : '<m:subHide m:val="1"/>') +
                (supXml ? "" : '<m:supHide m:val="1"/>') +
              "</m:naryPr>" +
              "<m:sub>" + subXml + "</m:sub>" +
              "<m:sup>" + supXml + "</m:sup>" +
              "<m:e/>" +
            "</m:nary>"
          );
        }

        case "lim": {
          const baseXml = createOmmlRun("lim", true);
          const subXml = node.sub ? astNodesToOmml(node.sub) : "";
          return subXml
            ? "<m:limLow><m:e>" + baseXml + "</m:e><m:lim>" + subXml + "</m:lim></m:limLow>"
            : baseXml;
        }

        case "accent": {
          const bodyXml = astNodesToOmml(node.body || []);
          if (node.cmd === "overline" || node.cmd === "bar") {
            return '<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>' + bodyXml + "</m:e></m:bar>";
          }
          if (node.cmd === "underline") {
            return '<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr><m:e>' + bodyXml + "</m:e></m:bar>";
          }
          const accChr =
            node.cmd === "vec" || node.cmd === "overrightarrow"
              ? "\u20D7"
              : node.cmd === "tilde" || node.cmd === "widetilde"
                ? "\u0303"
                : node.cmd === "dot"
                  ? "\u0307"
                  : node.cmd === "ddot"
                    ? "\u0308"
                    : "\u0302";
          return (
            "<m:acc>" +
              '<m:accPr><m:chr m:val="' + escapeXml(accChr) + '"/></m:accPr>' +
              "<m:e>" + bodyXml + "</m:e>" +
            "</m:acc>"
          );
        }

        case "env": {
          const rowsXml = (node.rows || [])
            .map(
              (row) =>
                "<m:mr>" +
                  row.map((cell) => "<m:e>" + astNodesToOmml(cell) + "</m:e>").join("") +
                "</m:mr>"
            )
            .join("");
          const matXml = "<m:m>" + rowsXml + "</m:m>";
          if (node.env === "cases") {
            return '<m:d><m:dPr><m:begChr m:val="{"/><m:endChr m:val=""/></m:dPr><m:e>' + matXml + "</m:e></m:d>";
          }
          if (node.env === "pmatrix") {
            return '<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>' + matXml + "</m:e></m:d>";
          }
          if (node.env === "bmatrix") {
            return '<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e>' + matXml + "</m:e></m:d>";
          }
          return matXml;
        }

        default:
          return "";
      }
    }

    for (let i = 0; i < nodes.length; i++) {
      const node = nodes[i];
      const next = nodes[i + 1];

      // Gắn chỉ số trên (^) và chỉ số dưới (_) liền sau phần tử cơ sở thành <m:sSup>/<m:sSub>/<m:sSubSup>
      if (next && next.kind === "script" && node.kind !== "nary" && node.kind !== "lim") {
        const baseXml = convertSingleNode(node) || createOmmlRun(" ", "plain");
        const subXml = next.sub ? astNodesToOmml(next.sub) : "";
        const supXml = next.sup ? astNodesToOmml(next.sup) : "";
        if (subXml && supXml) {
          out +=
            "<m:sSubSup>" +
              "<m:sSubSupPr>" + OMML_CTRL_PR + "</m:sSubSupPr>" +
              "<m:e>" + baseXml + "</m:e>" +
              "<m:sub>" + subXml + "</m:sub>" +
              "<m:sup>" + supXml + "</m:sup>" +
            "</m:sSubSup>";
        } else if (supXml) {
          out +=
            "<m:sSup>" +
              "<m:sSupPr>" + OMML_CTRL_PR + "</m:sSupPr>" +
              "<m:e>" + baseXml + "</m:e>" +
              "<m:sup>" + supXml + "</m:sup>" +
            "</m:sSup>";
        } else if (subXml) {
          out +=
            "<m:sSub>" +
              "<m:sSubPr>" + OMML_CTRL_PR + "</m:sSubPr>" +
              "<m:e>" + baseXml + "</m:e>" +
              "<m:sub>" + subXml + "</m:sub>" +
            "</m:sSub>";
        } else {
          out += baseXml;
        }
        i += 1;
        continue;
      }

      out += convertSingleNode(node);
    }

    return out;
  }

  function fallbackLatexToOmml(latex) {
    const clean = stripOuterDelimiters(latex);
    if (!clean) {
      throw new Error("Nội dung công thức LaTeX đang trống.");
    }
    const ast = parseLatexToAst(clean);
    const innerXml = astNodesToOmml(ast);
    if (!innerXml) {
      throw new Error("Bộ dịch dự phòng không tạo được nội dung OMML.");
    }
    return "<m:oMath>" + innerXml + "</m:oMath>";
  }

  window.fallbackLatexToOmml = fallbackLatexToOmml;

  function latexToOmmlElement(latex, displayMode) {
    const clean = stripOuterDelimiters(latex);
    if (!clean) {
      throw new Error("Nội dung công thức LaTeX đang trống.");
    }

    try {
      const mathml = renderLatexToMathML(clean, displayMode);
      return mathmlToOmml(mathml);
    } catch (temmlErr) {
      // Nếu công thức thiếu/thừa ngoặc nhọn thì giữ nguyên lỗi cú pháp để báo đỏ tại chỗ
      if (!hasBalancedLatexBraces(clean)) {
        throw temmlErr;
      }
      // Dự phòng bằng bộ dịch LaTeX -> OMML thuần JS khi gặp cú pháp đặc biệt mà Temml báo lỗi
      try {
        const fallbackOmath = fallbackLatexToOmml(clean);
        if (fallbackOmath) {
          return fallbackOmath;
        }
      } catch (_fallbackErr) {
        // Giữ lỗi gốc để hiển thị rõ cho người dùng
      }
      throw temmlErr;
    }
  }

  function latexToWordOoxml(latex, displayMode) {
    const oMath = latexToOmmlElement(latex, displayMode);
    return buildWordOoxml(oMath, displayMode);
  }

  /* ------------------------------------------------------------------------
   * 4. CHUYỂN NGƯỢC WORD OMML -> LATEX
   * ------------------------------------------------------------------------ */
  function mapUnicodeTextToLatex(text) {
    let result = "";
    for (const ch of text) {
      result += UNICODE_TO_LATEX[ch] !== undefined ? UNICODE_TO_LATEX[ch] : ch;
    }
    return result;
  }

  function getChildByLocalName(parent, localName) {
    return Array.from(parent.children).find((el) => el.localName === localName) || null;
  }

  function getChildrenByLocalName(parent, localName) {
    return Array.from(parent.children).filter((el) => el.localName === localName);
  }

  function ommlNodeToLatex(node) {
    if (!node || node.nodeType !== 1) return "";
    const tag = node.localName;

    switch (tag) {
      case "oMathPara":
      case "oMath":
      case "e":
      case "num":
      case "den":
      case "sub":
      case "sup":
      case "deg":
      case "lim":
        return Array.from(node.children).map(ommlNodeToLatex).join("");

      case "r": {
        const raw = getChildrenByLocalName(node, "t")
          .map((t) => t.textContent || "")
          .join("");
        const mapped = mapUnicodeTextToLatex(raw);
        const trimmed = mapped.trim();
        if (["sin", "cos", "tan", "cot", "log", "ln", "lim", "max", "min", "exp"].includes(trimmed)) {
          return "\\" + trimmed + " ";
        }
        return mapped;
      }

      case "f": {
        const fPr = getChildByLocalName(node, "fPr");
        const typeEl = fPr ? getChildByLocalName(fPr, "type") : null;
        const isNoBar =
          typeEl &&
          (typeEl.getAttribute("m:val") === "noBar" || typeEl.getAttribute("val") === "noBar");
        const num = ommlNodeToLatex(getChildByLocalName(node, "num")).trim();
        const den = ommlNodeToLatex(getChildByLocalName(node, "den")).trim();
        return (isNoBar ? "\\binom{" : "\\frac{") + num + "}{" + den + "}";
      }

      case "rad": {
        const deg = ommlNodeToLatex(getChildByLocalName(node, "deg")).trim();
        const e = ommlNodeToLatex(getChildByLocalName(node, "e")).trim();
        return deg ? "\\sqrt[" + deg + "]{" + e + "}" : "\\sqrt{" + e + "}";
      }

      case "sSup": {
        const base = ommlNodeToLatex(getChildByLocalName(node, "e")).trim();
        const sup = ommlNodeToLatex(getChildByLocalName(node, "sup")).trim();
        return base + "^" + formatLatexScriptArg(sup);
      }

      case "sSub": {
        const base = ommlNodeToLatex(getChildByLocalName(node, "e")).trim();
        const sub = ommlNodeToLatex(getChildByLocalName(node, "sub")).trim();
        return base + "_" + formatLatexScriptArg(sub);
      }

      case "sSubSup": {
        const base = ommlNodeToLatex(getChildByLocalName(node, "e")).trim();
        const sub = ommlNodeToLatex(getChildByLocalName(node, "sub")).trim();
        const sup = ommlNodeToLatex(getChildByLocalName(node, "sup")).trim();
        return base + "_" + formatLatexScriptArg(sub) + "^" + formatLatexScriptArg(sup);
      }

      case "nary": {
        const pr = getChildByLocalName(node, "naryPr");
        const chrEl = pr ? getChildByLocalName(pr, "chr") : null;
        const rawChr = (chrEl && (chrEl.getAttribute("m:val") || chrEl.getAttribute("val"))) || "∫";
        const cmd = UNICODE_TO_LATEX[rawChr] || rawChr;
        const sub = ommlNodeToLatex(getChildByLocalName(node, "sub")).trim();
        const sup = ommlNodeToLatex(getChildByLocalName(node, "sup")).trim();
        const e = ommlNodeToLatex(getChildByLocalName(node, "e")).trim();
        return cmd + (sub ? "_{" + sub + "}" : "") + (sup ? "^{" + sup + "}" : "") + (e ? " " + e : "");
      }

      case "limLow":
        return (
          ommlNodeToLatex(getChildByLocalName(node, "e")).trim() +
          "_{" +
          ommlNodeToLatex(getChildByLocalName(node, "lim")).trim() +
          "} "
        );

      case "limUpp":
        return (
          ommlNodeToLatex(getChildByLocalName(node, "e")).trim() +
          "^{" +
          ommlNodeToLatex(getChildByLocalName(node, "lim")).trim() +
          "} "
        );

      case "acc": {
        const pr = getChildByLocalName(node, "accPr");
        const chrEl = pr ? getChildByLocalName(pr, "chr") : null;
        const chr = (chrEl && (chrEl.getAttribute("m:val") || chrEl.getAttribute("val"))) || "\u0302";
        const e = ommlNodeToLatex(getChildByLocalName(node, "e")).trim();
        if (chr === "\u20D7" || chr === "→" || chr === "⃗") return "\\vec{" + e + "}";
        if (chr === "\u0303" || chr === "~") return "\\tilde{" + e + "}";
        if (chr === "\u0307" || chr === "˙") return "\\dot{" + e + "}";
        return "\\hat{" + e + "}";
      }

      case "bar": {
        const pr = getChildByLocalName(node, "barPr");
        const posEl = pr ? getChildByLocalName(pr, "pos") : null;
        const pos = (posEl && (posEl.getAttribute("m:val") || posEl.getAttribute("val"))) || "top";
        const e = ommlNodeToLatex(getChildByLocalName(node, "e")).trim();
        return (pos === "bot" ? "\\underline{" : "\\overline{") + e + "}";
      }

      case "d": {
        const pr = getChildByLocalName(node, "dPr");
        const begEl = pr ? getChildByLocalName(pr, "begChr") : null;
        const endEl = pr ? getChildByLocalName(pr, "endChr") : null;
        const beg = begEl ? (begEl.getAttribute("m:val") ?? begEl.getAttribute("val") ?? "(") : "(";
        const end = endEl ? (endEl.getAttribute("m:val") ?? endEl.getAttribute("val") ?? ")") : ")";
        const eEls = getChildrenByLocalName(node, "e");

        if (beg === "{" && end === "" && eEls.length === 1 && getChildByLocalName(eEls[0], "m")) {
          const matrix = getChildByLocalName(eEls[0], "m");
          const rows = getChildrenByLocalName(matrix, "mr").map((mr) =>
            getChildrenByLocalName(mr, "e").map((c) => ommlNodeToLatex(c).trim()).join(" & ")
          );
          return "\\begin{cases} " + rows.join(" \\\\ ") + " \\end{cases}";
        }

        const content = eEls.map((eEl) => ommlNodeToLatex(eEl).trim()).join(", ");
        const leftDelim = beg ? (beg === "{" ? "\\left\\{" : "\\left" + beg) : "\\left.";
        const rightDelim = end ? (end === "}" ? "\\right\\}" : "\\right" + end) : "\\right.";
        return leftDelim + " " + content + " " + rightDelim;
      }

      case "m": {
        const rows = getChildrenByLocalName(node, "mr").map((mr) =>
          getChildrenByLocalName(mr, "e").map((c) => ommlNodeToLatex(c).trim()).join(" & ")
        );
        return "\\begin{matrix} " + rows.join(" \\\\ ") + " \\end{matrix}";
      }

      default:
        return Array.from(node.children).map(ommlNodeToLatex).join("");
    }
  }

  function extractLatexFromOoxml(ooxmlString) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(ooxmlString, "application/xml");
    const mathNodes = Array.from(doc.getElementsByTagName("*")).filter(
      (el) => el.localName === "oMath"
    );
    if (mathNodes.length === 0) return null;

    const first = mathNodes[0];
    const isDisplay = Boolean(first.parentElement && first.parentElement.localName === "oMathPara");
    const latex = ommlNodeToLatex(first).replace(/\s+/g, " ").trim();
    return { latex, isDisplay, count: mathNodes.length };
  }

  function formatLatexScriptArg(str) {
    const s = (str || "").trim();
    if (!s) return "{}";
    return /^[a-zA-Z0-9]$/.test(s) ? s : "{" + s + "}";
  }

  /* ------------------------------------------------------------------------
   * 5. BỘ TÁCH CÔNG THỨC LATEX BẰNG REGEX CHUẨN JS ($...$, $$...$$, \(...\), \[...\])
   * ------------------------------------------------------------------------ */
  function extractLatexMatches(text, options) {
    const matches = [];
    if (!text) return matches;

    const opts = options || { useInline: true, useDisplay: true, skipCurrency: true };
    const useInline = opts.useInline !== false;
    const useDisplay = opts.useDisplay !== false;
    const skipCurrency = Boolean(opts.skipCurrency);

    function isEscapedAt(str, idx) {
      let backslashes = 0;
      for (let k = idx - 1; k >= 0 && str[k] === "\\"; k--) {
        backslashes += 1;
      }
      return backslashes % 2 === 1;
    }

    // 1. Dùng biểu thức chính quy chuẩn JS trích xuất chính xác $$...$$ và $...$ (bao gồm $2^x$, $a^{x+1}$, $x^{2n+1}$)
    const dollarRegex = /\$\$([\s\S]+?)\$\$|\$([^\$]+?)\$/g;
    let m;
    while ((m = dollarRegex.exec(text)) !== null) {
      const start = m.index;
      const raw = m[0];
      const end = start + raw.length;

      // Bỏ qua ký tự thoát \$
      if (isEscapedAt(text, start)) {
        dollarRegex.lastIndex = start + 1;
        continue;
      }

      if (m[1] !== undefined) {
        if (!useDisplay) continue;
        const inner = m[1].trim();
        if (!inner) continue;
        matches.push({ raw, latex: inner, displayMode: true, start, end });
      } else {
        if (!useInline) continue;
        if (isEscapedAt(text, end - 1)) continue;
        const inner = (m[2] || "").trim();
        if (!inner) continue;
        const looksLikeCurrency = skipCurrency && /^\d[\d.,]*$/.test(inner);
        if (looksLikeCurrency) continue;
        matches.push({ raw, latex: inner, displayMode: false, start, end });
      }
    }

    // 2. Hỗ trợ thêm \[...\] và \(...\) không chồng lấn với các chuỗi $...$ đã tìm
    const bracketRegex = /\\\[([\s\S]+?)\\\]|\\\(([\s\S]+?)\\\)/g;
    while ((m = bracketRegex.exec(text)) !== null) {
      const start = m.index;
      const raw = m[0];
      const end = start + raw.length;

      const overlaps = matches.some((ex) => start < ex.end && end > ex.start);
      if (overlaps) continue;

      if (m[1] !== undefined) {
        if (!useDisplay) continue;
        const inner = m[1].trim();
        if (!inner) continue;
        matches.push({ raw, latex: inner, displayMode: true, start, end });
      } else {
        if (!useInline) continue;
        const inner = (m[2] || "").trim();
        if (!inner) continue;
        matches.push({ raw, latex: inner, displayMode: false, start, end });
      }
    }

    matches.sort((a, b) => a.start - b.start);
    return matches;
  }

  function getScanOptions() {
    return {
      scope: state.scanScope,
      useInline: dom.optInlineDollar.checked,
      useDisplay: dom.optDisplayBlock.checked,
      skipCurrency: dom.optSkipCurrency.checked,
      highlightError: dom.optHighlightError.checked
    };
  }

  /* ------------------------------------------------------------------------
   * 6. TƯƠNG TÁC MICROSOFT WORD (OFFICE.JS — Word.run)
   * ------------------------------------------------------------------------ */
  function ensureWordAvailable() {
    if (typeof window.Word === "undefined" || typeof window.Word.run !== "function") {
      showStatus(
        "Không tìm thấy môi trường Word (Word.run). Vui lòng mở Add-in trong Microsoft Word.",
        "warn"
      );
      return false;
    }
    return true;
  }

  /**
   * Tự động xác định vùng mục tiêu (Vùng đang bôi đen hoặc Toàn bộ tài liệu).
   * Nếu người dùng chọn "Vùng đang bôi đen" nhưng chưa bôi đen văn bản nào,
   * hệ thống tự động chuyển sang quét "Toàn bộ tài liệu".
   */
  async function resolveTargetContainer(context, requestedScope) {
    if (requestedScope === "body") {
      return {
        container: context.document.body,
        effectiveScope: "body",
        autoSwitchedToBody: false
      };
    }

    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const hasSelectionText = (selection.text || "").trim().length > 0;
    if (!hasSelectionText) {
      return {
        container: context.document.body,
        effectiveScope: "body",
        autoSwitchedToBody: true
      };
    }

    return {
      container: selection,
      effectiveScope: "selection",
      autoSwitchedToBody: false
    };
  }

  /**
   * Tìm chính xác một chuỗi văn bản trong đoạn Word bằng paragraph.search(..., { matchWildcards: false }).
   * Đồng thời xử lý riêng đặc thù của Microsoft Word Desktop: khi chuỗi chứa dấu lũy thừa '^'
   * (ví dụ $2^x$, $x^2$, $a^{x+1}$), Word Desktop có thể hiểu '^' là mã ký tự đặc biệt nên
   * tự động thử lại bằng cách thoát '^' -> '^^' (vẫn tắt matchWildcards: false).
   */
  async function searchExactInParagraph(context, para, exactString) {
    if (!exactString || exactString.length > WORD_SEARCH_MAX_LEN) {
      return { items: [] };
    }

    // Lần 1: tìm chính xác nguyên văn với matchWildcards: false
    try {
      const results = para.search(exactString, {
        matchCase: true,
        matchWildcards: false
      });
      results.load("items");
      await context.sync();
      if (results.items && results.items.length > 0) {
        return results;
      }
    } catch (_firstErr) {
      // Một số phiên bản Word Desktop báo lỗi khi chuỗi chứa ^ (như ^x, ^2)
    }

    // Lần 2: nếu chuỗi chứa dấu lũy thừa '^', thoát '^' thành '^^' cho Word Desktop Find
    if (exactString.indexOf("^") !== -1) {
      const escapedCaret = exactString.replace(/\^/g, "^^");
      if (escapedCaret.length <= WORD_SEARCH_MAX_LEN) {
        try {
          const retryResults = para.search(escapedCaret, {
            matchCase: true,
            matchWildcards: false
          });
          retryResults.load("items");
          await context.sync();
          if (retryResults.items && retryResults.items.length > 0) {
            return retryResults;
          }
        } catch (_retryErr) {
          // Chuyển sang cơ chế thay thế trực tiếp qua OOXML của đoạn văn
        }
      }
    }

    return { items: [] };
  }

  /**
   * Dự phòng cấp 3: Thay thế trực tiếp chuỗi công thức trong OOXML của đoạn văn <w:p>
   * bằng nút <m:oMath> (hoặc <m:oMathPara>) khi paragraph.search không định vị được
   * (ví dụ công thức dài > 250 ký tự hoặc bị tách run đặc biệt trong Word).
   */
  async function replaceRawFormulaViaParaOoxml(context, para, rawText, oMathXml, displayMode) {
    if (typeof para.getOoxml !== "function" || typeof para.insertOoxml !== "function") {
      return 0;
    }

    const ooxmlRes = para.getOoxml();
    await context.sync();

    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(ooxmlRes.value || "", "application/xml");
    if (xmlDoc.querySelector("parsererror")) {
      return 0;
    }

    const mNs = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    const wDocEl = Array.from(xmlDoc.getElementsByTagName("*")).find(
      (el) => el.localName === "document"
    );
    if (wDocEl && !wDocEl.getAttribute("xmlns:m")) {
      wDocEl.setAttribute("xmlns:m", mNs);
    }

    // Tạo DOM node <m:oMath> hoặc <m:oMathPara> hợp lệ có khai báo đủ namespace
    const mathWrapperXml =
      '<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="' +
      mNs +
      '">' +
      (displayMode ? "<m:oMathPara>" + oMathXml + "</m:oMathPara>" : oMathXml) +
      "</root>";
    const mathDoc = parser.parseFromString(mathWrapperXml, "application/xml");
    if (mathDoc.querySelector("parsererror") || !mathDoc.documentElement.firstElementChild) {
      return 0;
    }
    const mathTemplateNode = mathDoc.documentElement.firstElementChild;

    const pNodes = Array.from(xmlDoc.getElementsByTagName("*")).filter(
      (el) => el.localName === "p"
    );
    let replacedTotal = 0;

    pNodes.forEach((pEl) => {
      const { spans, fullText } = collectParagraphTextSpans(pEl);
      if (!fullText || fullText.indexOf(rawText) === -1) return;

      const occurrences = [];
      let searchPos = 0;
      while (searchPos <= fullText.length - rawText.length) {
        const idx = fullText.indexOf(rawText, searchPos);
        if (idx === -1) break;
        occurrences.push({ start: idx, end: idx + rawText.length });
        searchPos = idx + rawText.length;
      }

      // Thay thế từ phải sang trái để giữ nguyên chỉ mục và định dạng <w:rPr> xung quanh
      for (let i = occurrences.length - 1; i >= 0; i--) {
        const { start, end } = occurrences[i];
        const affected = spans.filter((sp) => sp.end > start && sp.start < end);
        if (affected.length === 0) continue;

        const firstSpan = affected[0];
        const lastSpan = affected[affected.length - 1];
        const anchorParent = firstSpan.runEl.parentNode;
        if (!anchorParent) continue;

        const beforeText = firstSpan.text.slice(0, Math.max(0, start - firstSpan.start));
        const afterText = lastSpan.text.slice(Math.max(0, end - lastSpan.start));

        const frag = xmlDoc.createDocumentFragment();
        if (beforeText) {
          frag.appendChild(createPreservedTextRun(xmlDoc, beforeText, firstSpan.rPr, false));
        }
        frag.appendChild(xmlDoc.importNode(mathTemplateNode, true));
        if (afterText) {
          frag.appendChild(createPreservedTextRun(xmlDoc, afterText, lastSpan.rPr, false));
        }

        anchorParent.insertBefore(frag, firstSpan.runEl);
        affected.forEach((sp) => {
          if (sp.runEl.parentNode) sp.runEl.parentNode.removeChild(sp.runEl);
        });

        const rebuilt = collectParagraphTextSpans(pEl).spans;
        spans.length = 0;
        rebuilt.forEach((s) => spans.push(s));
        replacedTotal += 1;
      }
    });

    if (replacedTotal > 0) {
      const serializer = new XMLSerializer();
      para.insertOoxml(serializer.serializeToString(xmlDoc), Word.InsertLocation.replace);
      await context.sync();
    }

    return replacedTotal;
  }

  /**
   * Tô màu đỏ cảnh báo tại chỗ cho công thức LaTeX bị lỗi cú pháp trong Word.
   */
  async function markSyntaxErrorInPlace(context, para, rawText) {
    if (!rawText || rawText.length > WORD_SEARCH_MAX_LEN) return;
    try {
      const errRanges = await searchExactInParagraph(context, para, rawText);
      if (!errRanges.items || errRanges.items.length === 0) return;

      for (let i = 0; i < errRanges.items.length; i++) {
        const r = errRanges.items[i];
        r.font.color = "#DC2626";
        r.font.bold = true;
        try {
          r.font.highlightColor = "Yellow";
        } catch {
          // Một số phiên bản Word không hỗ trợ highlightColor
        }
      }
      await context.sync();
    } catch {
      // Không làm gián đoạn tiến trình nếu không thể tô màu
    }
  }

  async function handleScanPreview() {
    if (!ensureWordAvailable() || state.isBusy) return;
    const opts = getScanOptions();
    if (!opts.useInline && !opts.useDisplay) {
      showStatus("Hãy chọn ít nhất một kiểu cú pháp công thức (cùng dòng hoặc riêng dòng).", "warn");
      return;
    }

    setBusy(true);
    try {
      const found = [];
      let autoSwitched = false;

      await Word.run(async (context) => {
        const target = await resolveTargetContainer(context, opts.scope);
        autoSwitched = target.autoSwitchedToBody;

        const paragraphs = target.container.paragraphs;
        paragraphs.load("items/text");
        await context.sync();

        paragraphs.items.forEach((para, idx) => {
          const matches = extractLatexMatches(para.text || "", opts);
          matches.forEach((m) => {
            let syntaxError = "";
            try {
              latexToOmmlElement(m.latex, m.displayMode);
            } catch (e) {
              syntaxError = e && e.message ? e.message : String(e);
            }

            found.push({
              paragraphIndex: idx + 1,
              raw: m.raw,
              latex: m.latex,
              displayMode: m.displayMode,
              syntaxError
            });
          });
        });
      });

      state.lastScanItems = found;
      renderScanResults(found);

      const errTotal = found.filter((f) => f.syntaxError).length;
      const scopeNotice = autoSwitched
        ? "Không có vùng bôi đen — tự động quét Toàn bộ tài liệu. "
        : "";

      if (found.length === 0) {
        showStatus(scopeNotice + "Không tìm thấy công thức LaTeX nào.", "warn");
      } else if (errTotal > 0) {
        showStatus(
          scopeNotice +
            "Phát hiện " +
            found.length +
            " công thức (gồm " +
            errTotal +
            " công thức lỗi cú pháp sẽ bị bỏ qua & tô đỏ).",
          "warn"
        );
      } else {
        showStatus(
          scopeNotice + "Tìm thấy " + found.length + " công thức hợp lệ sẵn sàng chuyển đổi.",
          "success"
        );
      }
    } catch (err) {
      showStatus("Lỗi khi quét tài liệu: " + (err && err.message ? err.message : String(err)), "error");
    } finally {
      setBusy(false);
    }
  }

  async function handleBatchConvert() {
    if (!ensureWordAvailable() || state.isBusy) return;
    const opts = getScanOptions();
    if (!opts.useInline && !opts.useDisplay) {
      showStatus("Hãy chọn ít nhất một kiểu cú pháp nhận diện LaTeX.", "warn");
      return;
    }

    state.cancelRequested = false;
    setBusy(true, true);
    renderErrorReport([]);

    let totalInstances = 0;
    let processedInstances = 0;
    let successCount = 0;
    let errorCount = 0;
    let autoSwitched = false;
    const errorItems = [];

    try {
      await Word.run(async (context) => {
        const target = await resolveTargetContainer(context, opts.scope);
        autoSwitched = target.autoSwitchedToBody;

        const paragraphs = target.container.paragraphs;
        paragraphs.load("items/text");
        await context.sync();

        // Duyệt từng đoạn văn, dùng Regex chuẩn JS trích xuất chính xác danh sách chuỗi LaTeX
        const tasks = [];
        paragraphs.items.forEach((para, paraIdx) => {
          const matches = extractLatexMatches(para.text || "", opts);
          if (matches.length === 0) return;

          // Gom theo chuỗi raw duy nhất trong cùng 1 đoạn để search không bị lặp lại
          const grouped = new Map();
          matches.forEach((m) => {
            if (!grouped.has(m.raw)) {
              grouped.set(m.raw, { match: m, count: 1 });
            } else {
              grouped.get(m.raw).count += 1;
            }
          });

          grouped.forEach((entry) => {
            totalInstances += entry.count;
            tasks.push({
              para,
              paragraphIndex: paraIdx + 1,
              match: entry.match,
              count: entry.count
            });
          });
        });

        if (totalInstances === 0) {
          updateProgressUI(0, 0, 0, 0);
          return;
        }

        updateProgressUI(
          0,
          totalInstances,
          0,
          0,
          "Đang xử lý 0/" + totalInstances + " công thức..."
        );
        await yieldToUi();

        for (let t = 0; t < tasks.length; t++) {
          if (state.cancelRequested) break;

          const { para, paragraphIndex, match, count } = tasks[t];

          // Bước 1: Dùng Temml (chuẩn hóa MathML xmlns + <msup>/<msub>/<msubsup>) -> OMML -> OOXML (<pkg:package>)
          let oMathXml = "";
          let ooxml = "";
          try {
            oMathXml = latexToOmmlElement(match.latex, match.displayMode);
            ooxml = buildWordOoxml(oMathXml, match.displayMode);
          } catch (syntaxErr) {
            // Bỏ qua công thức lỗi, tô màu đỏ cảnh báo tại chỗ và tiếp tục
            if (opts.highlightError) {
              await markSyntaxErrorInPlace(context, para, match.raw);
            }
            errorCount += count;
            processedInstances += count;
            errorItems.push({
              paragraphIndex,
              raw: match.raw,
              latex: match.latex,
              displayMode: match.displayMode,
              message: syntaxErr && syntaxErr.message ? syntaxErr.message : String(syntaxErr)
            });
            updateProgressUI(processedInstances, totalInstances, successCount, errorCount);
            await yieldToUi();
            continue;
          }

          // Bước 2: Định vị bằng paragraph.search(exactString, { matchWildcards: false }) (có xử lý '^' -> '^^' cho Word Desktop)
          // rồi thay thế bằng insertOoxml(..., Word.InsertLocation.replace)
          try {
            const results = await searchExactInParagraph(context, para, match.raw);
            let replacedHere = 0;

            if (results.items && results.items.length > 0) {
              for (let k = 0; k < results.items.length; k++) {
                if (state.cancelRequested) break;
                results.items[k].insertOoxml(ooxml, Word.InsertLocation.replace);
                replacedHere += 1;
              }
              await context.sync();
            } else {
              // Dự phòng khi Word API không định vị được bằng search (vd > 250 ký tự hoặc tách run)
              replacedHere = await replaceRawFormulaViaParaOoxml(
                context,
                para,
                match.raw,
                oMathXml,
                match.displayMode
              );
            }

            if (replacedHere > 0) {
              successCount += replacedHere;
              processedInstances += Math.max(replacedHere, count);
            } else {
              if (opts.highlightError) {
                await markSyntaxErrorInPlace(context, para, match.raw);
              }
              errorCount += count;
              processedInstances += count;
              errorItems.push({
                paragraphIndex,
                raw: match.raw,
                latex: match.latex,
                displayMode: match.displayMode,
                message: "Không định vị được chuỗi công thức trong đoạn văn để chèn Word Equation."
              });
            }
          } catch (insertErr) {
            // Lỗi khi chèn OOXML: tô đỏ cảnh báo tại chỗ và tiếp tục công thức tiếp theo
            if (opts.highlightError) {
              await markSyntaxErrorInPlace(context, para, match.raw);
            }
            errorCount += count;
            processedInstances += count;
            errorItems.push({
              paragraphIndex,
              raw: match.raw,
              latex: match.latex,
              displayMode: match.displayMode,
              message:
                "Không thể chèn OOXML: " +
                (insertErr && insertErr.message ? insertErr.message : String(insertErr))
            });
          }

          updateProgressUI(processedInstances, totalInstances, successCount, errorCount);
          await yieldToUi();
        }
      });

      renderErrorReport(errorItems);
      renderScanResults([]);

      const scopePrefix = autoSwitched
        ? "Không có vùng bôi đen (tự động quét toàn tài liệu). "
        : "";

      if (totalInstances === 0) {
        showStatus(scopePrefix + "Không tìm thấy công thức LaTeX nào để chuyển đổi.", "warn");
      } else if (state.cancelRequested) {
        updateProgressUI(
          processedInstances,
          totalInstances,
          successCount,
          errorCount,
          "Đã dừng ở " + processedInstances + "/" + totalInstances + " công thức"
        );
        showStatus(
          scopePrefix +
            "Đã dừng theo yêu cầu: thành công " +
            successCount +
            ", lỗi " +
            errorCount +
            " công thức.",
          "warn"
        );
      } else {
        updateProgressUI(
          totalInstances,
          totalInstances,
          successCount,
          errorCount,
          "Hoàn tất " + totalInstances + "/" + totalInstances + " công thức"
        );
        if (errorCount > 0) {
          showStatus(
            scopePrefix +
              "Đã chuyển " +
              successCount +
              " công thức. Bỏ qua & tô đỏ " +
              errorCount +
              " công thức bị lỗi cú pháp.",
            "warn"
          );
        } else {
          showStatus(
            scopePrefix +
              "Chuyển đổi hoàn tất " +
              successCount +
              " công thức sang Word Equation.",
            "success"
          );
        }
      }
    } catch (err) {
      updateProgressUI(0, 0, 0, 0);
      showStatus(
        "Lỗi không mong muốn khi chuyển đổi: " + (err && err.message ? err.message : String(err)),
        "error"
      );
    } finally {
      setBusy(false, false);
      state.cancelRequested = false;
    }
  }

  async function handleBatchReverse() {
    if (!ensureWordAvailable() || state.isBusy) return;
    const opts = getScanOptions();

    setBusy(true);
    try {
      let replacedCount = 0;
      let autoSwitched = false;

      await Word.run(async (context) => {
        const target = await resolveTargetContainer(context, opts.scope);
        autoSwitched = target.autoSwitchedToBody;

        const ooxmlRes = target.container.getOoxml();
        await context.sync();

        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(ooxmlRes.value, "application/xml");
        const wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        const allNodes = Array.from(xmlDoc.getElementsByTagName("*"));
        const mathTargets = allNodes.filter(
          (el) =>
            el.localName === "oMathPara" ||
            (el.localName === "oMath" &&
              (!el.parentElement || el.parentElement.localName !== "oMathPara"))
        );

        mathTargets.forEach((node) => {
          const isDisplay = node.localName === "oMathPara";
          const latex = ommlNodeToLatex(node).replace(/\s+/g, " ").trim();
          if (!latex || !node.parentNode) return;

          const run = xmlDoc.createElementNS(wNs, "w:r");
          const textEl = xmlDoc.createElementNS(wNs, "w:t");
          textEl.setAttribute("xml:space", "preserve");
          textEl.textContent = isDisplay ? "$$" + latex + "$$" : "$" + latex + "$";
          run.appendChild(textEl);
          node.parentNode.replaceChild(run, node);
          replacedCount += 1;
        });

        if (replacedCount > 0) {
          const serializer = new XMLSerializer();
          const updatedOoxml = serializer.serializeToString(xmlDoc);
          target.container.insertOoxml(updatedOoxml, Word.InsertLocation.replace);
          await context.sync();
        }
      });

      const prefix = autoSwitched ? "Tự động áp dụng cho toàn tài liệu. " : "";
      if (replacedCount === 0) {
        showStatus(prefix + "Không tìm thấy Word Equation nào để chuyển ngược.", "warn");
      } else {
        showStatus(
          prefix + "Đã chuyển ngược " + replacedCount + " Word Equation về mã LaTeX.",
          "success"
        );
      }
    } catch (err) {
      showStatus(
        "Lỗi chuyển ngược Equation: " + (err && err.message ? err.message : String(err)),
        "error"
      );
    } finally {
      setBusy(false);
    }
  }

  async function handleInsertEquation(overrideLatex, overrideDisplay) {
    if (!ensureWordAvailable() || state.isBusy) return;

    const rawLatex = typeof overrideLatex === "string" ? overrideLatex : dom.latexInput.value;
    const latex = stripOuterDelimiters(rawLatex);
    const displayMode =
      typeof overrideDisplay === "boolean" ? overrideDisplay : state.displayMode;

    if (!latex) {
      showStatus("Vui lòng nhập công thức LaTeX trước khi chèn.", "warn");
      return;
    }

    setBusy(true);
    try {
      const ooxml = latexToWordOoxml(latex, displayMode);

      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertOoxml(ooxml, Word.InsertLocation.replace);
        await context.sync();
      });

      addToHistory(latex, displayMode);
      showStatus("Đã chèn công thức Equation vào Word thành công.", "success");
    } catch (err) {
      showStatus(
        "Không thể chèn công thức: " + (err && err.message ? err.message : String(err)),
        "error"
      );
    } finally {
      setBusy(false);
    }
  }

  async function handleInsertRawLatexText() {
    if (!ensureWordAvailable() || state.isBusy) return;
    const latex = stripOuterDelimiters(dom.latexInput.value);
    if (!latex) {
      showStatus("Vui lòng nhập mã LaTeX trước khi chèn.", "warn");
      return;
    }

    const wrapped = state.displayMode ? "$$" + latex + "$$" : "$" + latex + "$";
    setBusy(true);
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertText(wrapped, Word.InsertLocation.replace);
        await context.sync();
      });
      showStatus("Đã chèn văn bản mã nguồn LaTeX vào Word.", "success");
    } catch (err) {
      showStatus(
        "Lỗi khi chèn văn bản: " + (err && err.message ? err.message : String(err)),
        "error"
      );
    } finally {
      setBusy(false);
    }
  }

  async function handleReadEquationFromWord() {
    if (!ensureWordAvailable() || state.isBusy) return;

    setBusy(true);
    try {
      let ooxmlValue = "";
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const ooxml = selection.getOoxml();
        await context.sync();
        ooxmlValue = ooxml.value || "";
      });

      const extracted = extractLatexFromOoxml(ooxmlValue);
      if (!extracted || !extracted.latex) {
        showStatus("Vùng đang chọn chưa có công thức Word Equation (OMML).", "warn");
        return;
      }

      dom.latexInput.value = extracted.latex;
      state.displayMode = extracted.isDisplay;
      dom.modeBtns.forEach((b) =>
        b.classList.toggle("active", (b.dataset.mode === "display") === state.displayMode)
      );
      updateLivePreview();
      showStatus("Đã đọc công thức từ Word sang LaTeX.", "success");
    } catch (err) {
      showStatus(
        "Không thể đọc công thức từ Word: " + (err && err.message ? err.message : String(err)),
        "error"
      );
    } finally {
      setBusy(false);
    }
  }

  async function handleCopyMathML() {
    if (!state.lastMathML) {
      showStatus("Chưa có MathML hợp lệ để sao chép.", "warn");
      return;
    }
    try {
      await navigator.clipboard.writeText(state.lastMathML);
      showStatus("Đã sao chép mã MathML vào khay nhớ tạm.", "success");
    } catch {
      showStatus("Trình duyệt không cho phép truy cập clipboard tự động.", "warn");
    }
  }

  /* ------------------------------------------------------------------------
   * 7. HIỂN THỊ BÁO CÁO LỖI, KẾT QUẢ QUÉT, MẪU & LỊCH SỬ
   * ------------------------------------------------------------------------ */
  function openFormulaInEditor(latex, displayMode) {
    dom.latexInput.value = latex;
    state.displayMode = Boolean(displayMode);
    dom.modeBtns.forEach((b) =>
      b.classList.toggle("active", (b.dataset.mode === "display") === state.displayMode)
    );
    updateLivePreview();
    switchTab("editor");
  }

  function renderErrorReport(items) {
    if (!items || items.length === 0) {
      dom.errorReportCard.classList.add("hidden");
      dom.errorReportList.innerHTML = "";
      return;
    }

    dom.errorReportCard.classList.remove("hidden");
    dom.errorCountLabel.textContent = items.length + " lỗi";
    dom.errorReportList.innerHTML = "";

    items.forEach((item, index) => {
      const card = document.createElement("div");
      card.className = "list-item is-error";

      const head = document.createElement("div");
      head.className = "list-item-head";

      const title = document.createElement("span");
      title.className = "list-item-title";
      title.textContent = "Lỗi #" + (index + 1) + " · Đoạn " + item.paragraphIndex;

      const meta = document.createElement("span");
      meta.className = "list-item-meta";
      meta.textContent = "Đã tô đỏ";

      head.appendChild(title);
      head.appendChild(meta);

      const code = document.createElement("div");
      code.className = "list-item-code";
      code.textContent = item.raw;
      code.title = item.raw;

      const reason = document.createElement("div");
      reason.className = "list-item-err-msg";
      reason.textContent = item.message;

      const actions = document.createElement("div");
      actions.className = "list-item-actions";
      const btnFix = document.createElement("button");
      btnFix.type = "button";
      btnFix.className = "mini-btn";
      btnFix.textContent = "Mở sửa trong Soạn thảo";
      btnFix.addEventListener("click", () => openFormulaInEditor(item.latex, item.displayMode));
      actions.appendChild(btnFix);

      card.appendChild(head);
      card.appendChild(code);
      card.appendChild(reason);
      card.appendChild(actions);
      dom.errorReportList.appendChild(card);
    });
  }

  function renderScanResults(items) {
    dom.scanCountLabel.textContent = items.length > 0 ? items.length + " công thức" : "0 công thức";
    dom.scanResultsList.innerHTML = "";

    if (items.length === 0) {
      dom.scanResultsList.innerHTML =
        '<div class="empty-note">Không có công thức LaTeX chờ chuyển đổi.</div>';
      return;
    }

    items.forEach((item, index) => {
      const card = document.createElement("div");
      card.className = "list-item" + (item.syntaxError ? " is-error" : "");

      const head = document.createElement("div");
      head.className = "list-item-head";

      const title = document.createElement("span");
      title.className = "list-item-title";
      title.textContent = "#" + (index + 1) + " · Đoạn " + item.paragraphIndex;

      const meta = document.createElement("span");
      meta.className = "list-item-meta";
      meta.textContent = item.syntaxError
        ? "Lỗi cú pháp"
        : item.displayMode
          ? "Riêng dòng"
          : "Cùng dòng";

      head.appendChild(title);
      head.appendChild(meta);

      const code = document.createElement("div");
      code.className = "list-item-code";
      code.textContent = item.raw;
      code.title = item.raw;

      card.appendChild(head);
      card.appendChild(code);

      if (item.syntaxError) {
        const errMsg = document.createElement("div");
        errMsg.className = "list-item-err-msg";
        errMsg.textContent = item.syntaxError;
        card.appendChild(errMsg);
      }

      const actions = document.createElement("div");
      actions.className = "list-item-actions";
      const btnEdit = document.createElement("button");
      btnEdit.type = "button";
      btnEdit.className = "mini-btn";
      btnEdit.textContent = "Mở trong Soạn thảo";
      btnEdit.addEventListener("click", () => openFormulaInEditor(item.latex, item.displayMode));
      actions.appendChild(btnEdit);

      card.appendChild(actions);
      dom.scanResultsList.appendChild(card);
    });
  }

  function renderTemplates() {
    dom.templateList.innerHTML = "";
    MATH_TEMPLATES.forEach((tpl) => {
      const card = document.createElement("div");
      card.className = "list-item";

      const head = document.createElement("div");
      head.className = "list-item-head";

      const title = document.createElement("span");
      title.className = "list-item-title";
      title.textContent = tpl.title;

      const meta = document.createElement("span");
      meta.className = "list-item-meta";
      meta.textContent = tpl.display ? "Riêng dòng" : "Cùng dòng";

      head.appendChild(title);
      head.appendChild(meta);

      const code = document.createElement("div");
      code.className = "list-item-code";
      code.textContent = tpl.latex;
      code.title = tpl.latex;

      const actions = document.createElement("div");
      actions.className = "list-item-actions";

      const btnLoad = document.createElement("button");
      btnLoad.type = "button";
      btnLoad.className = "mini-btn";
      btnLoad.textContent = "Nạp vào ô sửa";
      btnLoad.addEventListener("click", () => openFormulaInEditor(tpl.latex, tpl.display));

      const btnQuickInsert = document.createElement("button");
      btnQuickInsert.type = "button";
      btnQuickInsert.className = "mini-btn accent";
      btnQuickInsert.textContent = "Chèn ngay";
      btnQuickInsert.addEventListener("click", () => handleInsertEquation(tpl.latex, tpl.display));

      actions.appendChild(btnLoad);
      actions.appendChild(btnQuickInsert);

      card.appendChild(head);
      card.appendChild(code);
      card.appendChild(actions);
      dom.templateList.appendChild(card);
    });
  }

  function loadHistory() {
    try {
      const raw = localStorage.getItem(HISTORY_STORAGE_KEY);
      const parsed = raw ? JSON.parse(raw) : [];
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  function saveHistory(items) {
    try {
      localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(items.slice(0, MAX_HISTORY_ITEMS)));
    } catch {
      // Bỏ qua nếu WebView hạn chế localStorage
    }
  }

  function addToHistory(latex, displayMode) {
    const items = loadHistory().filter((it) => it.latex !== latex);
    items.unshift({ latex, displayMode, savedAt: Date.now() });
    saveHistory(items);
    renderHistory();
  }

  function renderHistory() {
    const items = loadHistory();
    dom.historyList.innerHTML = "";
    if (items.length === 0) {
      dom.historyList.innerHTML =
        '<div class="empty-note">Chưa có công thức nào được chèn gần đây.</div>';
      return;
    }

    items.forEach((item) => {
      const card = document.createElement("div");
      card.className = "list-item";

      const code = document.createElement("div");
      code.className = "list-item-code";
      code.textContent = item.latex;
      code.title = item.latex;

      const actions = document.createElement("div");
      actions.className = "list-item-actions";

      const btnLoad = document.createElement("button");
      btnLoad.type = "button";
      btnLoad.className = "mini-btn";
      btnLoad.textContent = "Sửa lại";
      btnLoad.addEventListener("click", () =>
        openFormulaInEditor(item.latex, Boolean(item.displayMode))
      );

      const btnInsert = document.createElement("button");
      btnInsert.type = "button";
      btnInsert.className = "mini-btn accent";
      btnInsert.textContent = "Chèn lại";
      btnInsert.addEventListener("click", () =>
        handleInsertEquation(item.latex, Boolean(item.displayMode))
      );

      actions.appendChild(btnLoad);
      actions.appendChild(btnInsert);
      card.appendChild(code);
      card.appendChild(actions);
      dom.historyList.appendChild(card);
    });
  }

  /* ------------------------------------------------------------------------
   * 8. MODULE ĐỘC LẬP: CHUYỂN ĐỔI MATHTYPE CŨ (OLE2 CFB / MTEF) -> LATEX
   *    Giải mã trực tiếp nhị phân oleObject*.bin bằng Uint8Array & DataView
   *    không cần cài đặt phần mềm MathType trên máy tính.
   * ------------------------------------------------------------------------ */
  const MATHTYPE_PROGID_RE = /^(Equation\.DSMT\d*|Equation\.\d+|Equation\.KSE|MathType)/i;

  // Ánh xạ mã ký tự phông Symbol (dùng cho MTEF 3 / Equation 3.0 cũ)
  const SYMBOL_FONT_TO_UNICODE = {
    0x22: "∀", 0x24: "∃", 0x27: "∋", 0x2a: "∗", 0x2d: "−",
    0x40: "≅", 0x41: "Α", 0x42: "Β", 0x43: "Χ", 0x44: "Δ",
    0x45: "Ε", 0x46: "Φ", 0x47: "Γ", 0x48: "Η", 0x49: "Ι",
    0x4b: "Κ", 0x4c: "Λ", 0x4d: "Μ", 0x4e: "Ν", 0x4f: "Ο",
    0x50: "Π", 0x51: "Θ", 0x52: "Ρ", 0x53: "Σ", 0x54: "Τ",
    0x55: "Υ", 0x57: "Ω", 0x58: "Ξ", 0x59: "Ψ", 0x5a: "Ζ",
    0x5c: "∴", 0x5e: "⊥", 0x61: "α", 0x62: "β", 0x63: "χ",
    0x64: "δ", 0x65: "ε", 0x66: "φ", 0x67: "γ", 0x68: "η",
    0x69: "ι", 0x6b: "κ", 0x6c: "λ", 0x6d: "μ", 0x6e: "ν",
    0x6f: "ο", 0x70: "π", 0x71: "θ", 0x72: "ρ", 0x73: "σ",
    0x74: "τ", 0x75: "υ", 0x77: "ω", 0x78: "ξ", 0x79: "ψ",
    0x7a: "ζ", 0x7e: "∼", 0xa3: "≤", 0xa5: "∞", 0xab: "↔",
    0xac: "←", 0xad: "↑", 0xae: "→", 0xaf: "↓", 0xb0: "°",
    0xb1: "±", 0xb3: "≥", 0xb4: "×", 0xb5: "∝", 0xb6: "∂",
    0xb7: "·", 0xb8: "÷", 0xb9: "≠", 0xba: "≡", 0xbb: "≈",
    0xc4: "⊗", 0xc5: "⊕", 0xc6: "∅", 0xc7: "∩", 0xc8: "∪",
    0xc9: "⊃", 0xca: "⊇", 0xcc: "⊂", 0xcd: "⊆", 0xce: "∈",
    0xcf: "∉", 0xd0: "∠", 0xd1: "∇", 0xd5: "∏", 0xd6: "√",
    0xd7: "⋅", 0xd8: "¬", 0xd9: "∧", 0xda: "∨", 0xdb: "⇔",
    0xdc: "⇐", 0xde: "⇒", 0xe5: "∑", 0xf2: "∫"
  };

  function base64ToUint8Array(base64Str) {
    const clean = (base64Str || "").replace(/\s+/g, "");
    const binary = window.atob(clean);
    const len = binary.length;
    const out = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      out[i] = binary.charCodeAt(i);
    }
    return out;
  }

  /**
   * Bộ giải mã OLE2 Compound File Binary Format (MS-CFB) bằng Uint8Array & DataView.
   * Trích xuất chính xác luồng "Equation Native" (hỗ trợ cả FAT lẫn MiniFAT/MiniStream).
   */
  function extractEquationNativeFromCfb(bytes) {
    if (!bytes || bytes.length < 512) {
      throw new Error("Tệp nhị phân OLE quá ngắn (< 512 byte).");
    }
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

    // Kiểm tra chữ ký OLE2 CFB: D0 CF 11 E0 A1 B1 1A E1
    const sigLo = view.getUint32(0, true);
    const sigHi = view.getUint32(4, true);
    if (sigLo !== 0xe011cfd0 || sigHi !== 0xe11ab1a1) {
      throw new Error("Không phải định dạng OLE2 Compound File hợp lệ.");
    }

    const sectorShift = view.getUint16(30, true);
    const miniSectorShift = view.getUint16(32, true);
    const sectorSize = 1 << sectorShift;
    const miniSectorSize = 1 << miniSectorShift;

    const firstDirSector = view.getUint32(48, true);
    const miniStreamCutoff = view.getUint32(56, true);
    const firstMiniFatSector = view.getUint32(60, true);
    const firstDifatSector = view.getUint32(68, true);
    const numDifatSectors = view.getUint32(72, true);

    const ENDOFCHAIN = 0xfffffffe;
    const FREESECT = 0xffffffff;
    const MAXREGSECT = 0xfffffffa;

    // 1. Đọc danh sách các sector FAT từ DIFAT (109 mục trong header + chuỗi DIFAT mở rộng)
    const fatSectorIds = [];
    for (let i = 0; i < 109; i++) {
      const sid = view.getUint32(76 + i * 4, true);
      if (sid <= MAXREGSECT) fatSectorIds.push(sid);
    }

    let difatSid = firstDifatSector;
    let difatVisited = 0;
    const difatEntriesPerSector = sectorSize / 4 - 1;
    while (difatSid <= MAXREGSECT && difatVisited < numDifatSectors && difatVisited < 1024) {
      const off = (difatSid + 1) * sectorSize;
      if (off + sectorSize > bytes.length) break;
      for (let i = 0; i < difatEntriesPerSector; i++) {
        const sid = view.getUint32(off + i * 4, true);
        if (sid <= MAXREGSECT) fatSectorIds.push(sid);
      }
      difatSid = view.getUint32(off + difatEntriesPerSector * 4, true);
      difatVisited++;
    }

    // 2. Xây dựng bảng FAT đầy đủ
    const entriesPerSector = sectorSize / 4;
    const fat = new Uint32Array(fatSectorIds.length * entriesPerSector);
    for (let s = 0; s < fatSectorIds.length; s++) {
      const off = (fatSectorIds[s] + 1) * sectorSize;
      if (off + sectorSize > bytes.length) continue;
      for (let i = 0; i < entriesPerSector; i++) {
        fat[s * entriesPerSector + i] = view.getUint32(off + i * 4, true);
      }
    }

    function readFatChain(startSid, maxBytes) {
      const chunks = [];
      let total = 0;
      let sid = startSid;
      const visited = new Set();

      while (sid <= MAXREGSECT && !visited.has(sid)) {
        visited.add(sid);
        const off = (sid + 1) * sectorSize;
        if (off >= bytes.length) break;
        const end = Math.min(bytes.length, off + sectorSize);
        chunks.push(bytes.subarray(off, end));
        total += end - off;
        if (maxBytes !== undefined && total >= maxBytes) break;
        sid = sid < fat.length ? fat[sid] : ENDOFCHAIN;
      }

      const finalLen = maxBytes !== undefined ? Math.min(total, maxBytes) : total;
      const out = new Uint8Array(finalLen);
      let cursor = 0;
      for (let i = 0; i < chunks.length && cursor < finalLen; i++) {
        const take = Math.min(chunks[i].length, finalLen - cursor);
        out.set(chunks[i].subarray(0, take), cursor);
        cursor += take;
      }
      return out;
    }

    // 3. Đọc bảng thư mục (Directory Entries, mỗi mục 128 byte)
    const dirBytes = readFatChain(firstDirSector);
    const dirView = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength);
    const entryCount = Math.floor(dirBytes.length / 128);

    let rootEntry = null;
    let eqNativeEntry = null;

    for (let idx = 0; idx < entryCount; idx++) {
      const base = idx * 128;
      const nameLen = dirView.getUint16(base + 64, true);
      const entryType = dirView.getUint8(base + 66);
      if (entryType === 0) continue;

      const charCount = Math.max(0, Math.min(32, Math.floor((nameLen - 2) / 2)));
      let name = "";
      for (let c = 0; c < charCount; c++) {
        const ch = dirView.getUint16(base + c * 2, true);
        if (ch === 0) break;
        name += String.fromCharCode(ch);
      }

      const startSector = dirView.getUint32(base + 116, true);
      const streamSize = dirView.getUint32(base + 120, true);
      const info = { name, entryType, startSector, streamSize };

      if (entryType === 5 && !rootEntry) {
        rootEntry = info;
      }
      if (entryType === 2 && name.trim().toLowerCase() === "equation native") {
        eqNativeEntry = info;
      }
    }

    if (!eqNativeEntry) {
      throw new Error('Không tìm thấy luồng "Equation Native" trong tệp OLE MathType.');
    }

    // 4. Đọc dữ liệu từ MiniStream (nếu kích thước < miniStreamCutoff) hoặc trực tiếp từ FAT
    if (
      eqNativeEntry.streamSize < miniStreamCutoff &&
      rootEntry &&
      rootEntry.startSector <= MAXREGSECT &&
      firstMiniFatSector <= MAXREGSECT
    ) {
      const miniStream = readFatChain(rootEntry.startSector, rootEntry.streamSize);
      const miniFatBytes = readFatChain(firstMiniFatSector);
      const miniFatView = new DataView(
        miniFatBytes.buffer,
        miniFatBytes.byteOffset,
        miniFatBytes.byteLength
      );
      const miniFatLen = Math.floor(miniFatBytes.length / 4);
      const miniFat = new Uint32Array(miniFatLen);
      for (let i = 0; i < miniFatLen; i++) {
        miniFat[i] = miniFatView.getUint32(i * 4, true);
      }

      const out = new Uint8Array(eqNativeEntry.streamSize);
      let cursor = 0;
      let msid = eqNativeEntry.startSector;
      const visited = new Set();

      while (
        msid <= MAXREGSECT &&
        msid !== ENDOFCHAIN &&
        msid !== FREESECT &&
        !visited.has(msid) &&
        cursor < out.length
      ) {
        visited.add(msid);
        const off = msid * miniSectorSize;
        if (off >= miniStream.length) break;
        const take = Math.min(miniSectorSize, miniStream.length - off, out.length - cursor);
        out.set(miniStream.subarray(off, off + take), cursor);
        cursor += take;
        msid = msid < miniFat.length ? miniFat[msid] : ENDOFCHAIN;
      }

      return out;
    }

    return readFatChain(eqNativeEntry.startSector, eqNativeEntry.streamSize);
  }

  /**
   * Bộ giải mã nhị phân MTEF (MathType Equation Format v3 & v5) sang chuỗi LaTeX.
   */
  function parseMtefToLatex(nativeBytes) {
    if (!nativeBytes || nativeBytes.length < 6) {
      throw new Error("Luồng Equation Native trống hoặc không hợp lệ.");
    }

    const view = new DataView(
      nativeBytes.buffer,
      nativeBytes.byteOffset,
      nativeBytes.byteLength
    );

    let pos = 0;
    let limit = nativeBytes.length;

    // Bỏ qua header EQNOLEFILEHDR (28 byte chuẩn MathType OLE)
    if (nativeBytes.length >= 28 && view.getUint16(0, true) === 28) {
      const cbObject = view.getUint32(8, true);
      pos = 28;
      if (cbObject > 0 && 28 + cbObject <= nativeBytes.length) {
        limit = 28 + cbObject;
      }
    }

    function eof() {
      return pos >= limit;
    }

    function readU8() {
      if (pos >= limit) return 0;
      return nativeBytes[pos++];
    }

    function readU16() {
      if (pos + 2 > limit) {
        pos = limit;
        return 0;
      }
      const val = view.getUint16(pos, true);
      pos += 2;
      return val;
    }

    function readI16() {
      if (pos + 2 > limit) {
        pos = limit;
        return 0;
      }
      const val = view.getInt16(pos, true);
      pos += 2;
      return val;
    }

    function readMtefUInt() {
      const b = readU8();
      return b === 255 ? readU16() : b;
    }

    function readCString() {
      let str = "";
      while (!eof()) {
        const b = readU8();
        if (b === 0) break;
        str += String.fromCharCode(b);
      }
      return str;
    }

    function readNudge() {
      const dx = readU8();
      const dy = readU8();
      if (dx === 128 && dy === 128) {
        readI16();
        readI16();
      }
    }

    function skipDimensionArray(count) {
      if (!count || count <= 0) return;
      let done = 0;
      while (done < count && !eof()) {
        const b = readU8();
        const hi = (b >> 4) & 0x0f;
        const lo = b & 0x0f;
        if (hi === 0x0f) {
          done++;
          if (done >= count) break;
        }
        if (lo === 0x0f) {
          done++;
        }
      }
    }

    function skipRuler() {
      const nStops = readU8();
      for (let i = 0; i < nStops; i++) {
        readU8();  // tab type
        readU16(); // tab offset
      }
    }

    // Đọc MTEF Header
    const mtefVersion = readU8();
    if (mtefVersion < 1 || mtefVersion > 6) {
      throw new Error("Phiên bản MTEF không được hỗ trợ (" + mtefVersion + ").");
    }
    readU8(); // platform
    readU8(); // product
    readU8(); // productVersion
    readU8(); // productSubversion

    let isInline = true;
    if (mtefVersion >= 5) {
      readCString(); // application key (ví dụ: "DSMT4")
      const eqnOpts = readU8();
      isInline = eqnOpts !== 0;
    }

    function mtCodeToLatex(code, typeface, fontPos) {
      let ch = "";
      if (code && code > 0) {
        // Bảng chữ in đậm/tập hợp số PUA của MathType (0xEF00..)
        if (code === 0xef00 || code === 0x211d) return "\\mathbb{R}";
        if (code === 0x2124) return "\\mathbb{Z}";
        if (code === 0x2115) return "\\mathbb{N}";
        if (code === 0x211a) return "\\mathbb{Q}";
        if (code === 0x2102) return "\\mathbb{C}";
        ch = String.fromCharCode(code);
      } else if (fontPos > 0) {
        if (typeface === 4 || typeface === 5 || typeface === 6) {
          ch = SYMBOL_FONT_TO_UNICODE[fontPos] || String.fromCharCode(fontPos);
        } else {
          ch = String.fromCharCode(fontPos);
        }
      }

      if (!ch) return "";
      if (UNICODE_TO_LATEX[ch] !== undefined) return UNICODE_TO_LATEX[ch];
      if (ch === "{") return "\\{";
      if (ch === "}") return "\\}";
      if (ch === "%") return "\\%";
      return ch;
    }

    function applyEmbellishments(base, embells) {
      let out = base;
      for (let i = 0; i < embells.length; i++) {
        switch (embells[i]) {
          case 2: out = "\\dot{" + out + "}"; break;
          case 3: out = "\\ddot{" + out + "}"; break;
          case 5: out = out + "'"; break;
          case 6: out = out + "''"; break;
          case 8: out = "\\tilde{" + out + "}"; break;
          case 9: out = "\\hat{" + out + "}"; break;
          case 10: out = "\\not " + out; break;
          case 11: out = "\\vec{" + out + "}"; break;
          case 13: out = "\\bar{" + out + "}"; break;
          case 14: out = "\\overline{" + out + "}"; break;
          default: break;
        }
      }
      return out;
    }

    const MTEF_LATEX_FUNCTIONS = new Set([
      "sin", "cos", "tan", "cot", "sec", "csc",
      "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh",
      "log", "ln", "lim", "max", "min", "exp", "det", "deg",
      "gcd", "sup", "inf", "arg", "dim", "hom", "ker", "Pr"
    ]);

    function formatFunctionNames(latexStr) {
      return latexStr.replace(
        /(^|[^\\a-zA-Z])(sin|cos|tan|cot|log|ln|lim|max|min|exp)(?![a-zA-Z])/g,
        "$1\\$2 "
      );
    }

    function appendLatexToken(acc, token) {
      if (!token) return acc;
      if (/\\[a-zA-Z]+$/.test(acc) && /^[a-zA-Z]/.test(token)) {
        return acc + " " + token;
      }
      return acc + token;
    }

    function joinLineChildren(children) {
      let out = "";
      let fnBuffer = "";

      function flushFnBuffer() {
        if (!fnBuffer) return;
        if (MTEF_LATEX_FUNCTIONS.has(fnBuffer)) {
          out = appendLatexToken(out, "\\" + fnBuffer + " ");
        } else {
          out = appendLatexToken(out, fnBuffer);
        }
        fnBuffer = "";
      }

      for (let i = 0; i < children.length; i++) {
        const ch = children[i];
        if (ch && ch.kind === "char" && ch.isFn) {
          fnBuffer += ch.text;
          continue;
        }
        flushFnBuffer();
        const piece = ch && ch.kind === "char" ? (ch.text || "") : slotToText(ch);
        out = appendLatexToken(out, piece);
      }
      flushFnBuffer();
      return formatFunctionNames(out).replace(/\s{2,}/g, " ").trim();
    }

    function slotToText(slot) {
      if (!slot) return "";
      if (typeof slot === "string") return slot.trim();
      if (slot.kind === "pile") return slot.rows.join(" \\\\ ");
      if (slot.kind === "matrix") {
        const matrixRows = [];
        for (let r = 0; r < slot.rows; r++) {
          const rowCells = [];
          for (let c = 0; c < slot.cols; c++) {
            rowCells.push(slotToText(slot.cells[r * slot.cols + c]));
          }
          matrixRows.push(rowCells.join(" & "));
        }
        return "\\begin{matrix} " + matrixRows.join(" \\\\ ") + " \\end{matrix}";
      }
      if (slot.kind === "char") return String(slot.text || "");
      return String(slot.text || "").trim();
    }

    function renderTemplate(selector, variation, slots) {
      const s0 = slotToText(slots[0]);
      const s1 = slotToText(slots[1]);
      const s2 = slotToText(slots[2]);

      switch (selector) {
        case 0: // tmANGLE
          return "\\left\\langle " + s0 + " \\right\\rangle";
        case 1: // tmPAREN
          if (slots[0] && slots[0].kind === "matrix") {
            return slotToText(slots[0]).replace(/\{matrix\}/g, "{pmatrix}");
          }
          return "\\left( " + s0 + " \\right)";
        case 2: { // tmBRACE (hỗ trợ cả hệ phương trình \begin{cases})
          const hasLeft = (variation & 0x01) !== 0 || variation === 0;
          const hasRight = (variation & 0x02) !== 0;
          if (hasLeft && !hasRight && slots[0] && (slots[0].kind === "pile" || slots[0].kind === "matrix")) {
            const body =
              slots[0].kind === "pile"
                ? slots[0].rows.join(" \\\\ ")
                : slotToText(slots[0])
                    .replace(/^\\begin\{matrix\}\s*/, "")
                    .replace(/\s*\\end\{matrix\}$/, "");
            return "\\begin{cases} " + body + " \\end{cases}";
          }
          const leftB = hasLeft ? "\\left\\{" : "\\left.";
          const rightB = hasRight ? "\\right\\}" : "\\right.";
          return leftB + " " + s0 + " " + rightB;
        }
        case 3: // tmBRACK
          if (slots[0] && slots[0].kind === "matrix") {
            return slotToText(slots[0]).replace(/\{matrix\}/g, "{bmatrix}");
          }
          return "\\left[ " + s0 + " \\right]";
        case 4: // tmBAR
          return "\\left| " + s0 + " \\right|";
        case 5: // tmDBAR
          return "\\left\\| " + s0 + " \\right\\|";
        case 6: // tmFLOOR
          return "\\left\\lfloor " + s0 + " \\right\\rfloor";
        case 7: // tmCEILING
          return "\\left\\lceil " + s0 + " \\right\\rceil";
        case 11: // tmROOT (căn bậc hai / căn bậc n)
          return variation === 1 && s1
            ? "\\sqrt[" + s1 + "]{" + s0 + "}"
            : "\\sqrt{" + s0 + "}";
        case 12: // tmFRACT (phân số)
          return "\\frac{" + s0 + "}{" + s1 + "}";
        case 13: // tmUBAR
          return "\\underline{" + s0 + "}";
        case 14: // tmOBAR
          return "\\overline{" + s0 + "}";
        case 15: // tmARROW
        case 31: // tmVEC
          return s0.length <= 1 ? "\\vec{" + s0 + "}" : "\\overrightarrow{" + s0 + "}";
        case 16: { // tmINTEG (tích phân)
          let cmd = "\\int";
          if (variation & 0x04) cmd = "\\iint";
          else if (variation & 0x08) cmd = "\\iiint";
          else if (variation & 0x10) cmd = "\\oint";
          return cmd + (s1 ? "_{" + s1 + "}" : "") + (s2 ? "^{" + s2 + "}" : "") + (s0 ? " " + s0 : "");
        }
        case 17: // tmSUM
          return "\\sum" + (s1 ? "_{" + s1 + "}" : "") + (s2 ? "^{" + s2 + "}" : "") + (s0 ? " " + s0 : "");
        case 18: // tmPROD
          return "\\prod" + (s1 ? "_{" + s1 + "}" : "") + (s2 ? "^{" + s2 + "}" : "") + (s0 ? " " + s0 : "");
        case 19: // tmCOPROD
          return "\\coprod" + (s1 ? "_{" + s1 + "}" : "") + (s2 ? "^{" + s2 + "}" : "") + (s0 ? " " + s0 : "");
        case 20: // tmUNION
          return "\\bigcup" + (s1 ? "_{" + s1 + "}" : "") + (s2 ? "^{" + s2 + "}" : "") + (s0 ? " " + s0 : "");
        case 21: // tmINTER
          return "\\bigcap" + (s1 ? "_{" + s1 + "}" : "") + (s2 ? "^{" + s2 + "}" : "") + (s0 ? " " + s0 : "");
        case 24: // tmLIM (giới hạn)
          return (s0 || "\\lim") + (s1 ? "_{" + s1 + "}" : "") + (s2 ? "^{" + s2 + "}" : "") + " ";
        case 25: // tmHBRACE
          return (variation & 0x01)
            ? "\\overbrace{" + s0 + "}" + (s1 ? "^{" + s1 + "}" : "")
            : "\\underbrace{" + s0 + "}" + (s1 ? "_{" + s1 + "}" : "");
        case 27: // tmSUB (chỉ số dưới: slot 0 = sub, slot 1 = L_NULL)
          return "_" + formatLatexScriptArg(s0 || s1);
        case 28: // tmSUP (số mũ: hỗ trợ cả chuẩn ScrBoxClass [L_NULL, sup] lẫn 1-slot [sup])
          return "^" + formatLatexScriptArg(s1 || s0);
        case 29: // tmSUBSUP (cả chỉ số dưới & số mũ)
          return (s0 ? "_" + formatLatexScriptArg(s0) : "") + (s1 ? "^" + formatLatexScriptArg(s1) : "");
        case 30: // tmDIRAC
          return "\\left\\langle " + s0 + " \\middle| " + s1 + " \\right\\rangle";
        case 32: // tmHAT
          return "\\widehat{" + s0 + "}";
        default:
          return slots.map(slotToText).filter(Boolean).join(" ");
      }
    }

    function readRecordList(depth) {
      if (depth > 64) return [];
      const items = [];

      while (!eof()) {
        const rawTag = readU8();
        let recType = rawTag;
        let opts = 0;

        if (mtefVersion < 5) {
          recType = rawTag & 0x0f;
          opts = (rawTag >> 4) & 0x0f;
        }

        if (recType === 0) {
          // END record
          break;
        }

        // FUTURE record (>= 100 trong MTEF 5)
        if (mtefVersion >= 5 && recType >= 100) {
          const skipLen = readMtefUInt();
          pos = Math.min(limit, pos + skipLen);
          continue;
        }

        switch (recType) {
          case 1: { // LINE
            if (mtefVersion >= 5) opts = readU8();
            if (opts & 0x08) readNudge();
            if (opts & 0x04) readU16(); // line spacing
            if (opts & 0x02) skipRuler();
            if (opts & 0x01) {
              // L_NULL: slot rỗng, không có danh sách con
              items.push({ kind: "line", text: "" });
            } else {
              const children = readRecordList(depth + 1);
              if (children.length === 1 && (children[0].kind === "pile" || children[0].kind === "matrix")) {
                items.push(children[0]);
              } else {
                items.push({ kind: "line", text: joinLineChildren(children) });
              }
            }
            break;
          }

          case 2: { // CHAR
            if (mtefVersion >= 5) opts = readU8();
            if (opts & 0x08) readNudge();

            const rawTypeface = readU8();
            const typeface = mtefVersion >= 5 ? rawTypeface - 128 : rawTypeface;
            let mtCode = 0;
            let fontPos = 0;

            if (mtefVersion >= 5) {
              if (!(opts & 0x20)) mtCode = readU16(); // xfENC_NO_MTCODE
              if (opts & 0x04) fontPos = readU8();    // xfENC_CHAR_8
              else if (opts & 0x10) fontPos = readU16(); // xfENC_CHAR_16
            } else {
              fontPos = readU16();
            }

            const embells = [];
            if (opts & 0x01) {
              // Danh sách EMBELL kết thúc bằng END
              const embRecords = readRecordList(depth + 1);
              embRecords.forEach((er) => {
                if (er && er.kind === "embell") embells.push(er.embellType);
              });
            }

            const baseChar = mtCodeToLatex(mtCode, typeface, fontPos);
            const finalChar = embells.length > 0 ? applyEmbellishments(baseChar, embells) : baseChar;
            const isFn = typeface === 2 && embells.length === 0 && /^[a-zA-Z]$/.test(finalChar);
            items.push({ kind: "char", text: finalChar, isFn });
            break;
          }

          case 3: { // TMPL (Template: phân số, căn, mũ, tích phân, tổng, ngoặc...)
            if (mtefVersion >= 5) opts = readU8();
            if (opts & 0x08) readNudge();
            const selector = readU8();
            let variation = readU8();
            if (variation & 0x80) {
              const hi = readU8();
              variation = (variation & 0x7f) | (hi << 8);
            }
            readU8(); // template-specific options
            const slots = readRecordList(depth + 1);
            items.push({ kind: "text", text: renderTemplate(selector, variation, slots) });
            break;
          }

          case 4: { // PILE (xếp chồng nhiều dòng, dùng trong hệ phương trình)
            if (mtefVersion >= 5) opts = readU8();
            if (opts & 0x08) readNudge();
            readU8(); // halign
            readU8(); // valign
            if (opts & 0x02) skipRuler();
            const lines = readRecordList(depth + 1);
            items.push({ kind: "pile", rows: lines.map(slotToText) });
            break;
          }

          case 5: { // MATRIX
            if (mtefVersion >= 5) opts = readU8();
            if (opts & 0x08) readNudge();
            readU8(); // valign
            readU8(); // h_just
            readU8(); // v_just
            const rows = readU8();
            const cols = readU8();
            const rowPartBytes = Math.ceil((rows + 1) / 4);
            const colPartBytes = Math.ceil((cols + 1) / 4);
            pos = Math.min(limit, pos + rowPartBytes + colPartBytes);
            const cells = readRecordList(depth + 1);
            items.push({ kind: "matrix", rows, cols, cells });
            break;
          }

          case 6: { // EMBELL
            if (mtefVersion >= 5) opts = readU8();
            if (opts & 0x08) readNudge();
            const embellType = readU8();
            items.push({ kind: "embell", embellType });
            break;
          }

          case 7: // RULER
            skipRuler();
            break;

          case 8: // FONT_STYLE_DEF
            readMtefUInt();
            readU8();
            break;

          case 9: { // SIZE
            const lsize = readU8();
            if (lsize === 101) readI16();
            else if (lsize === 100) {
              readU8();
              readI16();
            } else {
              readU8();
            }
            break;
          }

          case 10: // FULL
          case 11: // SUB
          case 12: // SUB2
          case 13: // SYM
          case 14: // SUBSYM
            break;

          case 15: // COLOR
            readMtefUInt();
            break;

          case 16: { // COLOR_DEF
            const cOpts = readU8();
            const compCount = (cOpts & 0x01) ? 4 : 3; // CMYK vs RGB
            for (let c = 0; c < compCount; c++) readU16();
            if (cOpts & 0x04) readCString();
            break;
          }

          case 17: // FONT_DEF
            readMtefUInt();
            readCString();
            break;

          case 18: { // EQN_PREFS
            readU8(); // options
            const sizeCount = readU8();
            skipDimensionArray(sizeCount);
            const spaceCount = readU8();
            skipDimensionArray(spaceCount);
            const styleCount = readU8();
            for (let i = 0; i < styleCount; i++) {
              const fontDef = readMtefUInt();
              if (fontDef !== 0) readU8(); // char style
            }
            break;
          }

          case 19: // ENCODING_DEF
            readCString();
            break;

          default:
            // Bản ghi không xác định: dừng đọc danh sách hiện tại một cách an toàn
            return items;
        }
      }

      return items;
    }

    const topRecords = readRecordList(0);
    const latex = topRecords
      .map(slotToText)
      .filter(Boolean)
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();

    if (!latex) {
      throw new Error("Giải mã MTEF không thu được ký tự toán học nào.");
    }

    return { latex, isInline, mtefVersion };
  }

  function normalizePackagePath(partPath) {
    const clean = (partPath || "").trim().replace(/\\/g, "/");
    if (clean.startsWith("/")) return clean.toLowerCase();
    if (clean.startsWith("word/")) return ("/" + clean).toLowerCase();
    return ("/word/" + clean.replace(/^\.\//, "")).toLowerCase();
  }

  function getAttrByLocalName(el, localName) {
    if (!el || !el.attributes) return "";
    for (let i = 0; i < el.attributes.length; i++) {
      const attr = el.attributes[i];
      if (attr.localName === localName || attr.name === localName) {
        return attr.value || "";
      }
    }
    return "";
  }

  /**
   * Module độc lập: Chuyển đổi trực tiếp đối tượng MathType OLE (<w:object>)
   * sang chuỗi LaTeX $...$ ngay trong Word mà không cần cài MathType.
   */
  async function convertMathTypeToLatex(options) {
    if (!ensureWordAvailable() || state.isBusy) return null;

    const requestedScope = (options && options.scope) || state.scanScope || "selection";
    state.cancelRequested = false;
    setBusy(true, true);
    renderErrorReport([]);

    let totalFound = 0;
    let convertedCount = 0;
    let errorCount = 0;
    let autoSwitched = false;
    const errorItems = [];

    try {
      await Word.run(async (context) => {
        const target = await resolveTargetContainer(context, requestedScope);
        autoSwitched = target.autoSwitchedToBody;

        // 1. Lấy toàn bộ Flat OPC OOXML của vùng chọn hoặc toàn tài liệu
        const ooxmlRes = target.container.getOoxml();
        await context.sync();

        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(ooxmlRes.value || "", "application/xml");
        if (xmlDoc.querySelector("parsererror")) {
          throw new Error("Không thể phân tích cú pháp OOXML trả về từ Word.");
        }

        const wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const allParts = Array.from(xmlDoc.getElementsByTagName("*")).filter(
          (el) => el.localName === "part"
        );

        // 2. Xây dựng ánh xạ quan hệ (r:id -> /word/embeddings/oleObject*.bin) và dữ liệu Base64
        const relMap = new Map();
        const binaryPartMap = new Map();
        let documentPart = null;

        allParts.forEach((part) => {
          const partName = (getAttrByLocalName(part, "name") || "").toLowerCase();
          if (partName === "/word/document.xml") {
            documentPart = part;
          } else if (partName.endsWith("document.xml.rels")) {
            const relNodes = Array.from(part.getElementsByTagName("*")).filter(
              (el) => el.localName === "Relationship"
            );
            relNodes.forEach((rel) => {
              const id = getAttrByLocalName(rel, "Id");
              const targetPath = getAttrByLocalName(rel, "Target");
              if (id && targetPath) {
                relMap.set(id, normalizePackagePath(targetPath));
              }
            });
          }

          const binNode = Array.from(part.children).find((c) => c.localName === "binaryData");
          if (binNode && partName) {
            binaryPartMap.set(partName, binNode.textContent || "");
          }
        });

        const searchRoot = documentPart || xmlDoc;

        // 3. Tìm tất cả thẻ <w:object> chứa <o:OLEObject ProgID="Equation.DSMT4" ...>
        const wObjects = Array.from(searchRoot.getElementsByTagName("*")).filter(
          (el) => el.localName === "object"
        );

        const mathTypeTargets = [];
        wObjects.forEach((wObj) => {
          const oleNode = Array.from(wObj.getElementsByTagName("*")).find(
            (el) => el.localName === "OLEObject"
          );
          if (!oleNode) return;
          const progId = getAttrByLocalName(oleNode, "ProgID");
          if (!MATHTYPE_PROGID_RE.test(progId)) return;

          const relId = getAttrByLocalName(oleNode, "id");
          mathTypeTargets.push({ wObj, oleNode, progId, relId });
        });

        totalFound = mathTypeTargets.length;
        if (totalFound === 0) {
          updateProgressUI(0, 0, 0, 0);
          return;
        }

        updateProgressUI(
          0,
          totalFound,
          0,
          0,
          "Đang giải mã MathType 0/" + totalFound + "..."
        );
        await yieldToUi();

        // 4. Giải mã nhị phân từng đối tượng MathType OLE và thay thế <w:object> -> <w:r><w:t>$...$</w:t></w:r>
        for (let idx = 0; idx < mathTypeTargets.length; idx++) {
          if (state.cancelRequested) break;

          const { wObj, progId, relId } = mathTypeTargets[idx];
          try {
            if (!relId) {
              throw new Error("Thiếu thuộc tính r:id trên thẻ OLEObject.");
            }
            const binPath = relMap.get(relId);
            if (!binPath || !binaryPartMap.has(binPath)) {
              throw new Error("Không tìm thấy tệp nhị phân OLE cho " + relId + ".");
            }

            const oleBytes = base64ToUint8Array(binaryPartMap.get(binPath));
            const nativeStream = extractEquationNativeFromCfb(oleBytes);
            const decoded = parseMtefToLatex(nativeStream);

            const latexWrapped = decoded.isInline
              ? "$" + decoded.latex + "$"
              : "$$" + decoded.latex + "$$";

            const parent = wObj.parentNode;
            if (!parent) continue;

            // Tạo thẻ văn bản <w:r><w:t xml:space="preserve">$...$</w:t></w:r>
            const newRun = xmlDoc.createElementNS(wNs, "w:r");
            const newText = xmlDoc.createElementNS(wNs, "w:t");
            newText.setAttribute("xml:space", "preserve");
            newText.textContent = latexWrapped;
            newRun.appendChild(newText);

            if (parent.localName === "r" && parent.parentNode) {
              // Kiểm tra xem thẻ <w:r> bao ngoài có chỉ chứa <w:object> (và tùy chọn <w:rPr>) không
              const meaningfulSiblings = Array.from(parent.children).filter(
                (ch) => ch !== wObj && ch.localName !== "rPr"
              );
              if (meaningfulSiblings.length === 0) {
                parent.parentNode.replaceChild(newRun, parent);
              } else {
                parent.replaceChild(newText, wObj);
              }
            } else {
              parent.replaceChild(newRun, wObj);
            }

            convertedCount += 1;
          } catch (itemErr) {
            errorCount += 1;
            errorItems.push({
              paragraphIndex: idx + 1,
              raw: "[" + progId + " · " + (relId || "no-id") + "]",
              latex: "",
              displayMode: false,
              message: itemErr && itemErr.message ? itemErr.message : String(itemErr)
            });
          }

          updateProgressUI(
            idx + 1,
            totalFound,
            convertedCount,
            errorCount,
            "Đang giải mã MathType " + (idx + 1) + "/" + totalFound + "..."
          );
          await yieldToUi();
        }

        // 5. Ghi lại OOXML đã thay thế vào Word
        if (convertedCount > 0) {
          const serializer = new XMLSerializer();
          const updatedOoxml = serializer.serializeToString(xmlDoc);
          target.container.insertOoxml(updatedOoxml, Word.InsertLocation.replace);
          await context.sync();
        }
      });

      renderErrorReport(errorItems);
      const prefix = autoSwitched ? "Tự động áp dụng cho toàn tài liệu. " : "";

      if (totalFound === 0) {
        showStatus(
          prefix + "Không tìm thấy đối tượng MathType (Equation.DSMT4/Equation.3) nào.",
          "warn"
        );
      } else {
        updateProgressUI(
          totalFound,
          totalFound,
          convertedCount,
          errorCount,
          "Hoàn tất giải mã " + convertedCount + "/" + totalFound + " đối tượng MathType"
        );
        showStatus(
          prefix +
            "Đã chuyển " +
            convertedCount +
            "/" +
            totalFound +
            " công thức MathType sang LaTeX $...$." +
            (errorCount > 0 ? " (" + errorCount + " đối tượng lỗi được bỏ qua)." : ""),
          errorCount > 0 ? "warn" : "success"
        );
      }

      return { totalFound, convertedCount, errorCount, errorItems };
    } catch (err) {
      updateProgressUI(0, 0, 0, 0);
      showStatus(
        "Lỗi chuyển đổi MathType: " + (err && err.message ? err.message : String(err)),
        "error"
      );
      return null;
    } finally {
      setBusy(false, false);
      state.cancelRequested = false;
    }
  }

  // Xuất module độc lập lên window để có thể gọi trực tiếp từ bên ngoài nếu cần
  window.convertMathTypeToLatex = convertMathTypeToLatex;

  /* ------------------------------------------------------------------------
   * 9. MODULE ĐỘC LẬP: CHUYỂN ĐỔI HÀNG LOẠT LATEX ($...$, $$...$$) SANG
   *    MATHTYPE OLE (Equation.DSMT4) — MÃ HÓA NHỊ PHÂN MTEF v5 + CFBF + PNG
   * ------------------------------------------------------------------------ */
  const MT_FN_TEXT = 1;
  const MT_FN_FUNCTION = 2;
  const MT_FN_VARIABLE = 3;
  const MT_FN_LCGREEK = 4;
  const MT_FN_UCGREEK = 5;
  const MT_FN_SYMBOL = 6;
  const MT_FN_VECTOR = 7;
  const MT_FN_NUMBER = 8;

  // Bảng ánh xạ lệnh LaTeX sang ký tự Unicode & typeface chuẩn MTEF v5
  const LATEX_TO_MT_SYMBOL = {
    alpha: { ch: "α", code: 0x03b1, tf: MT_FN_LCGREEK },
    beta: { ch: "β", code: 0x03b2, tf: MT_FN_LCGREEK },
    gamma: { ch: "γ", code: 0x03b3, tf: MT_FN_LCGREEK },
    delta: { ch: "δ", code: 0x03b4, tf: MT_FN_LCGREEK },
    epsilon: { ch: "ε", code: 0x03b5, tf: MT_FN_LCGREEK },
    varepsilon: { ch: "ε", code: 0x03b5, tf: MT_FN_LCGREEK },
    zeta: { ch: "ζ", code: 0x03b6, tf: MT_FN_LCGREEK },
    eta: { ch: "η", code: 0x03b7, tf: MT_FN_LCGREEK },
    theta: { ch: "θ", code: 0x03b8, tf: MT_FN_LCGREEK },
    vartheta: { ch: "ϑ", code: 0x03d1, tf: MT_FN_LCGREEK },
    iota: { ch: "ι", code: 0x03b9, tf: MT_FN_LCGREEK },
    kappa: { ch: "κ", code: 0x03ba, tf: MT_FN_LCGREEK },
    lambda: { ch: "λ", code: 0x03bb, tf: MT_FN_LCGREEK },
    mu: { ch: "μ", code: 0x03bc, tf: MT_FN_LCGREEK },
    nu: { ch: "ν", code: 0x03bd, tf: MT_FN_LCGREEK },
    xi: { ch: "ξ", code: 0x03be, tf: MT_FN_LCGREEK },
    pi: { ch: "π", code: 0x03c0, tf: MT_FN_LCGREEK },
    rho: { ch: "ρ", code: 0x03c1, tf: MT_FN_LCGREEK },
    sigma: { ch: "σ", code: 0x03c3, tf: MT_FN_LCGREEK },
    tau: { ch: "τ", code: 0x03c4, tf: MT_FN_LCGREEK },
    upsilon: { ch: "υ", code: 0x03c5, tf: MT_FN_LCGREEK },
    phi: { ch: "φ", code: 0x03c6, tf: MT_FN_LCGREEK },
    varphi: { ch: "φ", code: 0x03c6, tf: MT_FN_LCGREEK },
    chi: { ch: "χ", code: 0x03c7, tf: MT_FN_LCGREEK },
    psi: { ch: "ψ", code: 0x03c8, tf: MT_FN_LCGREEK },
    omega: { ch: "ω", code: 0x03c9, tf: MT_FN_LCGREEK },
    Gamma: { ch: "Γ", code: 0x0393, tf: MT_FN_UCGREEK },
    Delta: { ch: "Δ", code: 0x0394, tf: MT_FN_UCGREEK },
    Theta: { ch: "Θ", code: 0x0398, tf: MT_FN_UCGREEK },
    Lambda: { ch: "Λ", code: 0x039b, tf: MT_FN_UCGREEK },
    Xi: { ch: "Ξ", code: 0x039e, tf: MT_FN_UCGREEK },
    Pi: { ch: "Π", code: 0x03a0, tf: MT_FN_UCGREEK },
    Sigma: { ch: "Σ", code: 0x03a3, tf: MT_FN_UCGREEK },
    Phi: { ch: "Φ", code: 0x03a6, tf: MT_FN_UCGREEK },
    Psi: { ch: "Ψ", code: 0x03a8, tf: MT_FN_UCGREEK },
    Omega: { ch: "Ω", code: 0x03a9, tf: MT_FN_UCGREEK },
    pm: { ch: "±", code: 0x00b1, tf: MT_FN_SYMBOL },
    mp: { ch: "∓", code: 0x2213, tf: MT_FN_SYMBOL },
    times: { ch: "×", code: 0x00d7, tf: MT_FN_SYMBOL },
    div: { ch: "÷", code: 0x00f7, tf: MT_FN_SYMBOL },
    cdot: { ch: "·", code: 0x22c5, tf: MT_FN_SYMBOL },
    ast: { ch: "∗", code: 0x2217, tf: MT_FN_SYMBOL },
    le: { ch: "≤", code: 0x2264, tf: MT_FN_SYMBOL },
    leq: { ch: "≤", code: 0x2264, tf: MT_FN_SYMBOL },
    ge: { ch: "≥", code: 0x2265, tf: MT_FN_SYMBOL },
    geq: { ch: "≥", code: 0x2265, tf: MT_FN_SYMBOL },
    ne: { ch: "≠", code: 0x2260, tf: MT_FN_SYMBOL },
    neq: { ch: "≠", code: 0x2260, tf: MT_FN_SYMBOL },
    approx: { ch: "≈", code: 0x2248, tf: MT_FN_SYMBOL },
    equiv: { ch: "≡", code: 0x2261, tf: MT_FN_SYMBOL },
    sim: { ch: "∼", code: 0x223c, tf: MT_FN_SYMBOL },
    infty: { ch: "∞", code: 0x221e, tf: MT_FN_SYMBOL },
    to: { ch: "→", code: 0x2192, tf: MT_FN_SYMBOL },
    rightarrow: { ch: "→", code: 0x2192, tf: MT_FN_SYMBOL },
    leftarrow: { ch: "←", code: 0x2190, tf: MT_FN_SYMBOL },
    Rightarrow: { ch: "⇒", code: 0x21d2, tf: MT_FN_SYMBOL },
    Leftrightarrow: { ch: "⇔", code: 0x21d4, tf: MT_FN_SYMBOL },
    in: { ch: "∈", code: 0x2208, tf: MT_FN_SYMBOL },
    notin: { ch: "∉", code: 0x2209, tf: MT_FN_SYMBOL },
    subset: { ch: "⊂", code: 0x2282, tf: MT_FN_SYMBOL },
    subseteq: { ch: "⊆", code: 0x2286, tf: MT_FN_SYMBOL },
    cup: { ch: "∪", code: 0x222a, tf: MT_FN_SYMBOL },
    cap: { ch: "∩", code: 0x2229, tf: MT_FN_SYMBOL },
    emptyset: { ch: "∅", code: 0x2205, tf: MT_FN_SYMBOL },
    forall: { ch: "∀", code: 0x2200, tf: MT_FN_SYMBOL },
    exists: { ch: "∃", code: 0x2203, tf: MT_FN_SYMBOL },
    partial: { ch: "∂", code: 0x2202, tf: MT_FN_SYMBOL },
    nabla: { ch: "∇", code: 0x2207, tf: MT_FN_SYMBOL },
    angle: { ch: "∠", code: 0x2220, tf: MT_FN_SYMBOL },
    perp: { ch: "⊥", code: 0x22a5, tf: MT_FN_SYMBOL },
    parallel: { ch: "∥", code: 0x2225, tf: MT_FN_SYMBOL },
    triangle: { ch: "△", code: 0x25b3, tf: MT_FN_SYMBOL },
    circ: { ch: "°", code: 0x00b0, tf: MT_FN_SYMBOL },
    dots: { ch: "…", code: 0x2026, tf: MT_FN_SYMBOL },
    ldots: { ch: "…", code: 0x2026, tf: MT_FN_SYMBOL },
    cdots: { ch: "⋯", code: 0x22ef, tf: MT_FN_SYMBOL },
    prime: { ch: "′", code: 0x2032, tf: MT_FN_SYMBOL },
    "{": { ch: "{", code: 0x007b, tf: MT_FN_SYMBOL },
    "}": { ch: "}", code: 0x007d, tf: MT_FN_SYMBOL },
    "|": { ch: "‖", code: 0x2016, tf: MT_FN_SYMBOL },
    "%": { ch: "%", code: 0x0025, tf: MT_FN_TEXT }
  };

  const LATEX_KNOWN_FUNCS = new Set([
    "sin", "cos", "tan", "cot", "sec", "csc",
    "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh",
    "log", "ln", "max", "min", "exp", "det", "deg",
    "gcd", "sup", "inf", "arg", "dim", "hom", "ker", "Pr"
  ]);

  const BLACKBOARD_MAP = {
    R: { ch: "ℝ", code: 0x211d },
    Z: { ch: "ℤ", code: 0x2124 },
    N: { ch: "ℕ", code: 0x2115 },
    Q: { ch: "ℚ", code: 0x211a },
    C: { ch: "ℂ", code: 0x2102 }
  };

  /**
   * Bộ ghi nhị phân tự động mở rộng dùng Uint8Array & DataView thuần.
   */
  function createBinaryWriter(initialCap) {
    let buf = new Uint8Array(initialCap || 256);
    let view = new DataView(buf.buffer);
    let len = 0;

    function ensure(extra) {
      if (len + extra <= buf.length) return;
      let nextSize = buf.length * 2;
      while (nextSize < len + extra) nextSize *= 2;
      const grown = new Uint8Array(nextSize);
      grown.set(buf.subarray(0, len));
      buf = grown;
      view = new DataView(buf.buffer);
    }

    return {
      u8(v) {
        ensure(1);
        buf[len++] = v & 0xff;
      },
      u16(v) {
        ensure(2);
        view.setUint16(len, v & 0xffff, true);
        len += 2;
      },
      u32(v) {
        ensure(4);
        view.setUint32(len, v >>> 0, true);
        len += 4;
      },
      bytes(arr) {
        ensure(arr.length);
        buf.set(arr, len);
        len += arr.length;
      },
      cstr(str) {
        ensure(str.length + 1);
        for (let i = 0; i < str.length; i++) {
          buf[len++] = str.charCodeAt(i) & 0xff;
        }
        buf[len++] = 0;
      },
      toUint8Array() {
        return buf.slice(0, len);
      }
    };
  }

  function uint8ArrayToBase64(bytes) {
    const chunkSize = 0x8000;
    let binary = "";
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const slice = bytes.subarray(i, Math.min(i + chunkSize, bytes.length));
      binary += String.fromCharCode.apply(null, Array.from(slice));
    }
    return window.btoa(binary);
  }

  /**
   * Bộ phân tích cú pháp LaTeX -> Cây biểu thức toán học (AST) dùng chung
   * cho cả bộ mã hóa nhị phân MTEF v5 và bộ vẽ ảnh xem trước PNG trên Canvas.
   */
  function parseLatexToAst(rawLatex) {
    const src = stripOuterDelimiters(rawLatex || "");
    if (!src) {
      throw new Error("Công thức LaTeX trống.");
    }

    // Kiểm tra nhanh ngoặc nhọn cân bằng
    let braceDepth = 0;
    for (let i = 0; i < src.length; i++) {
      if (src[i] === "\\" && i + 1 < src.length) {
        i += 1;
        continue;
      }
      if (src[i] === "{") braceDepth += 1;
      else if (src[i] === "}") {
        braceDepth -= 1;
        if (braceDepth < 0) throw new Error("Thừa dấu ngoặc nhọn đóng '}'.");
      }
    }
    if (braceDepth !== 0) {
      throw new Error("Thiếu dấu ngoặc nhọn đóng '}' trong biểu thức LaTeX.");
    }

    let pos = 0;

    function skipSpaces() {
      while (pos < src.length && /\s/.test(src[pos])) pos += 1;
    }

    function readCommandName() {
      if (pos >= src.length || src[pos] !== "\\") return "";
      pos += 1;
      if (pos >= src.length) return "";
      if (!/[a-zA-Z]/.test(src[pos])) {
        return src[pos++];
      }
      const start = pos;
      while (pos < src.length && /[a-zA-Z]/.test(src[pos])) pos += 1;
      return src.slice(start, pos);
    }

    function readBracedRaw() {
      skipSpaces();
      if (pos >= src.length) {
        throw new Error("Thiếu tham số cho lệnh LaTeX.");
      }
      if (src[pos] !== "{") {
        if (src[pos] === "\\") return "\\" + readCommandName();
        return src[pos++];
      }
      pos += 1;
      let depth = 1;
      const start = pos;
      while (pos < src.length && depth > 0) {
        if (src[pos] === "\\" && pos + 1 < src.length) {
          pos += 2;
          continue;
        }
        if (src[pos] === "{") depth += 1;
        else if (src[pos] === "}") depth -= 1;
        if (depth > 0) pos += 1;
      }
      const inside = src.slice(start, pos);
      if (pos < src.length && src[pos] === "}") pos += 1;
      return inside;
    }

    function readOptionalBracketRaw() {
      skipSpaces();
      if (pos >= src.length || src[pos] !== "[") return null;
      pos += 1;
      let depth = 1;
      const start = pos;
      while (pos < src.length && depth > 0) {
        if (src[pos] === "[") depth += 1;
        else if (src[pos] === "]") depth -= 1;
        if (depth > 0) pos += 1;
      }
      const inside = src.slice(start, pos);
      if (pos < src.length && src[pos] === "]") pos += 1;
      return inside;
    }

    function readGroupNodes() {
      skipSpaces();
      if (pos >= src.length) {
        throw new Error("Thiếu nhóm biểu thức LaTeX.");
      }
      if (src[pos] === "{") {
        pos += 1;
        const nodes = parseSequence("}");
        if (pos >= src.length || src[pos] !== "}") {
          throw new Error("Thiếu dấu '}' đóng nhóm.");
        }
        pos += 1;
        return nodes;
      }
      const atom = parseAtom();
      return atom ? [atom] : [];
    }

    function readDelimiterChar() {
      skipSpaces();
      if (pos >= src.length) return ".";
      if (src[pos] === "\\") {
        const cmd = readCommandName();
        if (cmd === "{" || cmd === "}") return cmd;
        if (cmd === "|") return "‖";
        if (cmd === "langle") return "⟨";
        if (cmd === "rangle") return "⟩";
        if (cmd === "lfloor") return "⌊";
        if (cmd === "rfloor") return "⌋";
        if (cmd === "lceil") return "⌈";
        if (cmd === "rceil") return "⌉";
        return ".";
      }
      return src[pos++];
    }

    function parseEnvironment(envName) {
      const endToken = "\\end{" + envName + "}";
      const endIdx = src.indexOf(endToken, pos);
      if (endIdx === -1) {
        throw new Error("Thiếu " + endToken + ".");
      }
      const bodyRaw = src.slice(pos, endIdx);
      pos = endIdx + endToken.length;

      // Tách hàng theo \\ và cột theo & ở cấp ngoặc 0
      const rowStrings = [];
      let cur = "";
      let depth = 0;
      for (let i = 0; i < bodyRaw.length; i++) {
        if (bodyRaw[i] === "{") depth += 1;
        else if (bodyRaw[i] === "}") depth = Math.max(0, depth - 1);
        if (depth === 0 && bodyRaw[i] === "\\" && bodyRaw[i + 1] === "\\") {
          rowStrings.push(cur.trim());
          cur = "";
          i += 1;
          continue;
        }
        cur += bodyRaw[i];
      }
      if (cur.trim()) rowStrings.push(cur.trim());

      const rows = rowStrings.map((rowStr) => {
        const cellStrings = [];
        let cell = "";
        let d = 0;
        for (let i = 0; i < rowStr.length; i++) {
          if (rowStr[i] === "\\" && i + 1 < rowStr.length) {
            cell += rowStr[i] + rowStr[i + 1];
            i += 1;
            continue;
          }
          if (rowStr[i] === "{") d += 1;
          else if (rowStr[i] === "}") d = Math.max(0, d - 1);
          if (d === 0 && rowStr[i] === "&") {
            cellStrings.push(cell.trim());
            cell = "";
            continue;
          }
          cell += rowStr[i];
        }
        cellStrings.push(cell.trim());
        return cellStrings.map((c) => (c ? parseLatexToAst(c) : []));
      });

      return { kind: "env", env: envName, rows };
    }

    function parseSubSupSuffix() {
      let sub = null;
      let sup = null;
      while (true) {
        skipSpaces();
        if (pos < src.length && src[pos] === "_" && !sub) {
          pos += 1;
          sub = readGroupNodes();
        } else if (pos < src.length && src[pos] === "^" && !sup) {
          pos += 1;
          sup = readGroupNodes();
        } else {
          break;
        }
      }
      return { sub, sup };
    }

    function parseAtom() {
      skipSpaces();
      if (pos >= src.length) return null;
      const ch = src[pos];

      if (ch === "{") {
        pos += 1;
        const inner = parseSequence("}");
        if (pos < src.length && src[pos] === "}") pos += 1;
        return { kind: "group", items: inner };
      }

      if (ch === "(") {
        // Kiểm tra xem có dấu ')' đóng tương ứng ở cùng cấp ngoặc không
        let scan = pos + 1;
        let pDepth = 1;
        let bDepth = 0;
        let matchedClose = -1;
        while (scan < src.length) {
          const c = src[scan];
          if (c === "\\" && scan + 1 < src.length) {
            scan += 2;
            continue;
          }
          if (c === "{") bDepth += 1;
          else if (c === "}") {
            if (bDepth === 0) break;
            bDepth -= 1;
          } else if (bDepth === 0) {
            if (c === "(") pDepth += 1;
            else if (c === ")") {
              pDepth -= 1;
              if (pDepth === 0) {
                matchedClose = scan;
                break;
              }
            }
          }
          scan += 1;
        }
        if (matchedClose !== -1) {
          pos += 1;
          const body = parseSequence(")");
          if (pos < src.length && src[pos] === ")") pos += 1;
          return { kind: "fence", left: "(", right: ")", body };
        }
        pos += 1;
        return { kind: "char", ch: "(", code: 0x0028, tf: MT_FN_SYMBOL };
      }

      if (ch === "\\") {
        const cmd = readCommandName();
        if (!cmd) return null;

        // Các lệnh khoảng cách và định dạng hiển thị
        if (cmd === "," || cmd === ";" || cmd === ":" || cmd === " " || cmd === "quad" || cmd === "qquad") {
          return { kind: "space", ch: " ", code: 0x0020 };
        }
        if (cmd === "!" || cmd === "displaystyle" || cmd === "textstyle" || cmd === "limits" || cmd === "nolimits") {
          return null;
        }

        if (cmd === "frac" || cmd === "dfrac" || cmd === "tfrac") {
          const num = readGroupNodes();
          const den = readGroupNodes();
          return { kind: "frac", num, den };
        }

        if (cmd === "sqrt") {
          const degRaw = readOptionalBracketRaw();
          const body = readGroupNodes();
          const deg = degRaw ? parseLatexToAst(degRaw) : null;
          return { kind: "root", deg, body };
        }

        if (cmd === "left") {
          const leftDelim = readDelimiterChar();
          const body = parseSequence("\\right");
          if (src.slice(pos, pos + 6) === "\\right") {
            pos += 6;
            const rightDelim = readDelimiterChar();
            return { kind: "fence", left: leftDelim, right: rightDelim, body };
          }
          return { kind: "fence", left: leftDelim, right: ".", body };
        }

        if (cmd === "begin") {
          const envName = readBracedRaw().trim();
          return parseEnvironment(envName);
        }

        if (
          cmd === "int" ||
          cmd === "iint" ||
          cmd === "iiint" ||
          cmd === "oint" ||
          cmd === "sum" ||
          cmd === "prod" ||
          cmd === "coprod" ||
          cmd === "bigcup" ||
          cmd === "bigcap"
        ) {
          const limits = parseSubSupSuffix();
          return { kind: "nary", op: cmd, sub: limits.sub, sup: limits.sup };
        }

        if (cmd === "lim") {
          const limits = parseSubSupSuffix();
          return { kind: "lim", sub: limits.sub, sup: limits.sup };
        }

        if (
          cmd === "vec" ||
          cmd === "overrightarrow" ||
          cmd === "overline" ||
          cmd === "bar" ||
          cmd === "underline" ||
          cmd === "hat" ||
          cmd === "widehat" ||
          cmd === "tilde" ||
          cmd === "widetilde" ||
          cmd === "dot" ||
          cmd === "ddot"
        ) {
          const body = readGroupNodes();
          return { kind: "accent", cmd, body };
        }

        if (cmd === "mathbb") {
          const raw = readBracedRaw().trim();
          const mapped = BLACKBOARD_MAP[raw];
          if (mapped) {
            return { kind: "char", ch: mapped.ch, code: mapped.code, tf: MT_FN_SYMBOL };
          }
          return { kind: "text", text: raw, tf: MT_FN_VECTOR };
        }

        if (cmd === "text" || cmd === "mathrm" || cmd === "operatorname" || cmd === "mathbf") {
          const raw = readBracedRaw();
          const tf =
            cmd === "operatorname"
              ? MT_FN_FUNCTION
              : cmd === "mathbf"
                ? MT_FN_VECTOR
                : MT_FN_TEXT;
          return { kind: "text", text: raw, tf };
        }

        if (LATEX_KNOWN_FUNCS.has(cmd)) {
          return { kind: "func", name: cmd };
        }

        if (LATEX_TO_MT_SYMBOL[cmd]) {
          const sym = LATEX_TO_MT_SYMBOL[cmd];
          return { kind: "char", ch: sym.ch, code: sym.code, tf: sym.tf };
        }

        // Fallback cho lệnh chữ chưa có trong bảng
        return { kind: "text", text: cmd, tf: MT_FN_FUNCTION };
      }

      // Ký tự thông thường
      pos += 1;
      if (/[a-zA-Z]/.test(ch)) {
        return { kind: "char", ch, code: ch.charCodeAt(0), tf: MT_FN_VARIABLE };
      }
      if (/[0-9]/.test(ch)) {
        return { kind: "char", ch, code: ch.charCodeAt(0), tf: MT_FN_NUMBER };
      }
      if (ch === "-") {
        return { kind: "char", ch: "−", code: 0x2212, tf: MT_FN_SYMBOL };
      }
      if (ch === "'") {
        return { kind: "char", ch: "′", code: 0x2032, tf: MT_FN_SYMBOL };
      }
      if (/[+=<>|[\]/!*,;:.]/.test(ch)) {
        return { kind: "char", ch, code: ch.charCodeAt(0), tf: MT_FN_SYMBOL };
      }
      return { kind: "char", ch, code: ch.charCodeAt(0), tf: MT_FN_TEXT };
    }

    function parseSequence(stopToken) {
      const list = [];
      while (pos < src.length) {
        skipSpaces();
        if (pos >= src.length) break;

        if (stopToken === "}" && src[pos] === "}") break;
        if (stopToken === ")" && src[pos] === ")") break;
        if (stopToken === "\\right" && src.slice(pos, pos + 6) === "\\right") break;

        // Xử lý chỉ số trên/dưới ngay sau phần tử trước đó
        if (src[pos] === "_" || src[pos] === "^") {
          const scripts = parseSubSupSuffix();
          if (scripts.sub || scripts.sup) {
            list.push({ kind: "script", sub: scripts.sub, sup: scripts.sup });
          }
          continue;
        }

        const atom = parseAtom();
        if (atom) {
          list.push(atom);
          // Nếu atom không phải nary/lim đã tự đọc cận, kiểm tra _ và ^ liền sau
          if (atom.kind !== "nary" && atom.kind !== "lim") {
            const scripts = parseSubSupSuffix();
            if (scripts.sub || scripts.sup) {
              list.push({ kind: "script", sub: scripts.sub, sup: scripts.sup });
            }
          }
        }
      }
      return list;
    }

    return parseSequence(null);
  }

  /**
   * Mã hóa một mảng kích thước (Dimension Array) chuẩn MTEF v5 EQN_PREFS:
   * Toàn bộ các kích thước được nén liên tục thành 1 luồng nibble (4-bit, high nibble trước),
   * mỗi kích thước kết thúc bằng nibble 0xF, và CHỈ đệm thêm 1 nibble 0xF ở cuối mảng nếu tổng số nibble là lẻ.
   * unitCode: 2 = pt, 4 = %
   */
  function encodeMtefDimensionArray(dims) {
    const nibbles = [];
    for (let d = 0; d < dims.length; d++) {
      const unitCode = dims[d][0];
      const s = String(dims[d][1]);
      nibbles.push(unitCode & 0x0f);
      for (let i = 0; i < s.length; i++) {
        const c = s[i];
        if (c >= "0" && c <= "9") nibbles.push(c.charCodeAt(0) - 48);
        else if (c === ".") nibbles.push(0x0a);
        else if (c === "-") nibbles.push(0x0b);
      }
      nibbles.push(0x0f);
    }
    if (nibbles.length % 2 === 1) nibbles.push(0x0f);

    const out = [];
    for (let i = 0; i < nibbles.length; i += 2) {
      out.push(((nibbles[i] & 0x0f) << 4) | (nibbles[i + 1] & 0x0f));
    }
    return out;
  }

  /**
   * Bộ mã hóa nhị phân MTEF v5 (MathType Equation Format 5) từ cây AST.
   */
  function encodeAstToMtefV5(astNodes, options) {
    const displayMode = Boolean(options && options.displayMode);
    const fontSizePt = Number((options && options.fontSizePt) || 12);
    const w = createBinaryWriter(256);

    // 1. MTEF v5 Header chuẩn MathType 7.0
    w.u8(5); // Byte 0: MTEF version 5
    w.u8(1); // Byte 1: Platform = Windows (1)
    w.u8(0); // Byte 2: Product = MathType (0)
    w.u8(7); // Byte 3: Product version = MathType 7 (7)
    w.u8(0); // Byte 4: Product subversion = 0
    w.cstr("DSMT4"); // Application key chuẩn Equation.DSMT4
    w.u8(displayMode ? 0 : 1); // 0 = display equation, 1 = inline equation

    // 2. FONT_DEF (record 17): Times New Roman (encDef=1 MTCode) & Symbol (encDef=3 Symbol)
    w.u8(17); // FONT_DEF #1
    w.u8(1);  // encDefIndex = 1 (MTCode)
    w.cstr("Times New Roman");

    w.u8(17); // FONT_DEF #2
    w.u8(3);  // encDefIndex = 3 (Symbol — chuẩn MTEF v5)
    w.cstr("Symbol");

    // 3. EQN_PREFS (record 18): thiết lập cỡ chữ mặc định (vd: 12pt) và bảng kiểu chữ Times New Roman
    w.u8(18); // EQN_PREFS
    w.u8(0);  // options
    w.u8(7);  // 7 mức cỡ chữ chuẩn MathType
    w.bytes(
      encodeMtefDimensionArray([
        [2, String(fontSizePt)], // Full (12pt)
        [4, "58"],               // Subscript (58%)
        [4, "42"],               // Sub-subscript (42%)
        [4, "150"],              // Symbol (150%)
        [4, "100"],              // Sub-symbol (100%)
        [4, "75"],               // User 1 (75%)
        [4, "150"]               // User 2 (150%)
      ])
    );

    w.u8(0); // spaces count (dùng mặc định của MathType)

    // 11 styles chuẩn: [fontDefIndex, charStyle (0=plain, 1=bold, 2=italic)]
    const styles = [
      [1, 0], // 1: Text (Times New Roman regular)
      [1, 0], // 2: Function (Times New Roman regular)
      [1, 2], // 3: Variable (Times New Roman italic)
      [2, 2], // 4: L.C. Greek (Symbol italic)
      [2, 0], // 5: U.C. Greek (Symbol regular)
      [2, 0], // 6: Symbol (Symbol regular)
      [1, 1], // 7: Vector/Matrix (Times New Roman bold)
      [1, 0], // 8: Number (Times New Roman regular)
      [1, 0], // 9: User 1
      [1, 0], // 10: User 2
      [2, 0]  // 11: MT Extra / Symbol
    ];
    w.u8(styles.length);
    styles.forEach(([fontDef, charStyle]) => {
      w.u8(fontDef);
      w.u8(charStyle);
    });

    // 4. Đặt kích thước hiện tại về FULL (record 10)
    w.u8(10);

    function writeChar(code, typeface) {
      w.u8(2); // CHAR record
      w.u8(0); // options: có MTCode 16-bit, không embellishment
      w.u8((typeface + 128) & 0xff); // MTEF v5 typeface = style + 128
      w.u16(code & 0xffff);
    }

    function writeNullLine() {
      w.u8(1);    // LINE record
      w.u8(0x01); // L_NULL (không có danh sách con và không có END)
    }

    function writeLine(nodes) {
      if (!nodes || nodes.length === 0) {
        writeNullLine();
        return;
      }
      w.u8(1); // LINE record
      w.u8(0); // options
      writeNodeList(nodes);
      w.u8(0); // END LINE
    }

    function writeTemplate(selector, variation, slotWriters) {
      w.u8(3); // TMPL record
      w.u8(0); // options
      w.u8(selector & 0xff);
      if (variation > 0x7f) {
        w.u8((variation & 0x7f) | 0x80);
        w.u8((variation >> 8) & 0xff);
      } else {
        w.u8(variation & 0x7f);
      }
      w.u8(0); // template-specific options
      for (let i = 0; i < slotWriters.length; i++) {
        slotWriters[i]();
      }
      w.u8(0); // END TMPL
    }

    function writeTextString(str, tf) {
      for (let i = 0; i < str.length; i++) {
        writeChar(str.charCodeAt(i), tf);
      }
    }

    function writeNode(node) {
      if (!node) return;
      switch (node.kind) {
        case "char":
          writeChar(node.code, node.tf || MT_FN_VARIABLE);
          break;

        case "space":
          writeChar(0x0020, MT_FN_TEXT);
          break;

        case "text":
          writeTextString(node.text || "", node.tf || MT_FN_TEXT);
          break;

        case "func":
          writeTextString(node.name || "", MT_FN_FUNCTION);
          break;

        case "group":
          writeNodeList(node.items || []);
          break;

        case "frac":
          // tmFRACT (12): slot 0 = tử số, slot 1 = mẫu số
          writeTemplate(12, 0, [
            () => writeLine(node.num),
            () => writeLine(node.den)
          ]);
          break;

        case "root":
          // tmROOT (11): variation 0 = căn bậc hai, 1 = căn bậc n
          if (node.deg && node.deg.length > 0) {
            writeTemplate(11, 1, [
              () => writeLine(node.body),
              () => writeLine(node.deg)
            ]);
          } else {
            writeTemplate(11, 0, [
              () => writeLine(node.body)
            ]);
          }
          break;

        case "script":
          if (node.sub && node.sup) {
            // tmSUBSUP (29) — ScrBoxClass: slot 0 = sub, slot 1 = sup
            writeTemplate(29, 0, [
              () => writeLine(node.sub),
              () => writeLine(node.sup)
            ]);
          } else if (node.sub) {
            // tmSUB (27) — ScrBoxClass: slot 0 = sub, slot 1 = L_NULL
            writeTemplate(27, 0, [
              () => writeLine(node.sub),
              () => writeNullLine()
            ]);
          } else if (node.sup) {
            // tmSUP (28) — ScrBoxClass: slot 0 = L_NULL (sub trống), slot 1 = sup
            writeTemplate(28, 0, [
              () => writeNullLine(),
              () => writeLine(node.sup)
            ]);
          }
          break;

        case "fence": {
          const left = node.left || "(";
          const right = node.right || ")";
          let selector = 1; // tmPAREN
          let variation = 0x03; // cả trái & phải
          if (left === "(" || right === ")") selector = 1;
          else if (left === "[" || right === "]") selector = 3;
          else if (left === "{" || right === "}") {
            selector = 2;
            variation = (left === "{" ? 0x01 : 0) | (right === "}" ? 0x02 : 0);
          } else if (left === "|" || right === "|") selector = 4;
          else if (left === "‖" || right === "‖") selector = 5;
          else if (left === "⟨" || right === "⟩") selector = 0;
          else if (left === "⌊" || right === "⌋") selector = 6;
          else if (left === "⌈" || right === "⌉") selector = 7;

          writeTemplate(selector, variation, [() => writeLine(node.body)]);
          break;
        }

        case "nary": {
          let selector = 16; // tmINTEG
          let variation = 0x01; // có cận trên/dưới
          if (node.op === "iint") variation = 0x04;
          else if (node.op === "iiint") variation = 0x08;
          else if (node.op === "oint") variation = 0x10;
          else if (node.op === "sum") selector = 17;
          else if (node.op === "prod") selector = 18;
          else if (node.op === "coprod") selector = 19;
          else if (node.op === "bigcup") selector = 20;
          else if (node.op === "bigcap") selector = 21;

          // Slot 0 = biểu thức trong tích phân/tổng (L_NULL để nối tiếp tự nhiên), Slot 1 = cận dưới, Slot 2 = cận trên
          writeTemplate(selector, variation, [
            () => writeNullLine(),
            () => (node.sub ? writeLine(node.sub) : writeNullLine()),
            () => (node.sup ? writeLine(node.sup) : writeNullLine())
          ]);
          break;
        }

        case "lim":
          // tmLIM (24): slot 0 = "lim", slot 1 = cận dưới, slot 2 = cận trên
          writeTemplate(24, 0, [
            () => {
              w.u8(1);
              w.u8(0);
              writeTextString("lim", MT_FN_FUNCTION);
              w.u8(0);
            },
            () => (node.sub ? writeLine(node.sub) : writeNullLine()),
            () => (node.sup ? writeLine(node.sup) : writeNullLine())
          ]);
          break;

        case "accent": {
          let selector = 31; // tmVEC
          if (node.cmd === "overline" || node.cmd === "bar") selector = 14; // tmOBAR
          else if (node.cmd === "underline") selector = 13; // tmUBAR
          else if (node.cmd === "hat" || node.cmd === "widehat") selector = 32; // tmHAT
          else if (node.cmd === "vec" || node.cmd === "overrightarrow") selector = 31; // tmVEC
          else selector = 32;
          writeTemplate(selector, 0, [() => writeLine(node.body)]);
          break;
        }

        case "env": {
          const rows = node.rows || [];
          const rowCount = rows.length || 1;
          const colCount = Math.max(1, ...rows.map((r) => r.length));

          function writeMatrixRecord() {
            w.u8(5); // MATRIX record
            w.u8(0); // options
            w.u8(1); // valign: center
            w.u8(1); // h_just: center
            w.u8(1); // v_just: center
            w.u8(rowCount & 0xff);
            w.u8(colCount & 0xff);
            const rowPartBytes = Math.ceil((rowCount + 1) / 4);
            const colPartBytes = Math.ceil((colCount + 1) / 4);
            for (let i = 0; i < rowPartBytes + colPartBytes; i++) w.u8(0);
            for (let r = 0; r < rowCount; r++) {
              for (let c = 0; c < colCount; c++) {
                const cellNodes = (rows[r] && rows[r][c]) || [];
                writeLine(cellNodes);
              }
            }
            w.u8(0); // END MATRIX
          }

          if (node.env === "cases") {
            // tmBRACE (2) chỉ ngoặc nhọn trái (variation = 0x01) bọc PILE
            writeTemplate(2, 0x01, [
              () => {
                w.u8(1); // LINE bọc PILE
                w.u8(0);
                w.u8(4); // PILE
                w.u8(0); // options
                w.u8(1); // halign: left
                w.u8(1); // valign: center
                for (let r = 0; r < rows.length; r++) {
                  const flattenedRow = [];
                  rows[r].forEach((cell, cIdx) => {
                    if (cIdx > 0) flattenedRow.push({ kind: "space", ch: " ", code: 0x0020 });
                    cell.forEach((n) => flattenedRow.push(n));
                  });
                  writeLine(flattenedRow);
                }
                w.u8(0); // END PILE
                w.u8(0); // END LINE
              }
            ]);
          } else if (node.env === "pmatrix") {
            writeTemplate(1, 0x03, [
              () => {
                w.u8(1);
                w.u8(0);
                writeMatrixRecord();
                w.u8(0);
              }
            ]);
          } else if (node.env === "bmatrix") {
            writeTemplate(3, 0x03, [
              () => {
                w.u8(1);
                w.u8(0);
                writeMatrixRecord();
                w.u8(0);
              }
            ]);
          } else {
            writeMatrixRecord();
          }
          break;
        }

        default:
          break;
      }
    }

    function writeNodeList(nodes) {
      for (let i = 0; i < nodes.length; i++) {
        writeNode(nodes[i]);
      }
    }

    // 5. Dòng công thức chính (top-level LINE) + END toàn gói MTEF
    writeLine(astNodes);
    w.u8(0); // END MTEF

    return w.toUint8Array();
  }

  /**
   * CLSID chuẩn của Design Science / Wiris MathType 7.0 (Equation.DSMT4) theo thứ tự Little-Endian:
   * {0002CE02-0000-0000-C000-000000000046}
   */
  const MATHTYPE7_CLSID = new Uint8Array([
    0x02, 0xce, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
    0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46
  ]);

  /**
   * Đóng gói luồng byte MTEF v5 vào cấu trúc chuẩn luồng "Equation Native"
   * (Header 28-byte EQNOLEFILEHDR chuẩn Design Science MathType:
   *  cbHdr = 28, version = 0x00020000, cf = 0x0008, cbObject = độ dài MTEF thực tế, nReserved[4]).
   * Tự động đệm thêm byte 0x00 ở cuối để tổng dung lượng luồng đạt tối thiểu 4096 bytes,
   * đảm bảo luồng nằm hoàn toàn trong các FAT Sector 512-byte chuẩn (không dùng MiniFAT).
   */
  function wrapMtefInEquationNative(mtefBytes) {
    const rawLen = 28 + mtefBytes.length;
    const paddedLen = Math.max(4096, Math.ceil(rawLen / 512) * 512);
    const out = new Uint8Array(paddedLen); // Mặc định khởi tạo toàn bộ là 0x00
    const dv = new DataView(out.buffer);
    dv.setUint16(0, 28, true);               // cbHdr = 28 bytes
    dv.setUint32(2, 0x00020000, true);       // version = 2.0 (0x00020000)
    dv.setUint16(6, 0x0008, true);           // cf = 0x0008 (chuẩn Design Science MathType)
    dv.setUint32(8, mtefBytes.length, true); // cbObject = độ dài thực tế của mảng MTEF v5
    dv.setUint32(12, 0, true);               // nReserved1
    dv.setUint32(16, 0, true);               // nReserved2
    dv.setUint32(20, 0, true);               // nReserved3
    dv.setUint32(24, 0, true);               // nReserved4
    out.set(mtefBytes, 28);
    return out;
  }

  /**
   * Tạo luồng "\x01CompObj" chuẩn OLE2 cho MathType 7.0:
   * - CLSID: {0002CE02-0000-0000-C000-000000000046}
   * - UserType: "MathType 7.0 Equation" (độ dài kèm ký tự null \0 = 22 bytes)
   * - ProgID: "Equation.DSMT4" (độ dài kèm ký tự null \0 = 15 bytes)
   */
  function buildCompObjStream() {
    const w = createBinaryWriter(128);
    const userType = "MathType 7.0 Equation";
    const progId = "Equation.DSMT4";

    w.u16(0x0001); // Version
    w.u16(0xfffe); // Byte order (little-endian)
    w.u32(0x00000a03); // OS version (Win32)
    w.u32(0xffffffff); // Reserved

    // 16-byte CLSID chuẩn Design Science / Wiris MathType 7.0
    w.bytes(MATHTYPE7_CLSID);

    w.u32(userType.length + 1); // Độ dài UserType gồm cả ký tự kết thúc null '\0'
    w.cstr(userType);

    w.u32(0); // Clipboard format = none

    w.u32(progId.length + 1); // Độ dài ProgID gồm cả ký tự kết thúc null '\0'
    w.cstr(progId);

    w.u32(0x71b239f4); // Unicode marker
    w.u32(0); // UnicodeUserType length = 0
    w.u32(0); // UnicodeClipboardFormat length = 0
    w.u32(0); // Reserved2 length = 0
    return w.toUint8Array();
  }

  /**
   * Tạo luồng "\x01Ole" chuẩn (20 bytes) cho đối tượng nhúng OLE2.
   */
  function buildOleStream() {
    const out = new Uint8Array(20);
    const dv = new DataView(out.buffer);
    dv.setUint32(0, 0x02000001, true); // OLE version 2.0
    dv.setUint32(4, 0x00000000, true); // Embedded object flags
    return out;
  }

  /**
   * Đóng gói luồng "Equation Native" (cùng "\x01CompObj" và "\x01Ole") vào tệp nhị phân
   * chuẩn OLE Compound File Binary Format (CFBF v3 — 512-byte sectors) bằng Uint8Array & DataView.
   */
  function buildOleCfbWithEquationNative(nativeBytes) {
    const SECTOR_SIZE = 512;
    const MINI_SECTOR_SIZE = 64;
    const MINI_CUTOFF = 4096;
    const FREESECT = 0xffffffff;
    const ENDOFCHAIN = 0xfffffffe;
    const FATSECT = 0xfffffffd;
    const NOSTREAM = 0xffffffff;

    const compObjBytes = buildCompObjStream();
    const oleBytes = buildOleStream();

    // Phân loại các stream nhỏ (< 4096 byte) vào MiniStream và stream lớn vào FAT thường
    const streams = [
      { name: "\x01CompObj", data: compObjBytes, dirIdx: 1 },
      { name: "\x01Ole", data: oleBytes, dirIdx: 2 },
      { name: "Equation Native", data: nativeBytes, dirIdx: 3 }
    ];

    const miniChunks = [];
    const miniFatEntries = [];
    const streamMeta = new Map();

    streams.forEach((st) => {
      if (st.data.length < MINI_CUTOFF) {
        const startMini = miniFatEntries.length;
        const count = Math.max(1, Math.ceil(st.data.length / MINI_SECTOR_SIZE));
        for (let i = 0; i < count; i++) {
          const chunk = new Uint8Array(MINI_SECTOR_SIZE);
          const srcOff = i * MINI_SECTOR_SIZE;
          chunk.set(st.data.subarray(srcOff, Math.min(srcOff + MINI_SECTOR_SIZE, st.data.length)));
          miniChunks.push(chunk);
          miniFatEntries.push(i === count - 1 ? ENDOFCHAIN : startMini + i + 1);
        }
        streamMeta.set(st.dirIdx, { start: startMini, size: st.data.length, isMini: true });
      }
    });

    // Ghép toàn bộ MiniStream
    const miniStreamBytes = new Uint8Array(miniChunks.length * MINI_SECTOR_SIZE);
    miniChunks.forEach((chunk, i) => miniStreamBytes.set(chunk, i * MINI_SECTOR_SIZE));

    // Cấp phát các sector 512-byte:
    // Sector 0: FAT (sẽ hỗ trợ tối đa 128 sector = 64KB, đủ rộng cho cả công thức ma trận lớn)
    // Sector 1: Directory (4 entries x 128 bytes = 512 bytes)
    // Sector 2..: MiniFAT (nếu có) -> MiniStream (nếu có) -> Large streams (nếu có)
    const fat = [];
    const dataSectors = [];

    function allocChain(payloadBytes) {
      const count = Math.max(1, Math.ceil(payloadBytes.length / SECTOR_SIZE));
      const startSect = dataSectors.length;
      for (let i = 0; i < count; i++) {
        const sec = new Uint8Array(SECTOR_SIZE);
        const srcOff = i * SECTOR_SIZE;
        sec.set(payloadBytes.subarray(srcOff, Math.min(srcOff + SECTOR_SIZE, payloadBytes.length)));
        dataSectors.push(sec);
        fat.push(i === count - 1 ? ENDOFCHAIN : startSect + i + 1);
      }
      return { startSect, count };
    }

    // Dành chỗ Sector 0 cho FAT và Sector 1 cho Directory
    dataSectors.push(new Uint8Array(SECTOR_SIZE)); // placeholder sector 0 (FAT)
    fat.push(FATSECT);
    dataSectors.push(new Uint8Array(SECTOR_SIZE)); // placeholder sector 1 (Dir)
    fat.push(ENDOFCHAIN);

    let firstMiniFatSect = ENDOFCHAIN;
    let numMiniFatSects = 0;
    let firstMiniStreamSect = ENDOFCHAIN;

    if (miniFatEntries.length > 0) {
      const miniFatBytes = new Uint8Array(Math.ceil(miniFatEntries.length / 128) * SECTOR_SIZE);
      const mfView = new DataView(miniFatBytes.buffer);
      for (let i = 0; i < miniFatBytes.length / 4; i++) {
        mfView.setUint32(i * 4, i < miniFatEntries.length ? miniFatEntries[i] : FREESECT, true);
      }
      const mfAlloc = allocChain(miniFatBytes);
      firstMiniFatSect = mfAlloc.startSect;
      numMiniFatSects = mfAlloc.count;

      const msAlloc = allocChain(miniStreamBytes);
      firstMiniStreamSect = msAlloc.startSect;
    }

    // Cấp phát các stream >= 4096 byte trực tiếp vào FAT thường
    streams.forEach((st) => {
      if (st.data.length >= MINI_CUTOFF) {
        const alloc = allocChain(st.data);
        streamMeta.set(st.dirIdx, { start: alloc.startSect, size: st.data.length, isMini: false });
      }
    });

    // Ghi FAT vào sector 0 (nếu vượt 128 sector, mở rộng thêm FAT sector)
    const fatSector0 = dataSectors[0];
    const fatView = new DataView(fatSector0.buffer);
    for (let i = 0; i < 128; i++) {
      fatView.setUint32(i * 4, i < fat.length ? fat[i] : FREESECT, true);
    }

    // Ghi Directory Entries vào sector 1
    const dirSector = dataSectors[1];
    function writeDirEntry(slotIdx, name, objType, color, leftId, rightId, childId, startSect, size, clsidBytes) {
      const off = slotIdx * 128;
      const dv = new DataView(dirSector.buffer, off, 128);
      for (let i = 0; i < name.length && i < 31; i++) {
        dv.setUint16(i * 2, name.charCodeAt(i), true);
      }
      dv.setUint16(name.length * 2, 0, true);
      dv.setUint16(64, (name.length + 1) * 2, true); // Name length in bytes (UTF-16LE + NUL)
      dv.setUint8(66, objType);
      dv.setUint8(67, color);
      dv.setUint32(68, leftId, true);
      dv.setUint32(72, rightId, true);
      dv.setUint32(76, childId, true);
      if (clsidBytes) {
        dirSector.set(clsidBytes, off + 80);
      }
      dv.setUint32(116, startSect, true);
      dv.setUint32(120, size >>> 0, true);
    }

    // Ghi CLSID chuẩn Design Science / Wiris MathType 7.0 vào Root Entry (offset 0x50 = byte 80)
    // Cây đỏ-đen MS-CFB hợp lệ theo thứ tự độ dài tên: "\x01Ole" (4) < "\x01CompObj" (8) < "Equation Native" (15)
    writeDirEntry(0, "Root Entry", 5, 1, NOSTREAM, NOSTREAM, 1, firstMiniStreamSect, miniStreamBytes.length, MATHTYPE7_CLSID);
    const m1 = streamMeta.get(1);
    const m2 = streamMeta.get(2);
    const m3 = streamMeta.get(3);
    writeDirEntry(1, "\x01CompObj", 2, 1, 2, 3, NOSTREAM, m1.start, m1.size, null);
    writeDirEntry(2, "\x01Ole", 2, 0, NOSTREAM, NOSTREAM, NOSTREAM, m2.start, m2.size, null);
    writeDirEntry(3, "Equation Native", 2, 0, NOSTREAM, NOSTREAM, NOSTREAM, m3.start, m3.size, null);

    // Header 512 bytes chuẩn CFBF v3
    const header = new Uint8Array(SECTOR_SIZE);
    const hv = new DataView(header.buffer);
    header.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0); // Magic signature
    hv.setUint16(24, 0x003e, true); // Minor version
    hv.setUint16(26, 0x0003, true); // Major version 3
    hv.setUint16(28, 0xfffe, true); // Little-endian byte order
    hv.setUint16(30, 9, true);      // Sector shift (2^9 = 512)
    hv.setUint16(32, 6, true);      // Mini-sector shift (2^6 = 64)
    hv.setUint32(44, 1, true);      // Number of FAT sectors = 1
    hv.setUint32(48, 1, true);      // First Directory sector = 1
    hv.setUint32(56, MINI_CUTOFF, true); // Mini stream cutoff = 4096
    hv.setUint32(60, firstMiniFatSect, true);
    hv.setUint32(64, numMiniFatSects, true);
    hv.setUint32(68, ENDOFCHAIN, true); // First DIFAT sector
    hv.setUint32(72, 0, true);          // Num DIFAT sectors

    // DIFAT array (109 entries): entry 0 = sector 0 (FAT), còn lại FREESECT
    hv.setUint32(76, 0, true);
    for (let i = 1; i < 109; i++) {
      hv.setUint32(76 + i * 4, FREESECT, true);
    }

    // Gộp Header + tất cả các Sector thành tệp OLE2 hoàn chỉnh
    const total = new Uint8Array((1 + dataSectors.length) * SECTOR_SIZE);
    total.set(header, 0);
    for (let i = 0; i < dataSectors.length; i++) {
      total.set(dataSectors[i], (i + 1) * SECTOR_SIZE);
    }
    return total;
  }

  /**
   * Vẽ nhanh hình ảnh xem trước PNG Base64 từ cây AST trên HTML5 <canvas> ẩn
   * dùng font Times New Roman sắc nét (tỷ lệ 2x DPI, không bị lỗi taint canvas).
   */
  function renderAstPreviewPng(astNodes, fontSizePt, displayMode) {
    // Fallback PNG 1x1 trong suốt nếu môi trường không có Canvas 2D
    const FALLBACK_PNG =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";

    const canvas = dom.mtPreviewCanvas || document.createElement("canvas");
    if (!canvas || typeof canvas.getContext !== "function") {
      return { pngBase64: FALLBACK_PNG, widthPt: 48, heightPt: fontSizePt * 1.4, descentPt: fontSizePt * 0.25 };
    }
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      return { pngBase64: FALLBACK_PNG, widthPt: 48, heightPt: fontSizePt * 1.4, descentPt: fontSizePt * 0.25 };
    }

    const ptToPx = 96 / 72;
    const basePx = fontSizePt * ptToPx * (displayMode ? 1.12 : 1.0);
    const scale = 2; // 2x Retina để hiển thị rõ nét trên Word

    function fontStr(sizePx, italic, bold) {
      const style = italic ? "italic " : "";
      const weight = bold ? "bold " : "";
      return style + weight + Math.max(7, sizePx).toFixed(1) + 'px "Times New Roman", "Cambria Math", Georgia, serif';
    }

    function measureTextBox(text, sizePx, italic, bold) {
      ctx.font = fontStr(sizePx, italic, bold);
      const m = ctx.measureText(text);
      const w = Math.max(m.width, sizePx * 0.25);
      const ascent = sizePx * 0.78;
      const descent = sizePx * 0.24;
      return { width: w, ascent, descent, draw: (c, x, y) => {
        c.font = fontStr(sizePx, italic, bold);
        c.fillText(text, x, y);
      }};
    }

    function layoutSeq(nodes, sizePx) {
      const boxes = [];
      for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i];
        const next = nodes[i + 1];
        if (next && next.kind === "script" && node.kind !== "nary" && node.kind !== "lim") {
          const baseBox = layoutNode(node, sizePx);
          const subBox = next.sub ? layoutSeq(next.sub, sizePx * 0.68) : null;
          const supBox = next.sup ? layoutSeq(next.sup, sizePx * 0.68) : null;
          const scriptW = Math.max(subBox ? subBox.width : 0, supBox ? supBox.width : 0);
          const supRaise = sizePx * 0.42;
          const subDrop = sizePx * 0.28;
          const ascent = Math.max(baseBox.ascent, supBox ? supRaise + supBox.ascent : 0);
          const descent = Math.max(baseBox.descent, subBox ? subDrop + subBox.descent : 0);
          boxes.push({
            width: baseBox.width + scriptW + 1,
            ascent,
            descent,
            draw: (c, x, y) => {
              baseBox.draw(c, x, y);
              if (supBox) supBox.draw(c, x + baseBox.width + 0.5, y - supRaise);
              if (subBox) subBox.draw(c, x + baseBox.width + 0.5, y + subDrop);
            }
          });
          i += 1;
        } else {
          boxes.push(layoutNode(node, sizePx));
        }
      }

      let totalW = 0;
      let maxAsc = sizePx * 0.72;
      let maxDesc = sizePx * 0.22;
      boxes.forEach((b) => {
        totalW += b.width;
        if (b.ascent > maxAsc) maxAsc = b.ascent;
        if (b.descent > maxDesc) maxDesc = b.descent;
      });

      return {
        width: totalW,
        ascent: maxAsc,
        descent: maxDesc,
        draw: (c, x, y) => {
          let cx = x;
          boxes.forEach((b) => {
            b.draw(c, cx, y);
            cx += b.width;
          });
        }
      };
    }

    function layoutNode(node, sizePx) {
      if (!node) return { width: 0, ascent: 0, descent: 0, draw: () => {} };

      switch (node.kind) {
        case "char": {
          const isVar = node.tf === MT_FN_VARIABLE || node.tf === MT_FN_LCGREEK;
          const isOp = /^[+=−<>≤≥≠≈±×÷·∈→⇒⇔]$/.test(node.ch);
          const pad = isOp ? sizePx * 0.16 : 0;
          const inner = measureTextBox(node.ch, sizePx, isVar, node.tf === MT_FN_VECTOR);
          return {
            width: inner.width + pad * 2,
            ascent: inner.ascent,
            descent: inner.descent,
            draw: (c, x, y) => inner.draw(c, x + pad, y)
          };
        }

        case "space":
          return { width: sizePx * 0.26, ascent: sizePx * 0.5, descent: 0, draw: () => {} };

        case "text":
          return measureTextBox(node.text || "", sizePx, false, node.tf === MT_FN_VECTOR);

        case "func": {
          const b = measureTextBox(node.name || "", sizePx, false, false);
          return {
            width: b.width + sizePx * 0.16,
            ascent: b.ascent,
            descent: b.descent,
            draw: (c, x, y) => b.draw(c, x, y)
          };
        }

        case "group":
          return layoutSeq(node.items || [], sizePx);

        case "script": {
          const subBox = node.sub ? layoutSeq(node.sub, sizePx * 0.68) : null;
          const supBox = node.sup ? layoutSeq(node.sup, sizePx * 0.68) : null;
          const w = Math.max(subBox ? subBox.width : 0, supBox ? supBox.width : 0);
          return {
            width: w + 1,
            ascent: supBox ? sizePx * 0.42 + supBox.ascent : sizePx * 0.5,
            descent: subBox ? sizePx * 0.28 + subBox.descent : sizePx * 0.2,
            draw: (c, x, y) => {
              if (supBox) supBox.draw(c, x, y - sizePx * 0.42);
              if (subBox) subBox.draw(c, x, y + sizePx * 0.28);
            }
          };
        }

        case "frac": {
          const childSize = sizePx * 0.86;
          const numBox = layoutSeq(node.num || [], childSize);
          const denBox = layoutSeq(node.den || [], childSize);
          const barPad = sizePx * 0.18;
          const barW = Math.max(numBox.width, denBox.width) + barPad * 2;
          const axisY = sizePx * 0.30; // Trục phân số phía trên baseline
          const gap = sizePx * 0.14;
          const ascent = axisY + gap + numBox.descent + numBox.ascent;
          const descent = Math.max(sizePx * 0.22, gap + denBox.ascent + denBox.descent - axisY);
          return {
            width: barW + sizePx * 0.16,
            ascent,
            descent,
            draw: (c, x, y) => {
              const bx = x + sizePx * 0.08;
              const lineY = y - axisY;
              numBox.draw(c, bx + (barW - numBox.width) / 2, lineY - gap - numBox.descent);
              denBox.draw(c, bx + (barW - denBox.width) / 2, lineY + gap + denBox.ascent);
              c.beginPath();
              c.lineWidth = Math.max(0.85, sizePx * 0.065);
              c.moveTo(bx, lineY);
              c.lineTo(bx + barW, lineY);
              c.stroke();
            }
          };
        }

        case "root": {
          const bodyBox = layoutSeq(node.body || [], sizePx);
          const degBox = node.deg ? layoutSeq(node.deg, sizePx * 0.58) : null;
          const radW = sizePx * 0.58;
          const degW = degBox ? Math.max(0, degBox.width - radW * 0.45) : 0;
          const topPad = sizePx * 0.18;
          const ascent = Math.max(bodyBox.ascent + topPad, degBox ? degBox.ascent + sizePx * 0.55 : 0);
          const descent = Math.max(bodyBox.descent, sizePx * 0.22);
          return {
            width: degW + radW + bodyBox.width + sizePx * 0.14,
            ascent,
            descent,
            draw: (c, x, y) => {
              const rx = x + degW;
              const topY = y - bodyBox.ascent - topPad + 1;
              const botY = y + Math.min(bodyBox.descent, sizePx * 0.18);
              if (degBox) {
                degBox.draw(c, x, y - sizePx * 0.42);
              }
              c.beginPath();
              c.lineWidth = Math.max(0.9, sizePx * 0.065);
              c.moveTo(rx + radW * 0.05, y - sizePx * 0.22);
              c.lineTo(rx + radW * 0.24, y - sizePx * 0.28);
              c.lineTo(rx + radW * 0.48, botY);
              c.lineTo(rx + radW * 0.88, topY);
              c.lineTo(rx + radW + bodyBox.width + sizePx * 0.08, topY);
              c.stroke();
              bodyBox.draw(c, rx + radW, y);
            }
          };
        }

        case "fence": {
          const bodyBox = layoutSeq(node.body || [], sizePx);
          const left = node.left && node.left !== "." ? node.left : "";
          const right = node.right && node.right !== "." ? node.right : "";
          const totalH = bodyBox.ascent + bodyBox.descent;
          const delimSize = Math.max(sizePx, Math.min(totalH * 0.96, sizePx * 2.4));
          const lBox = left ? measureTextBox(left, delimSize, false, false) : { width: 0, ascent: 0, descent: 0, draw: () => {} };
          const rBox = right ? measureTextBox(right, delimSize, false, false) : { width: 0, ascent: 0, descent: 0, draw: () => {} };
          const delimYOffset = (bodyBox.descent - bodyBox.ascent) / 2 + delimSize * 0.30;
          return {
            width: lBox.width + bodyBox.width + rBox.width + 2,
            ascent: Math.max(bodyBox.ascent, delimSize * 0.72),
            descent: Math.max(bodyBox.descent, delimSize * 0.24),
            draw: (c, x, y) => {
              if (left) lBox.draw(c, x, y + delimYOffset);
              bodyBox.draw(c, x + lBox.width + 1, y);
              if (right) rBox.draw(c, x + lBox.width + bodyBox.width + 1, y + delimYOffset);
            }
          };
        }

        case "nary": {
          const symMap = {
            int: "∫", iint: "∬", iiint: "∭", oint: "∮",
            sum: "∑", prod: "∏", coprod: "∐", bigcup: "⋃", bigcap: "⋂"
          };
          const sym = symMap[node.op] || "∫";
          const bigSize = sizePx * 1.55;
          const symBox = measureTextBox(sym, bigSize, false, false);
          const subBox = node.sub ? layoutSeq(node.sub, sizePx * 0.65) : null;
          const supBox = node.sup ? layoutSeq(node.sup, sizePx * 0.65) : null;
          const isInt = node.op.indexOf("int") !== -1;

          if (isInt) {
            const limW = Math.max(subBox ? subBox.width : 0, supBox ? supBox.width : 0);
            const supY = sizePx * 0.85;
            const subY = sizePx * 0.58;
            return {
              width: symBox.width + limW + sizePx * 0.2,
              ascent: Math.max(symBox.ascent * 0.88, supBox ? supY + supBox.ascent : 0),
              descent: Math.max(symBox.descent + sizePx * 0.22, subBox ? subY + subBox.descent : 0),
              draw: (c, x, y) => {
                symBox.draw(c, x, y + sizePx * 0.16);
                if (supBox) supBox.draw(c, x + symBox.width + 1, y - supY);
                if (subBox) subBox.draw(c, x + symBox.width - 1, y + subY);
              }
            };
          }

          const colW = Math.max(symBox.width, subBox ? subBox.width : 0, supBox ? supBox.width : 0);
          const asc = symBox.ascent * 0.82 + (supBox ? supBox.ascent + supBox.descent + 2 : 0);
          const desc = symBox.descent + sizePx * 0.12 + (subBox ? subBox.ascent + subBox.descent + 2 : 0);
          return {
            width: colW + sizePx * 0.22,
            ascent: asc,
            descent: desc,
            draw: (c, x, y) => {
              symBox.draw(c, x + (colW - symBox.width) / 2, y + sizePx * 0.12);
              if (supBox) supBox.draw(c, x + (colW - supBox.width) / 2, y - symBox.ascent * 0.82 - supBox.descent - 1);
              if (subBox) subBox.draw(c, x + (colW - subBox.width) / 2, y + symBox.descent + sizePx * 0.12 + subBox.ascent + 1);
            }
          };
        }

        case "lim": {
          const limBox = measureTextBox("lim", sizePx, false, false);
          const subBox = node.sub ? layoutSeq(node.sub, sizePx * 0.65) : null;
          const colW = Math.max(limBox.width, subBox ? subBox.width : 0);
          return {
            width: colW + sizePx * 0.22,
            ascent: limBox.ascent,
            descent: limBox.descent + (subBox ? subBox.ascent + subBox.descent + 1 : 0),
            draw: (c, x, y) => {
              limBox.draw(c, x + (colW - limBox.width) / 2, y);
              if (subBox) subBox.draw(c, x + (colW - subBox.width) / 2, y + limBox.descent + subBox.ascent + 1);
            }
          };
        }

        case "accent": {
          const bodyBox = layoutSeq(node.body || [], sizePx);
          const extraTop = node.cmd === "underline" ? 0 : sizePx * 0.30;
          const extraBot = node.cmd === "underline" ? sizePx * 0.20 : 0;
          return {
            width: bodyBox.width + 2,
            ascent: bodyBox.ascent + extraTop,
            descent: bodyBox.descent + extraBot,
            draw: (c, x, y) => {
              bodyBox.draw(c, x + 1, y);
              c.beginPath();
              c.lineWidth = Math.max(0.85, sizePx * 0.06);
              if (node.cmd === "underline") {
                const uy = y + bodyBox.descent + sizePx * 0.10;
                c.moveTo(x + 1, uy);
                c.lineTo(x + 1 + bodyBox.width, uy);
                c.stroke();
              } else if (node.cmd === "vec" || node.cmd === "overrightarrow") {
                const ay = y - bodyBox.ascent - sizePx * 0.12;
                c.moveTo(x + 1, ay);
                c.lineTo(x + 1 + bodyBox.width, ay);
                c.lineTo(x + 1 + bodyBox.width - 3, ay - 2.5);
                c.moveTo(x + 1 + bodyBox.width, ay);
                c.lineTo(x + 1 + bodyBox.width - 3, ay + 2.5);
                c.stroke();
              } else if (node.cmd === "hat" || node.cmd === "widehat") {
                const hy = y - bodyBox.ascent - sizePx * 0.08;
                c.moveTo(x + 1, hy);
                c.lineTo(x + 1 + bodyBox.width / 2, hy - sizePx * 0.18);
                c.lineTo(x + 1 + bodyBox.width, hy);
                c.stroke();
              } else {
                const oy = y - bodyBox.ascent - sizePx * 0.10;
                c.moveTo(x + 1, oy);
                c.lineTo(x + 1 + bodyBox.width, oy);
                c.stroke();
              }
            }
          };
        }

        case "env": {
          const rows = (node.rows || []).map((r) => r.map((cell) => layoutSeq(cell, sizePx * 0.92)));
          const rowCount = rows.length || 1;
          const colCount = Math.max(1, ...rows.map((r) => r.length));
          const colWidths = new Array(colCount).fill(0);
          const rowHeights = rows.map((r) => {
            let maxA = sizePx * 0.7;
            let maxD = sizePx * 0.25;
            r.forEach((cell, cIdx) => {
              if (cell.width > colWidths[cIdx]) colWidths[cIdx] = cell.width;
              if (cell.ascent > maxA) maxA = cell.ascent;
              if (cell.descent > maxD) maxD = cell.descent;
            });
            return { ascent: maxA, descent: maxD };
          });

          const colGap = sizePx * 0.5;
          const rowGap = sizePx * 0.28;
          const gridW = colWidths.reduce((a, b) => a + b, 0) + colGap * Math.max(0, colCount - 1);
          const totalH =
            rowHeights.reduce((a, r) => a + r.ascent + r.descent, 0) +
            rowGap * Math.max(0, rowCount - 1);
          const leftDelim = node.env === "cases" ? "{" : node.env === "pmatrix" ? "(" : node.env === "bmatrix" ? "[" : "";
          const rightDelim = node.env === "pmatrix" ? ")" : node.env === "bmatrix" ? "]" : "";
          const delimW = leftDelim ? sizePx * 0.55 : 0;
          const rDelimW = rightDelim ? sizePx * 0.55 : 0;
          const halfH = totalH / 2;

          return {
            width: delimW + gridW + rDelimW + 4,
            ascent: halfH + sizePx * 0.25,
            descent: Math.max(sizePx * 0.25, halfH - sizePx * 0.25),
            draw: (c, x, y) => {
              const topY = y - halfH - sizePx * 0.25;
              if (leftDelim) {
                const ds = Math.min(totalH * 1.05, sizePx * 3.2);
                c.font = fontStr(ds, false, false);
                c.fillText(leftDelim, x, topY + totalH * 0.78);
              }
              let cy = topY;
              for (let r = 0; r < rows.length; r++) {
                const rh = rowHeights[r];
                let cx = x + delimW + 2;
                for (let col = 0; col < rows[r].length; col++) {
                  const cell = rows[r][col];
                  const alignX = node.env === "cases" ? cx : cx + (colWidths[col] - cell.width) / 2;
                  cell.draw(c, alignX, cy + rh.ascent);
                  cx += colWidths[col] + colGap;
                }
                cy += rh.ascent + rh.descent + rowGap;
              }
              if (rightDelim) {
                const ds = Math.min(totalH * 1.05, sizePx * 3.2);
                c.font = fontStr(ds, false, false);
                c.fillText(rightDelim, x + delimW + gridW + 3, topY + totalH * 0.78);
              }
            }
          };
        }

        default:
          return { width: 0, ascent: 0, descent: 0, draw: () => {} };
      }
    }

    const rootBox = layoutSeq(astNodes, basePx);
    const padX = Math.ceil(basePx * 0.22);
    const padY = Math.ceil(basePx * 0.16);
    const cssW = Math.max(16, Math.ceil(rootBox.width + padX * 2));
    const cssH = Math.max(16, Math.ceil(rootBox.ascent + rootBox.descent + padY * 2));

    canvas.width = cssW * scale;
    canvas.height = cssH * scale;
    ctx.setTransform(scale, 0, 0, scale, 0, 0);
    ctx.clearRect(0, 0, cssW, cssH);
    ctx.fillStyle = "#000000";
    ctx.strokeStyle = "#000000";
    ctx.textBaseline = "alphabetic";

    rootBox.draw(ctx, padX, padY + rootBox.ascent);

    const dataUrl = canvas.toDataURL("image/png");
    const pngBase64 = dataUrl.replace(/^data:image\/png;base64,/, "") || FALLBACK_PNG;

    return {
      pngBase64,
      widthPt: Number((cssW / ptToPx).toFixed(2)),
      heightPt: Number((cssH / ptToPx).toFixed(2)),
      descentPt: Number(((rootBox.descent + padY) / ptToPx).toFixed(2))
    };
  }

  /**
  /**
   * Đọc thông báo lỗi kỹ thuật đầy đủ từ Office.js / Word API (gồm error code và debugInfo).
   */
  function formatWordApiError(err) {
    if (!err) return "Lỗi không xác định từ Word API.";
    const parts = [];
    if (err.code) parts.push("[" + err.code + "]");
    if (err.message) parts.push(String(err.message));
    if (err.debugInfo && err.debugInfo.message && err.debugInfo.message !== err.message) {
      parts.push(String(err.debugInfo.message));
    }
    if (err.debugInfo && err.debugInfo.errorLocation) {
      parts.push("(tại " + err.debugInfo.errorLocation + ")");
    }
    return parts.length > 0 ? parts.join(" ") : String(err);
  }

  /**
   * Bộ tạo tệp nhị phân Vector EMF (Windows Enhanced Metafile — MS-EMF) thuần JavaScript
   * bằng Uint8Array & DataView (1 pt = 20 logical units, 1440 DPI chuẩn Word twips).
   * Sinh đầy đủ các bản ghi GDI:
   * - EMR_HEADER (108 bytes), EMR_SETMAPMODE, EMR_SETWINDOWEXTEX/VIEWPORTEXTEX
   * - EMR_SETBKMODE (TRANSPARENT), EMR_SETTEXTCOLOR, EMR_SETTEXTALIGN (TA_LEFT | TA_BASELINE)
   * - EMR_EXTCREATEFONTINDIRECTW (Times New Roman, in nghiêng cho biến, đứng cho số & phép toán)
   * - EMR_EXTTEXTOUTW (UTF-16LE + OutputDx), EMR_CREATEPEN + EMR_MOVETOEX + EMR_LINETO
   * - EMR_DELETEOBJECT giải phóng handle GDI & EMR_EOF.
   *
   * Trả về { emfBase64, emfBytes, widthPt, heightPt, depthPt } với depthPt là khoảng cách
   * từ đường cơ sở (baseline của dấu '=') xuống đáy công thức để căn <w:position> chuẩn xác.
   */
  function renderAstToEmf(astNodes, fontSizePt, displayMode) {
    const UNITS_PER_PT = 20; // 1 pt = 20 twips (1440 DPI -> khớp hoàn toàn đơn vị dxa của Word)
    const baseSizePt = Number(fontSizePt || 12) * (displayMode ? 1.08 : 1.0);
    const baseSizeU = Math.round(baseSizePt * UNITS_PER_PT);

    // Dùng Canvas 2D ẩn (nếu có) để đo bề rộng ký tự Times New Roman chính xác hơn; nếu không có thì dùng bảng tỷ lệ chuẩn
    let measureCtx = null;
    try {
      const c = dom.mtPreviewCanvas || (typeof document !== "undefined" ? document.createElement("canvas") : null);
      if (c && typeof c.getContext === "function") {
        measureCtx = c.getContext("2d");
      }
    } catch {
      measureCtx = null;
    }

    function estimateCharWidthRatio(ch, italic, bold) {
      if (ch === " ") return 0.28;
      if ("iljtI1.,;:'|!()[]".indexOf(ch) !== -1) return italic ? 0.34 : 0.32;
      if ("fr".indexOf(ch) !== -1) return 0.38;
      if ("mwMW∑∏∫∬∭∮∞⇔⇒".indexOf(ch) !== -1) return 0.80;
      if (/[A-Z]/.test(ch)) return bold ? 0.72 : 0.68;
      if (/[0-9]/.test(ch)) return 0.52;
      if ("+-−=<>≤≥≠≈±×÷·∈→".indexOf(ch) !== -1) return 0.58;
      return bold ? 0.54 : 0.50;
    }

    function measureCharAdvanceU(ch, sizeU, italic, bold) {
      if (measureCtx) {
        try {
          const pxSize = (sizeU / UNITS_PER_PT) * (96 / 72);
          const stylePrefix = (italic ? "italic " : "") + (bold ? "bold " : "");
          measureCtx.font = stylePrefix + Math.max(6, pxSize).toFixed(2) + 'px "Times New Roman", serif';
          const wPx = measureCtx.measureText(ch).width;
          if (wPx > 0) {
            return Math.max(Math.round(sizeU * 0.20), Math.round((wPx * 72 / 96) * UNITS_PER_PT));
          }
        } catch {
          // Rơi về bảng đo chuẩn
        }
      }
      return Math.max(Math.round(sizeU * 0.22), Math.round(sizeU * estimateCharWidthRatio(ch, italic, bold)));
    }

    function measureTextRun(text, sizeU, italic, bold) {
      const str = String(text || "");
      const dx = [];
      let totalW = 0;
      for (let i = 0; i < str.length; i++) {
        const adv = measureCharAdvanceU(str[i], sizeU, italic, bold);
        dx.push(adv);
        totalW += adv;
      }
      // Bù nhẹ độ nghiêng bên phải (italic overhang) cho biến chữ nghiêng để không bị cắt sát mép
      const italicExtra = italic && str.length > 0 ? Math.round(sizeU * 0.08) : 0;
      const ascent = Math.round(sizeU * 0.80);
      const descent = Math.round(sizeU * 0.24);
      return {
        width: totalW + italicExtra,
        ascent,
        descent,
        draw: (emf, x, baselineY) => {
          if (!str) return;
          emf.drawText(str, dx, x, baselineY, sizeU, italic, bold, totalW + italicExtra);
        }
      };
    }

    function layoutSeq(nodes, sizeU) {
      const boxes = [];
      for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i];
        const next = nodes[i + 1];

        if (next && next.kind === "script" && node.kind !== "nary" && node.kind !== "lim") {
          const baseBox = layoutNode(node, sizeU);
          const scriptSizeU = Math.max(100, Math.round(sizeU * 0.64));
          const subBox = next.sub ? layoutSeq(next.sub, scriptSizeU) : null;
          const supBox = next.sup ? layoutSeq(next.sup, scriptSizeU) : null;
          const scriptW = Math.max(subBox ? subBox.width : 0, supBox ? supBox.width : 0);
          const supRaise = Math.round(sizeU * 0.42);
          const subDrop = Math.round(sizeU * 0.26);
          const gapX = Math.round(sizeU * 0.04);
          const ascent = Math.max(baseBox.ascent, supBox ? supRaise + supBox.ascent : 0);
          const descent = Math.max(baseBox.descent, subBox ? subDrop + subBox.descent : 0);

          boxes.push({
            width: baseBox.width + gapX + scriptW,
            ascent,
            descent,
            draw: (emf, x, y) => {
              baseBox.draw(emf, x, y);
              if (supBox) supBox.draw(emf, x + baseBox.width + gapX, y - supRaise);
              if (subBox) subBox.draw(emf, x + baseBox.width + gapX, y + subDrop);
            }
          });
          i += 1;
        } else {
          boxes.push(layoutNode(node, sizeU));
        }
      }

      let totalW = 0;
      let maxAsc = Math.round(sizeU * 0.76);
      let maxDesc = Math.round(sizeU * 0.22);
      boxes.forEach((b) => {
        totalW += b.width;
        if (b.ascent > maxAsc) maxAsc = b.ascent;
        if (b.descent > maxDesc) maxDesc = b.descent;
      });

      return {
        width: totalW,
        ascent: maxAsc,
        descent: maxDesc,
        draw: (emf, x, y) => {
          let cx = x;
          boxes.forEach((b) => {
            b.draw(emf, cx, y);
            cx += b.width;
          });
        }
      };
    }

    function layoutNode(node, sizeU) {
      if (!node) return { width: 0, ascent: 0, descent: 0, draw: () => {} };

      switch (node.kind) {
        case "char": {
          const isItalic = node.tf === MT_FN_VARIABLE || node.tf === MT_FN_LCGREEK;
          const isBold = node.tf === MT_FN_VECTOR;
          const isBinaryOp = /^[+=−<>≤≥≠≈±×÷·∈→⇒⇔]$/.test(node.ch || "");
          const pad = isBinaryOp ? Math.round(sizeU * 0.16) : 0;
          const inner = measureTextRun(node.ch || "", sizeU, isItalic, isBold);
          return {
            width: inner.width + pad * 2,
            ascent: inner.ascent,
            descent: inner.descent,
            draw: (emf, x, y) => inner.draw(emf, x + pad, y)
          };
        }

        case "space":
          return { width: Math.round(sizeU * 0.26), ascent: Math.round(sizeU * 0.5), descent: 0, draw: () => {} };

        case "text":
          return measureTextRun(node.text || "", sizeU, false, node.tf === MT_FN_VECTOR);

        case "func": {
          const b = measureTextRun(node.name || "", sizeU, false, false);
          return {
            width: b.width + Math.round(sizeU * 0.14),
            ascent: b.ascent,
            descent: b.descent,
            draw: (emf, x, y) => b.draw(emf, x, y)
          };
        }

        case "group":
          return layoutSeq(node.items || [], sizeU);

        case "script": {
          const scriptSizeU = Math.max(100, Math.round(sizeU * 0.64));
          const subBox = node.sub ? layoutSeq(node.sub, scriptSizeU) : null;
          const supBox = node.sup ? layoutSeq(node.sup, scriptSizeU) : null;
          const w = Math.max(subBox ? subBox.width : 0, supBox ? supBox.width : 0);
          const supRaise = Math.round(sizeU * 0.42);
          const subDrop = Math.round(sizeU * 0.26);
          return {
            width: w + Math.round(sizeU * 0.05),
            ascent: supBox ? supRaise + supBox.ascent : Math.round(sizeU * 0.5),
            descent: subBox ? subDrop + subBox.descent : Math.round(sizeU * 0.22),
            draw: (emf, x, y) => {
              if (supBox) supBox.draw(emf, x, y - supRaise);
              if (subBox) subBox.draw(emf, x, y + subDrop);
            }
          };
        }

        case "frac": {
          const childSizeU = Math.max(120, Math.round(sizeU * 0.88));
          const numBox = layoutSeq(node.num || [], childSizeU);
          const denBox = layoutSeq(node.den || [], childSizeU);
          const barPad = Math.round(sizeU * 0.16);
          const sideMargin = Math.round(sizeU * 0.08);
          const barW = Math.max(numBox.width, denBox.width) + barPad * 2;
          // Trục toán học (ngang giữa dấu '=' và gạch phân số) nằm trên baseline 0.30 * sizeU
          const axisOffset = Math.round(sizeU * 0.30);
          const gap = Math.round(sizeU * 0.14);
          const strokeU = Math.max(10, Math.round(sizeU * 0.055));
          const ascent = axisOffset + gap + numBox.descent + numBox.ascent;
          const descent = Math.max(Math.round(sizeU * 0.24), gap + denBox.ascent + denBox.descent - axisOffset);

          return {
            width: barW + sideMargin * 2,
            ascent,
            descent,
            draw: (emf, x, y) => {
              const bx = x + sideMargin;
              const lineY = y - axisOffset;
              numBox.draw(emf, bx + Math.round((barW - numBox.width) / 2), lineY - gap - numBox.descent);
              denBox.draw(emf, bx + Math.round((barW - denBox.width) / 2), lineY + gap + denBox.ascent);
              emf.drawLine(bx, lineY, bx + barW, lineY, strokeU);
            }
          };
        }

        case "root": {
          const bodyBox = layoutSeq(node.body || [], sizeU);
          const degSizeU = Math.max(96, Math.round(sizeU * 0.56));
          const degBox = node.deg && node.deg.length > 0 ? layoutSeq(node.deg, degSizeU) : null;
          const radW = Math.round(sizeU * 0.58);
          const degW = degBox ? Math.max(0, degBox.width - Math.round(radW * 0.42)) : 0;
          const topPad = Math.round(sizeU * 0.18);
          const strokeU = Math.max(10, Math.round(sizeU * 0.055));
          const ascent = Math.max(bodyBox.ascent + topPad, degBox ? degBox.ascent + Math.round(sizeU * 0.54) : 0);
          const descent = Math.max(bodyBox.descent, Math.round(sizeU * 0.22));
          const tailPad = Math.round(sizeU * 0.12);

          return {
            width: degW + radW + bodyBox.width + tailPad,
            ascent,
            descent,
            draw: (emf, x, y) => {
              const rx = x + degW;
              const topY = y - bodyBox.ascent - topPad + Math.round(strokeU * 0.5);
              const botY = y + Math.min(bodyBox.descent, Math.round(sizeU * 0.18));
              if (degBox) {
                degBox.draw(emf, x, y - Math.round(sizeU * 0.42));
              }
              emf.drawPolyline(
                [
                  [rx + Math.round(radW * 0.05), y - Math.round(sizeU * 0.20)],
                  [rx + Math.round(radW * 0.24), y - Math.round(sizeU * 0.28)],
                  [rx + Math.round(radW * 0.48), botY],
                  [rx + Math.round(radW * 0.88), topY],
                  [rx + radW + bodyBox.width + Math.round(tailPad * 0.6), topY]
                ],
                strokeU
              );
              bodyBox.draw(emf, rx + radW, y);
            }
          };
        }

        case "fence": {
          const bodyBox = layoutSeq(node.body || [], sizeU);
          const left = node.left && node.left !== "." ? node.left : "";
          const right = node.right && node.right !== "." ? node.right : "";
          const totalH = bodyBox.ascent + bodyBox.descent;
          const delimSizeU = Math.max(sizeU, Math.min(Math.round(totalH * 0.96), Math.round(sizeU * 2.5)));
          const lBox = left ? measureTextRun(left, delimSizeU, false, false) : { width: 0, ascent: 0, descent: 0, draw: () => {} };
          const rBox = right ? measureTextRun(right, delimSizeU, false, false) : { width: 0, ascent: 0, descent: 0, draw: () => {} };
          const delimYOffset = Math.round((bodyBox.descent - bodyBox.ascent) / 2 + delimSizeU * 0.28);
          const pad = Math.round(sizeU * 0.04);

          return {
            width: lBox.width + bodyBox.width + rBox.width + pad * 2,
            ascent: Math.max(bodyBox.ascent, Math.round(delimSizeU * 0.72)),
            descent: Math.max(bodyBox.descent, Math.round(delimSizeU * 0.24)),
            draw: (emf, x, y) => {
              if (left) lBox.draw(emf, x, y + delimYOffset);
              bodyBox.draw(emf, x + lBox.width + pad, y);
              if (right) rBox.draw(emf, x + lBox.width + bodyBox.width + pad, y + delimYOffset);
            }
          };
        }

        case "nary": {
          const symMap = {
            int: "∫", iint: "∬", iiint: "∭", oint: "∮",
            sum: "∑", prod: "∏", coprod: "∐", bigcup: "⋃", bigcap: "⋂"
          };
          const sym = symMap[node.op] || "∫";
          const bigSizeU = Math.round(sizeU * 1.55);
          const limSizeU = Math.max(96, Math.round(sizeU * 0.62));
          const symBox = measureTextRun(sym, bigSizeU, false, false);
          const subBox = node.sub ? layoutSeq(node.sub, limSizeU) : null;
          const supBox = node.sup ? layoutSeq(node.sup, limSizeU) : null;
          const isInt = (node.op || "").indexOf("int") !== -1;

          if (isInt) {
            const limW = Math.max(subBox ? subBox.width : 0, supBox ? supBox.width : 0);
            const supY = Math.round(sizeU * 0.84);
            const subY = Math.round(sizeU * 0.56);
            return {
              width: symBox.width + limW + Math.round(sizeU * 0.18),
              ascent: Math.max(Math.round(symBox.ascent * 0.88), supBox ? supY + supBox.ascent : 0),
              descent: Math.max(symBox.descent + Math.round(sizeU * 0.22), subBox ? subY + subBox.descent : 0),
              draw: (emf, x, y) => {
                symBox.draw(emf, x, y + Math.round(sizeU * 0.16));
                if (supBox) supBox.draw(emf, x + symBox.width + Math.round(sizeU * 0.04), y - supY);
                if (subBox) subBox.draw(emf, x + symBox.width - Math.round(sizeU * 0.02), y + subY);
              }
            };
          }

          const colW = Math.max(symBox.width, subBox ? subBox.width : 0, supBox ? supBox.width : 0);
          const vGap = Math.round(sizeU * 0.10);
          const asc = Math.round(symBox.ascent * 0.82) + (supBox ? supBox.ascent + supBox.descent + vGap : 0);
          const desc = symBox.descent + Math.round(sizeU * 0.12) + (subBox ? subBox.ascent + subBox.descent + vGap : 0);
          return {
            width: colW + Math.round(sizeU * 0.20),
            ascent: asc,
            descent: desc,
            draw: (emf, x, y) => {
              symBox.draw(emf, x + Math.round((colW - symBox.width) / 2), y + Math.round(sizeU * 0.12));
              if (supBox) {
                supBox.draw(
                  emf,
                  x + Math.round((colW - supBox.width) / 2),
                  y - Math.round(symBox.ascent * 0.82) - supBox.descent - vGap
                );
              }
              if (subBox) {
                subBox.draw(
                  emf,
                  x + Math.round((colW - subBox.width) / 2),
                  y + symBox.descent + Math.round(sizeU * 0.12) + subBox.ascent + vGap
                );
              }
            }
          };
        }

        case "lim": {
          const limBox = measureTextRun("lim", sizeU, false, false);
          const subBox = node.sub ? layoutSeq(node.sub, Math.max(96, Math.round(sizeU * 0.62))) : null;
          const colW = Math.max(limBox.width, subBox ? subBox.width : 0);
          const vGap = Math.round(sizeU * 0.08);
          return {
            width: colW + Math.round(sizeU * 0.20),
            ascent: limBox.ascent,
            descent: limBox.descent + (subBox ? subBox.ascent + subBox.descent + vGap : 0),
            draw: (emf, x, y) => {
              limBox.draw(emf, x + Math.round((colW - limBox.width) / 2), y);
              if (subBox) {
                subBox.draw(emf, x + Math.round((colW - subBox.width) / 2), y + limBox.descent + subBox.ascent + vGap);
              }
            }
          };
        }

        case "accent": {
          const bodyBox = layoutSeq(node.body || [], sizeU);
          const strokeU = Math.max(10, Math.round(sizeU * 0.055));
          const extraTop = node.cmd === "underline" ? 0 : Math.round(sizeU * 0.30);
          const extraBot = node.cmd === "underline" ? Math.round(sizeU * 0.22) : 0;
          const pad = Math.round(sizeU * 0.06);
          return {
            width: bodyBox.width + pad * 2,
            ascent: bodyBox.ascent + extraTop,
            descent: bodyBox.descent + extraBot,
            draw: (emf, x, y) => {
              const bx = x + pad;
              bodyBox.draw(emf, bx, y);
              if (node.cmd === "underline") {
                const uy = y + bodyBox.descent + Math.round(sizeU * 0.10);
                emf.drawLine(bx, uy, bx + bodyBox.width, uy, strokeU);
              } else if (node.cmd === "vec" || node.cmd === "overrightarrow") {
                const ay = y - bodyBox.ascent - Math.round(sizeU * 0.12);
                const head = Math.max(24, Math.round(sizeU * 0.16));
                emf.drawLine(bx, ay, bx + bodyBox.width, ay, strokeU);
                emf.drawPolyline(
                  [
                    [bx + bodyBox.width - head, ay - Math.round(head * 0.7)],
                    [bx + bodyBox.width, ay],
                    [bx + bodyBox.width - head, ay + Math.round(head * 0.7)]
                  ],
                  strokeU
                );
              } else if (node.cmd === "hat" || node.cmd === "widehat") {
                const hy = y - bodyBox.ascent - Math.round(sizeU * 0.08);
                emf.drawPolyline(
                  [
                    [bx, hy],
                    [bx + Math.round(bodyBox.width / 2), hy - Math.round(sizeU * 0.18)],
                    [bx + bodyBox.width, hy]
                  ],
                  strokeU
                );
              } else {
                const oy = y - bodyBox.ascent - Math.round(sizeU * 0.10);
                emf.drawLine(bx, oy, bx + bodyBox.width, oy, strokeU);
              }
            }
          };
        }

        case "env": {
          const cellSizeU = Math.max(120, Math.round(sizeU * 0.92));
          const rows = (node.rows || []).map((r) => r.map((cell) => layoutSeq(cell, cellSizeU)));
          const rowCount = rows.length || 1;
          const colCount = Math.max(1, ...rows.map((r) => r.length));
          const colWidths = new Array(colCount).fill(0);
          const rowHeights = rows.map((r) => {
            let maxA = Math.round(cellSizeU * 0.74);
            let maxD = Math.round(cellSizeU * 0.24);
            r.forEach((cell, cIdx) => {
              if (cell.width > colWidths[cIdx]) colWidths[cIdx] = cell.width;
              if (cell.ascent > maxA) maxA = cell.ascent;
              if (cell.descent > maxD) maxD = cell.descent;
            });
            return { ascent: maxA, descent: maxD };
          });

          const colGap = Math.round(sizeU * 0.50);
          const rowGap = Math.round(sizeU * 0.28);
          const gridW = colWidths.reduce((a, b) => a + b, 0) + colGap * Math.max(0, colCount - 1);
          const totalH =
            rowHeights.reduce((a, r) => a + r.ascent + r.descent, 0) +
            rowGap * Math.max(0, rowCount - 1);
          const leftDelim = node.env === "cases" ? "{" : node.env === "pmatrix" ? "(" : node.env === "bmatrix" ? "[" : "";
          const rightDelim = node.env === "pmatrix" ? ")" : node.env === "bmatrix" ? "]" : "";
          const delimSizeU = Math.min(Math.round(totalH * 1.02), Math.round(sizeU * 3.2));
          const lBox = leftDelim ? measureTextRun(leftDelim, delimSizeU, false, false) : { width: 0, draw: () => {} };
          const rBox = rightDelim ? measureTextRun(rightDelim, delimSizeU, false, false) : { width: 0, draw: () => {} };
          const halfH = Math.round(totalH / 2);
          const axisOffset = Math.round(sizeU * 0.28);

          return {
            width: lBox.width + gridW + rBox.width + Math.round(sizeU * 0.16),
            ascent: halfH + axisOffset,
            descent: Math.max(Math.round(sizeU * 0.24), halfH - axisOffset),
            draw: (emf, x, y) => {
              const topY = y - halfH - axisOffset;
              if (leftDelim) {
                lBox.draw(emf, x, topY + Math.round(totalH * 0.78));
              }
              let cy = topY;
              for (let r = 0; r < rows.length; r++) {
                const rh = rowHeights[r];
                let cx = x + lBox.width + Math.round(sizeU * 0.08);
                for (let col = 0; col < rows[r].length; col++) {
                  const cell = rows[r][col];
                  const alignX = node.env === "cases" ? cx : cx + Math.round((colWidths[col] - cell.width) / 2);
                  cell.draw(emf, alignX, cy + rh.ascent);
                  cx += colWidths[col] + colGap;
                }
                cy += rh.ascent + rh.descent + rowGap;
              }
              if (rightDelim) {
                rBox.draw(emf, x + lBox.width + gridW + Math.round(sizeU * 0.10), topY + Math.round(totalH * 0.78));
              }
            }
          };
        }

        default:
          return { width: 0, ascent: 0, descent: 0, draw: () => {} };
      }
    }

    // 1. Tính toán bố cục vector toàn công thức theo đơn vị logical units (20 units = 1 pt)
    const rootBox = layoutSeq(astNodes, baseSizeU);
    const padX = Math.round(baseSizeU * 0.14);
    const padY = Math.round(baseSizeU * 0.10);
    const widthU = Math.max(120, Math.ceil(rootBox.width + padX * 2));
    const heightU = Math.max(160, Math.ceil(rootBox.ascent + rootBox.descent + padY * 2));
    const baselineY = padY + rootBox.ascent;
    const depthU = heightU - baselineY; // Độ sâu từ đường cơ sở (dấu '=') xuống đáy công thức

    const widthPt = Number((widthU / UNITS_PER_PT).toFixed(2));
    const heightPt = Number((heightU / UNITS_PER_PT).toFixed(2));
    const depthPt = Number((depthU / UNITS_PER_PT).toFixed(2));

    // 2. Bộ ghi bản ghi nhị phân MS-EMF (4-byte aligned records)
    const chunks = [];
    let totalBytes = 0;
    let recordCount = 0;

    function pushRecord(buf) {
      chunks.push(buf);
      totalBytes += buf.length;
      recordCount += 1;
    }

    function makeRec(type, size) {
      const buf = new Uint8Array(size);
      const dv = new DataView(buf.buffer);
      dv.setUint32(0, type, true);
      dv.setUint32(4, size, true);
      return { buf, dv };
    }

    function writeUtf16Fixed(dv, byteOffset, str, maxChars) {
      const len = Math.min(str.length, maxChars - 1);
      for (let i = 0; i < len; i++) {
        dv.setUint16(byteOffset + i * 2, str.charCodeAt(i), true);
      }
      dv.setUint16(byteOffset + len * 2, 0, true);
    }

    // Bản ghi 0: EMR_HEADER (108 bytes — sẽ cập nhật lại nBytes, nRecords, nHandles ở bước cuối)
    const headerRec = makeRec(1, 108);
    const hdv = headerRec.dv;
    // rclBounds (inclusive logical/device units)
    hdv.setInt32(8, 0, true);
    hdv.setInt32(12, 0, true);
    hdv.setInt32(16, widthU, true);
    hdv.setInt32(20, heightU, true);
    // rclFrame (0.01 mm HIMETRIC: 1 pt = 2540/72 = 35.2778)
    const frameW = Math.max(1, Math.round((widthU / UNITS_PER_PT) * 2540 / 72));
    const frameH = Math.max(1, Math.round((heightU / UNITS_PER_PT) * 2540 / 72));
    hdv.setInt32(24, 0, true);
    hdv.setInt32(28, 0, true);
    hdv.setInt32(32, frameW, true);
    hdv.setInt32(36, frameH, true);
    // dSignature = " EMF" (0x464D4520), nVersion = 0x00010000
    hdv.setUint32(40, 0x464d4520, true);
    hdv.setUint32(44, 0x00010000, true);
    // szlDevice = (14400, 14400) units = 10 inches @ 1440 DPI; szlMillimeters = (254, 254) mm = 10 inches
    hdv.setInt32(72, 14400, true);
    hdv.setInt32(76, 14400, true);
    hdv.setInt32(80, 254, true);
    hdv.setInt32(84, 254, true);
    // szlMicrometers = (254000, 254000) um
    hdv.setInt32(100, 254000, true);
    hdv.setInt32(104, 254000, true);
    pushRecord(headerRec.buf);

    // Thiết lập hệ tọa độ chuẩn (MM_ANISOTROPIC + Window/Viewport Ext/Org)
    function emitU32Rec(type, val) {
      const r = makeRec(type, 12);
      r.dv.setUint32(8, val, true);
      pushRecord(r.buf);
    }
    function emitPointRec(type, x, y) {
      const r = makeRec(type, 16);
      r.dv.setInt32(8, Math.round(x), true);
      r.dv.setInt32(12, Math.round(y), true);
      pushRecord(r.buf);
    }

    emitU32Rec(17, 8);                  // EMR_SETMAPMODE (MM_ANISOTROPIC = 8)
    emitPointRec(9, widthU, heightU);   // EMR_SETWINDOWEXTEX
    emitPointRec(11, widthU, heightU);  // EMR_SETVIEWPORTEXTEX
    emitPointRec(10, 0, 0);             // EMR_SETWINDOWORGEX
    emitPointRec(12, 0, 0);             // EMR_SETVIEWPORTORGEX
    emitU32Rec(18, 1);                  // EMR_SETBKMODE (TRANSPARENT = 1)
    emitU32Rec(24, 0x00000000);         // EMR_SETTEXTCOLOR (RGB Black)
    emitU32Rec(22, 0x00000018);         // EMR_SETTEXTALIGN (TA_LEFT | TA_BASELINE)

    // Quản lý handle GDI (font Times New Roman & pen nét vẽ)
    let nextHandle = 1;
    const allocatedHandles = [];
    const fontHandles = new Map();
    const penHandles = new Map();
    let activeFontHandle = 0;
    let activePenHandle = 0;

    function selectObject(ihObject) {
      emitU32Rec(37, ihObject); // EMR_SELECTOBJECT
    }

    function ensureFont(sizeU, italic, bold) {
      const h = Math.max(80, Math.round(sizeU));
      const key = h + "_" + (italic ? 1 : 0) + "_" + (bold ? 1 : 0);
      let handle = fontHandles.get(key);
      if (!handle) {
        handle = nextHandle++;
        fontHandles.set(key, handle);
        allocatedHandles.push(handle);

        // EMR_EXTCREATEFONTINDIRECTW (Type 82, 332 bytes — LogFontPanose)
        const rec = makeRec(82, 332);
        const dv = rec.dv;
        dv.setUint32(8, handle, true);
        dv.setInt32(12, -h, true); // lfHeight âm = chiều cao ký tự chính xác theo logical units
        dv.setInt32(16, 0, true);  // lfWidth = 0 (tự động theo tỷ lệ font)
        dv.setInt32(20, 0, true);  // lfEscapement
        dv.setInt32(24, 0, true);  // lfOrientation
        dv.setInt32(28, bold ? 700 : 400, true); // lfWeight
        dv.setUint8(32, italic ? 1 : 0);         // lfItalic
        dv.setUint8(33, 0);                      // lfUnderline
        dv.setUint8(34, 0);                      // lfStrikeOut
        dv.setUint8(35, 1);                      // lfCharSet = DEFAULT_CHARSET (Unicode đầy đủ)
        dv.setUint8(36, 0);                      // lfOutPrecision
        dv.setUint8(37, 0);                      // lfClipPrecision
        dv.setUint8(38, 4);                      // lfQuality = ANTIALIASED_QUALITY
        dv.setUint8(39, 0x12);                   // lfPitchAndFamily = VARIABLE_PITCH | FF_ROMAN
        writeUtf16Fixed(dv, 40, "Times New Roman", 32);
        writeUtf16Fixed(dv, 104, "Times New Roman", 64);
        writeUtf16Fixed(dv, 232, bold ? (italic ? "Bold Italic" : "Bold") : (italic ? "Italic" : "Regular"), 32);
        pushRecord(rec.buf);
      }

      if (activeFontHandle !== handle) {
        selectObject(handle);
        activeFontHandle = handle;
      }
    }

    function ensurePen(strokeU) {
      const w = Math.max(8, Math.round(strokeU));
      let handle = penHandles.get(w);
      if (!handle) {
        handle = nextHandle++;
        penHandles.set(w, handle);
        allocatedHandles.push(handle);

        // EMR_CREATEPEN (Type 38, 28 bytes)
        const rec = makeRec(38, 28);
        const dv = rec.dv;
        dv.setUint32(8, handle, true);
        dv.setUint32(12, 0, true);          // PS_SOLID
        dv.setInt32(16, w, true);           // Width.x
        dv.setInt32(20, 0, true);           // Width.y
        dv.setUint32(24, 0x00000000, true); // RGB Black
        pushRecord(rec.buf);
      }

      if (activePenHandle !== handle) {
        selectObject(handle);
        activePenHandle = handle;
      }
    }

    const emfDriver = {
      drawText(str, dxArr, x, y, sizeU, italic, bold, runW) {
        if (!str || str.length === 0) return;
        ensureFont(sizeU, italic, bold);

        const nChars = str.length;
        const strBytes = nChars * 2;
        const alignedStrBytes = Math.ceil(strBytes / 4) * 4;
        const offString = 76;
        const offDx = offString + alignedStrBytes;
        const recSize = offDx + nChars * 4;

        // EMR_EXTTEXTOUTW (Type 84)
        const rec = makeRec(84, recSize);
        const dv = rec.dv;
        const rx = Math.round(x);
        const ry = Math.round(y);
        // rclBounds
        dv.setInt32(8, Math.max(0, rx), true);
        dv.setInt32(12, Math.max(0, ry - Math.round(sizeU * 0.9)), true);
        dv.setInt32(16, Math.min(widthU, rx + Math.round(runW)), true);
        dv.setInt32(20, Math.min(heightU, ry + Math.round(sizeU * 0.35)), true);
        dv.setUint32(24, 1, true);               // iGraphicsMode = GM_COMPATIBLE (chuẩn GDI / MathType)
        dv.setFloat32(28, 2540 / 1440, true);    // exScale (0.01mm / logical unit ở 1440 DPI)
        dv.setFloat32(32, 2540 / 1440, true);    // eyScale
        // EmrText header
        dv.setInt32(36, rx, true);       // Reference.x
        dv.setInt32(40, ry, true);       // Reference.y (TA_BASELINE)
        dv.setUint32(44, nChars, true);  // Chars
        dv.setUint32(48, offString, true);
        dv.setUint32(52, 0, true);       // Options = ETO_NONE
        dv.setInt32(56, 0, true);        // Rectangle.left..bottom = 0
        dv.setInt32(60, 0, true);
        dv.setInt32(64, 0, true);
        dv.setInt32(68, 0, true);
        dv.setUint32(72, offDx, true);   // offDx

        for (let i = 0; i < nChars; i++) {
          dv.setUint16(offString + i * 2, str.charCodeAt(i), true);
        }
        for (let i = 0; i < nChars; i++) {
          dv.setUint32(offDx + i * 4, Math.max(1, Math.round(dxArr[i] || sizeU * 0.5)), true);
        }
        pushRecord(rec.buf);
      },

      drawLine(x1, y1, x2, y2, strokeU) {
        ensurePen(strokeU);
        emitPointRec(27, x1, y1); // EMR_MOVETOEX
        emitPointRec(54, x2, y2); // EMR_LINETO
      },

      drawPolyline(points, strokeU) {
        if (!points || points.length < 2) return;
        ensurePen(strokeU);
        emitPointRec(27, points[0][0], points[0][1]); // EMR_MOVETOEX
        for (let i = 1; i < points.length; i++) {
          emitPointRec(54, points[i][0], points[i][1]); // EMR_LINETO
        }
      }
    };

    // 3. Vẽ toàn bộ cây công thức toán học ra các bản ghi vector EMF
    rootBox.draw(emfDriver, padX, baselineY);

    // 4. Trả lại stock objects & xóa toàn bộ GDI handle đã cấp phát (EMR_DELETEOBJECT)
    if (activeFontHandle) selectObject(0x8000000d); // SYSTEM_FONT
    if (activePenHandle) selectObject(0x80000007);  // BLACK_PEN
    for (let i = 0; i < allocatedHandles.length; i++) {
      emitU32Rec(40, allocatedHandles[i]); // EMR_DELETEOBJECT
    }

    // 5. Bản ghi cuối: EMR_EOF (Type 14, 20 bytes)
    const eofRec = makeRec(14, 20);
    eofRec.dv.setUint32(8, 0, true);   // nPalEntries
    eofRec.dv.setUint32(12, 16, true); // offPalEntries
    eofRec.dv.setUint32(16, 20, true); // nSizeLast
    pushRecord(eofRec.buf);

    // 6. Cập nhật lại tổng kích thước, tổng số bản ghi & tổng số handle vào EMR_HEADER
    hdv.setUint32(48, totalBytes, true);   // nBytes
    hdv.setUint32(52, recordCount, true);  // nRecords
    hdv.setUint16(56, nextHandle, true);   // nHandles = maxHandle + 1

    const emfBytes = new Uint8Array(totalBytes);
    let cursor = 0;
    for (let i = 0; i < chunks.length; i++) {
      emfBytes.set(chunks[i], cursor);
      cursor += chunks[i].length;
    }

    return {
      emfBytes,
      emfBase64: uint8ArrayToBase64(emfBytes),
      widthPt,
      heightPt,
      depthPt
    };
  }

  window.renderAstToEmf = renderAstToEmf;

  /**
   * Tạo gói Flat OPC OOXML (<pkg:package>) chuẩn độc lập cho đối tượng MathType 7.0 (Equation.DSMT4):
   * - Ảnh hiển thị vector sắc nét: /word/media/image1.emf (pkg:contentType="image/x-emf")
   * - Căn chỉnh thẳng hàng dòng chữ (Inline Baseline Alignment):
   *   posHalfPt = -Math.round(depthPt * 2), chèn <w:rPr><w:position w:val="${posHalfPt}"/></w:rPr> ngay trước <w:object>
   * - Đủ các quan hệ /_rels/.rels và /word/_rels/document.xml.rels (rId_img -> media/image1.emf, rId_ole -> embeddings/oleObject1.bin)
   * - Nhúng tệp OLE2 CFBF MathType 7.0: /word/embeddings/oleObject1.bin (ProgID="Equation.DSMT4").
   */
  function createMathTypeOoxml(params) {
    const emfBase64 = String((params && (params.emfBase64 || params.imageBase64)) || "");
    const oleBase64 = String((params && params.oleBase64) || "");
    if (!emfBase64 || !oleBase64) {
      throw new Error("Thiếu dữ liệu ảnh vector EMF hoặc tệp nhị phân OLE để tạo gói OOXML MathType.");
    }

    const widthNum = Math.max(6, Number((params && params.widthPt) || 24));
    const heightNum = Math.max(8, Number((params && params.heightPt) || 14));
    const depthNum = Math.max(0, Number((params && (params.depthPt !== undefined ? params.depthPt : params.descentPt)) || 0));
    const widthPt = widthNum.toFixed(2);
    const heightPt = heightNum.toFixed(2);
    const dxaOrig = Math.max(20, Math.round(widthNum * 20));
    const dyaOrig = Math.max(20, Math.round(heightNum * 20));
    const shapeId = (params && params.shapeId) || "_x0000_i1025";
    const objectId = (params && params.objectId) || "_1700000001";
    const posHalfPt = -Math.round(depthNum * 2);

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
        '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">' +
          "<pkg:xmlData>" +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
            "</Relationships>" +
          "</pkg:xmlData>" +
        "</pkg:part>" +
        '<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">' +
          "<pkg:xmlData>" +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
              '<Relationship Id="rId_img" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.emf"/>' +
              '<Relationship Id="rId_ole" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>' +
            "</Relationships>" +
          "</pkg:xmlData>" +
        "</pkg:part>" +
        '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
          "<pkg:xmlData>" +
            '<w:document ' +
              'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
              'xmlns:v="urn:schemas-microsoft-com:vml" ' +
              'xmlns:o="urn:schemas-microsoft-com:office:office">' +
              "<w:body>" +
                "<w:p>" +
                  "<w:r>" +
                    '<w:rPr><w:position w:val="' + posHalfPt + '"/></w:rPr>' +
                    '<w:object w:dxaOrig="' + dxaOrig + '" w:dyaOrig="' + dyaOrig + '">' +
                      '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"/>' +
                      '<v:shape id="' + escapeXml(shapeId) + '" type="#_x0000_t75" style="width:' + widthPt + "pt;height:" + heightPt + 'pt" o:ole="">' +
                        '<v:imagedata r:id="rId_img" o:title=""/>' +
                      "</v:shape>" +
                      '<o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="' + escapeXml(shapeId) + '" DrawAspect="Content" ObjectID="' + escapeXml(objectId) + '" r:id="rId_ole"/>' +
                    "</w:object>" +
                  "</w:r>" +
                "</w:p>" +
              "</w:body>" +
            "</w:document>" +
          "</pkg:xmlData>" +
        "</pkg:part>" +
        '<pkg:part pkg:name="/word/media/image1.emf" pkg:contentType="image/x-emf" pkg:compression="store">' +
          "<pkg:binaryData>" + emfBase64 + "</pkg:binaryData>" +
        "</pkg:part>" +
        '<pkg:part pkg:name="/word/embeddings/oleObject1.bin" pkg:contentType="application/vnd.openxmlformats-officedocument.oleObject" pkg:compression="store">' +
          "<pkg:binaryData>" + oleBase64 + "</pkg:binaryData>" +
        "</pkg:part>" +
      "</pkg:package>"
    );
  }

  window.createMathTypeOoxml = createMathTypeOoxml;

  /**
   * Thu thập các đoạn văn bản liên tiếp (<w:r> chỉ chứa <w:t>) trong một <w:p>
   * để nhận diện công thức ngay cả khi Word tự tách 1 công thức qua nhiều thẻ <w:r>
   * mà vẫn giữ nguyên vẹn <w:rPr> (Times New Roman, cỡ chữ, in đậm/nghiêng) xung quanh.
   */
  function collectParagraphTextSpans(pEl) {
    const spans = [];
    let fullText = "";

    Array.from(pEl.children).forEach((child) => {
      if (child.localName !== "r") return;
      // Chỉ xử lý thẻ <w:r> thuần văn bản (không chứa object/drawing/math)
      const hasComplex = Array.from(child.children).some(
        (c) => c.localName !== "rPr" && c.localName !== "t" && c.localName !== "tab"
      );
      if (hasComplex) return;

      const tNodes = Array.from(child.children).filter((c) => c.localName === "t");
      if (tNodes.length === 0) return;

      const text = tNodes.map((t) => t.textContent || "").join("");
      if (!text) return;

      const rPr = Array.from(child.children).find((c) => c.localName === "rPr") || null;
      spans.push({
        runEl: child,
        rPr,
        text,
        start: fullText.length,
        end: fullText.length + text.length
      });
      fullText += text;
    });

    return { spans, fullText };
  }

  function createPreservedTextRun(xmlDoc, text, templateRPr, highlightError) {
    const wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const run = xmlDoc.createElementNS(wNs, "w:r");
    const rPr = templateRPr
      ? templateRPr.cloneNode(true)
      : highlightError
        ? xmlDoc.createElementNS(wNs, "w:rPr")
        : null;

    if (highlightError && rPr) {
      const color = xmlDoc.createElementNS(wNs, "w:color");
      color.setAttributeNS(wNs, "w:val", "DC2626");
      const bold = xmlDoc.createElementNS(wNs, "w:b");
      const hl = xmlDoc.createElementNS(wNs, "w:highlight");
      hl.setAttributeNS(wNs, "w:val", "yellow");
      rPr.appendChild(color);
      rPr.appendChild(bold);
      rPr.appendChild(hl);
    }

    if (rPr && rPr.childNodes.length > 0) {
      run.appendChild(rPr);
    }
    const tEl = xmlDoc.createElementNS(wNs, "w:t");
    tEl.setAttribute("xml:space", "preserve");
    tEl.textContent = text;
    run.appendChild(tEl);
    return run;
  }

  /**
   * Module độc lập: Chuyển đổi hàng loạt LaTeX ($...$, $$...$$, \(...\), \[...\])
   * trong vùng chọn hoặc toàn tài liệu sang đối tượng MathType 7.0 nhúng (Equation.DSMT4)
   * hoàn toàn bằng JavaScript thuần (không cần cài phần mềm MathType trên máy):
   * 1. Tạo ảnh vector EMF sắc nét (media/image1.emf, image/x-emf) bằng Uint8Array/DataView.
   * 2. Tự động căn chỉnh đường cơ sở dòng chữ qua <w:rPr><w:position w:val="${-Math.round(depthPt*2)}"/></w:rPr>.
   * 3. Mã hóa MTEF v5 (header 5,1,0,7,0) -> Equation Native (>= 4096 bytes FAT) -> OLE2 CFBF
   *    với CLSID chuẩn Design Science/Wiris (02 CE 02 00...) và \x01CompObj ("MathType 7.0 Equation").
   * 4. Định vị từng công thức bằng searchExactInParagraph (matchWildcards: false) và
   *    bọc item.insertOoxml(..., replace) + context.sync() trong try...catch riêng biệt.
   */
  async function convertLatexToMathType(options) {
    if (!ensureWordAvailable() || state.isBusy) return null;

    const requestedScope = (options && options.scope) || state.scanScope || "selection";
    const fontSizePt = Number(
      (options && options.fontSizePt) ||
        (dom.optMathTypeFontSize && dom.optMathTypeFontSize.value) ||
        12
    );
    const shouldHighlightError =
      options && typeof options.highlightError === "boolean"
        ? options.highlightError
        : dom.optHighlightError
          ? dom.optHighlightError.checked
          : true;

    const extractOpts = getScanOptions();
    if (!extractOpts.useInline && !extractOpts.useDisplay) {
      showStatus("Vui lòng bật ít nhất một cú pháp nhận diện LaTeX.", "warn");
      return null;
    }

    state.cancelRequested = false;
    setBusy(true, true);
    renderErrorReport([]);

    let totalFound = 0;
    let processedCount = 0;
    let convertedCount = 0;
    let errorCount = 0;
    let autoSwitched = false;
    let lastTechError = "";
    const errorItems = [];

    try {
      await Word.run(async (context) => {
        const target = await resolveTargetContainer(context, requestedScope);
        autoSwitched = target.autoSwitchedToBody;

        const paragraphs = target.container.paragraphs;
        paragraphs.load("items/text");
        await context.sync();

        // Duyệt từng đoạn văn và trích xuất công thức LaTeX bằng Regex chuẩn JS
        const tasks = [];
        paragraphs.items.forEach((para, paraIdx) => {
          const matches = extractLatexMatches(para.text || "", extractOpts);
          if (matches.length === 0) return;

          // Gom theo chuỗi raw duy nhất trong cùng 1 đoạn để tránh tìm lặp
          const grouped = new Map();
          matches.forEach((m) => {
            if (!grouped.has(m.raw)) {
              grouped.set(m.raw, { match: m, count: 1 });
            } else {
              grouped.get(m.raw).count += 1;
            }
          });

          grouped.forEach((entry) => {
            totalFound += entry.count;
            tasks.push({
              para,
              paragraphIndex: paraIdx + 1,
              match: entry.match,
              count: entry.count
            });
          });
        });

        if (totalFound === 0) {
          updateProgressUI(0, 0, 0, 0);
          return;
        }

        updateProgressUI(
          0,
          totalFound,
          0,
          0,
          "Đang tạo MathType 0/" + totalFound + " công thức..."
        );
        await yieldToUi();

        let seqNum = 0;
        const baseObjId = 1700000000 + Math.floor(Math.random() * 900000);

        for (let t = 0; t < tasks.length; t++) {
          if (state.cancelRequested) break;
          const { para, paragraphIndex, match, count } = tasks[t];

          // Bước 1 & 2: Phân tích LaTeX -> MTEF v5 (MathType 7.0) -> Equation Native (>=4096B) -> OLE CFBF + ảnh vector EMF
          let ast = null;
          let oleBase64 = "";
          let emfPreview = null;
          try {
            if (typeof temml !== "undefined") {
              try {
                renderLatexToMathML(match.latex, match.displayMode);
              } catch (temmlSyntaxErr) {
                if (!hasBalancedLatexBraces(match.latex)) {
                  throw temmlSyntaxErr;
                }
              }
            }

            ast = parseLatexToAst(match.latex);
            const mtefBytes = encodeAstToMtefV5(ast, {
              displayMode: match.displayMode,
              fontSizePt
            });
            const nativeBytes = wrapMtefInEquationNative(mtefBytes);
            const oleCfbBytes = buildOleCfbWithEquationNative(nativeBytes);
            oleBase64 = uint8ArrayToBase64(oleCfbBytes);
            emfPreview = renderAstToEmf(ast, fontSizePt, match.displayMode);
          } catch (buildErr) {
            const errMsg = formatWordApiError(buildErr);
            lastTechError = errMsg;
            if (shouldHighlightError) {
              await markSyntaxErrorInPlace(context, para, match.raw);
            }
            errorCount += count;
            processedCount += count;
            errorItems.push({
              paragraphIndex,
              raw: match.raw,
              latex: match.latex,
              displayMode: match.displayMode,
              message: "Lỗi tạo dữ liệu MathType 7.0 / EMF: " + errMsg
            });
            updateProgressUI(
              processedCount,
              totalFound,
              convertedCount,
              errorCount,
              "Đang tạo MathType " + processedCount + "/" + totalFound + " công thức..."
            );
            await yieldToUi();
            continue;
          }

          // Bước 3 & 4: Định vị bằng searchExactInParagraph (matchWildcards: false) và chèn OOXML từng công thức
          let searchResults = { items: [] };
          try {
            searchResults = await searchExactInParagraph(context, para, match.raw);
          } catch (searchErr) {
            const errMsg = formatWordApiError(searchErr);
            lastTechError = errMsg;
            errorCount += count;
            processedCount += count;
            errorItems.push({
              paragraphIndex,
              raw: match.raw,
              latex: match.latex,
              displayMode: match.displayMode,
              message: "Lỗi tìm vị trí công thức trong Word: " + errMsg
            });
            updateProgressUI(
              processedCount,
              totalFound,
              convertedCount,
              errorCount,
              "Đang tạo MathType " + processedCount + "/" + totalFound + " công thức..."
            );
            await yieldToUi();
            continue;
          }

          if (!searchResults.items || searchResults.items.length === 0) {
            const notFoundMsg =
              "Word không định vị được chuỗi \"" +
              match.raw +
              "\" trong đoạn " +
              paragraphIndex +
              " (có thể công thức dài > 250 ký tự hoặc bị ngắt định dạng).";
            lastTechError = notFoundMsg;
            errorCount += count;
            processedCount += count;
            errorItems.push({
              paragraphIndex,
              raw: match.raw,
              latex: match.latex,
              displayMode: match.displayMode,
              message: notFoundMsg
            });
            updateProgressUI(
              processedCount,
              totalFound,
              convertedCount,
              errorCount,
              "Đang tạo MathType " + processedCount + "/" + totalFound + " công thức..."
            );
            await yieldToUi();
            continue;
          }

          // Chèn và đồng bộ từng vị trí với try...catch độc lập để bắt chính xác lỗi nếu Word từ chối OOXML
          for (let k = 0; k < searchResults.items.length; k++) {
            if (state.cancelRequested) break;
            const item = searchResults.items[k];
            seqNum += 1;

            const shapeId = "_x0000_i" + (1024 + seqNum);
            const objectId = "_" + (baseObjId + seqNum);

            try {
              const ooxml = createMathTypeOoxml({
                emfBase64: emfPreview.emfBase64,
                oleBase64,
                widthPt: emfPreview.widthPt,
                heightPt: emfPreview.heightPt,
                depthPt: emfPreview.depthPt,
                shapeId,
                objectId,
                fontSizePt,
                displayMode: match.displayMode
              });

              item.insertOoxml(ooxml, Word.InsertLocation.replace);
              await context.sync();
              convertedCount += 1;
            } catch (insertErr) {
              const techErr = formatWordApiError(insertErr);
              lastTechError = techErr;
              errorCount += 1;
              if (shouldHighlightError) {
                try {
                  item.font.color = "#DC2626";
                  item.font.bold = true;
                  await context.sync();
                } catch {
                  // Bỏ qua nếu không thể tô màu trực tiếp trên range lỗi
                }
              }
              errorItems.push({
                paragraphIndex,
                raw: match.raw,
                latex: match.latex,
                displayMode: match.displayMode,
                message: "Word từ chối chèn OOXML MathType: " + techErr
              });
              showStatus("Lỗi chèn MathType (" + match.raw + "): " + techErr, "error");
            }

            processedCount += 1;
            updateProgressUI(
              processedCount,
              totalFound,
              convertedCount,
              errorCount,
              "Đang tạo MathType " + processedCount + "/" + totalFound + " công thức..."
            );
            await yieldToUi();
          }
        }
      });

      renderErrorReport(errorItems);
      const prefix = autoSwitched ? "Tự động áp dụng cho toàn tài liệu. " : "";

      if (totalFound === 0) {
        showStatus(
          prefix + "Không tìm thấy công thức LaTeX ($...$, $$...$$) nào để đổi sang MathType.",
          "warn"
        );
      } else if (state.cancelRequested) {
        updateProgressUI(
          processedCount,
          totalFound,
          convertedCount,
          errorCount,
          "Đã dừng tại " + processedCount + "/" + totalFound + " công thức"
        );
        showStatus(
          prefix +
            "Đã dừng chuyển đổi MathType. Thành công " +
            convertedCount +
            " công thức" +
            (errorCount > 0 ? ", lỗi " + errorCount + " công thức." : "."),
          "warn"
        );
      } else if (convertedCount === 0 && errorCount > 0) {
        updateProgressUI(
          totalFound,
          totalFound,
          0,
          errorCount,
          "Thất bại 0/" + totalFound + " công thức (" + errorCount + " lỗi)"
        );
        showStatus(
          prefix +
            "Không thể chèn đối tượng MathType (0/" +
            totalFound +
            " thành công). Lỗi kỹ thuật: " +
            lastTechError,
          "error"
        );
      } else {
        updateProgressUI(
          totalFound,
          totalFound,
          convertedCount,
          errorCount,
          "Hoàn tất tạo " + convertedCount + "/" + totalFound + " đối tượng MathType 7.0 (" + fontSizePt + "pt · EMF)"
        );
        showStatus(
          prefix +
            "Đã chuyển " +
            convertedCount +
            "/" +
            totalFound +
            " công thức sang MathType 7.0 (Equation.DSMT4 · EMF · " +
            fontSizePt +
            "pt)." +
            (errorCount > 0 ? " Có " + errorCount + " lỗi: " + lastTechError : ""),
          errorCount > 0 ? "warn" : "success"
        );
      }

      return { totalFound, convertedCount, errorCount, errorItems };
    } catch (err) {
      const techErr = formatWordApiError(err);
      updateProgressUI(0, 0, 0, 0);
      showStatus("Lỗi chuyển LaTeX sang MathType: " + techErr, "error");
      return null;
    } finally {
      setBusy(false, false);
      state.cancelRequested = false;
    }
  }

  // Xuất module độc lập lên window để có thể gọi trực tiếp từ bên ngoài nếu cần
  window.convertLatexToMathType = convertLatexToMathType;

  /* ------------------------------------------------------------------------
   * 10. KHỞI CHẠY OFFICE.ONREADY
   * ------------------------------------------------------------------------ */
  if (typeof window.Office !== "undefined" && typeof window.Office.onReady === "function") {
    window.Office.onReady((info) => {
      const isWord = info && info.host === window.Office.HostType.Word;
      initApp(isWord);
    });
  } else {
    document.addEventListener("DOMContentLoaded", () => initApp(false));
  }
})();


