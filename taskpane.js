/* ==========================================================================
 * Chuyển đổi LaTeX Word — Microsoft Word Web Add-in (Vanilla ES6+)
 * Host tĩnh trên GitHub Pages: https://tex-math.github.io/latex-word-converter/
 * Chức năng:
 *  - Chuyển đổi LaTeX -> MathML (Temml) -> Word OMML (OOXML chuẩn Equation Word)
 *  - Chèn công thức tại con trỏ Word bằng Word.run + insertOoxml
 *  - Quét & chuyển đổi hàng loạt $...$, $$...$$, \(...\), \[...\]
 *  - Chuyển đổi ngược Microsoft Word Equation (OMML) -> LaTeX
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

  // Các ký tự toán tử N-ary lớn trong Word OMML
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
    displayMode: false,
    isBusy: false,
    lastMathML: "",
    lastScanItems: []
  };

  const dom = {};

  /* ------------------------------------------------------------------------
   * 1. KHỞI TẠO OFFICE.JS & GIAO DIỆN
   * ------------------------------------------------------------------------ */
  function initApp(isWord) {
    state.isWordHost = Boolean(isWord);
    cacheDom();
    bindEvents();
    renderTemplates();
    renderHistory();
    updateHostIndicator();
    updateLivePreview();
  }

  function cacheDom() {
    dom.hostIndicator = document.getElementById("hostIndicator");
    dom.tabBtns = Array.from(document.querySelectorAll(".tab-btn"));
    dom.tabPanels = {
      editor: document.getElementById("panel-editor"),
      batch: document.getElementById("panel-batch"),
      library: document.getElementById("panel-library")
    };
    dom.modeBtns = Array.from(document.querySelectorAll(".mode-btn"));
    dom.symbolToolbar = document.getElementById("symbolToolbar");
    dom.latexInput = document.getElementById("latexInput");
    dom.btnClearInput = document.getElementById("btnClearInput");
    dom.btnCopyMathml = document.getElementById("btnCopyMathml");
    dom.mathPreview = document.getElementById("mathPreview");
    dom.btnInsertWord = document.getElementById("btnInsertWord");
    dom.btnReadFromWord = document.getElementById("btnReadFromWord");
    dom.btnWrapInsertLatex = document.getElementById("btnWrapInsertLatex");

    dom.optDollar = document.getElementById("optDollar");
    dom.optBracket = document.getElementById("optBracket");
    dom.optSkipCurrency = document.getElementById("optSkipCurrency");
    dom.btnScanPreview = document.getElementById("btnScanPreview");
    dom.btnBatchConvert = document.getElementById("btnBatchConvert");
    dom.btnBatchReverse = document.getElementById("btnBatchReverse");
    dom.batchProgressWrap = document.getElementById("batchProgressWrap");
    dom.batchProgressFill = document.getElementById("batchProgressFill");
    dom.batchProgressText = document.getElementById("batchProgressText");
    dom.scanCountLabel = document.getElementById("scanCountLabel");
    dom.scanResultsList = document.getElementById("scanResultsList");

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

  function bindEvents() {
    // Chuyển tab
    dom.tabBtns.forEach((btn) => {
      btn.addEventListener("click", () => switchTab(btn.dataset.tab));
    });

    // Chuyển chế độ inline / display
    dom.modeBtns.forEach((btn) => {
      btn.addEventListener("click", () => {
        state.displayMode = btn.dataset.mode === "display";
        dom.modeBtns.forEach((b) => b.classList.toggle("active", b === btn));
        updateLivePreview();
      });
    });

    // Bảng chèn nhanh ký hiệu
    dom.symbolToolbar.addEventListener("click", (event) => {
      const target = event.target.closest(".sym-btn");
      if (!target) return;
      insertSnippetAtCursor(target.dataset.snippet || "");
    });

    // Nhập LaTeX + phím tắt Ctrl+Enter
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

    // Tab quét hàng loạt
    dom.btnScanPreview.addEventListener("click", handleScanPreview);
    dom.btnBatchConvert.addEventListener("click", handleBatchConvert);
    dom.btnBatchReverse.addEventListener("click", handleBatchReverse);

    // Tab lịch sử
    dom.btnClearHistory.addEventListener("click", () => {
      saveHistory([]);
      renderHistory();
      showStatus("Đã xóa toàn bộ lịch sử công thức.", "success");
    });
  }

  function switchTab(tabId) {
    dom.tabBtns.forEach((btn) => {
      const isActive = btn.dataset.tab === tabId;
      btn.classList.toggle("active", isActive);
      btn.setAttribute("aria-selected", String(isActive));
    });
    Object.keys(dom.tabPanels).forEach((key) => {
      dom.tabPanels[key].classList.toggle("active", key === tabId);
    });
  }

  function insertSnippetAtCursor(snippet) {
    const input = dom.latexInput;
    const start = input.selectionStart ?? input.value.length;
    const end = input.selectionEnd ?? input.value.length;
    const before = input.value.slice(0, start);
    const after = input.value.slice(end);
    input.value = before + snippet + after;
    const cursorPos = start + snippet.length;
    input.setSelectionRange(cursorPos, cursorPos);
    input.focus();
    updateLivePreview();
  }

  function showStatus(message, type = "info") {
    dom.statusBar.textContent = message;
    dom.statusBar.className = "status-bar visible" + (type !== "info" ? " " + type : "");
  }

  function setBusy(busy) {
    state.isBusy = busy;
    [
      dom.btnInsertWord,
      dom.btnReadFromWord,
      dom.btnWrapInsertLatex,
      dom.btnScanPreview,
      dom.btnBatchConvert,
      dom.btnBatchReverse
    ].forEach((btn) => {
      if (btn) btn.disabled = busy;
    });
  }

  function setProgress(current, total, label) {
    if (!total) {
      dom.batchProgressWrap.classList.remove("visible");
      return;
    }
    dom.batchProgressWrap.classList.add("visible");
    const pct = Math.min(100, Math.round((current / total) * 100));
    dom.batchProgressFill.style.width = pct + "%";
    dom.batchProgressText.textContent = label || ("Đang xử lý " + current + "/" + total + " công thức...");
  }

  /* ------------------------------------------------------------------------
   * 2. XEM TRƯỚC TRỰC TIẾP VỚI TEMML (LaTeX -> MathML)
   * ------------------------------------------------------------------------ */
  function renderLatexToMathML(latex, displayMode) {
    if (typeof window.temml === "undefined") {
      throw new Error("Thư viện Temml chưa tải xong. Vui lòng kiểm tra kết nối mạng.");
    }
    const clean = stripOuterDelimiters(latex);
    return window.temml.renderToString(clean, {
      displayMode: Boolean(displayMode),
      xml: true,
      throwOnError: true,
      annotate: false
    });
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
      dom.mathPreview.innerHTML = '<span class="preview-placeholder">Nhập mã LaTeX để xem trước công thức...</span>';
      return;
    }

    try {
      const mathml = renderLatexToMathML(raw, state.displayMode);
      state.lastMathML = mathml;
      dom.mathPreview.innerHTML = mathml;
    } catch (err) {
      state.lastMathML = "";
      dom.mathPreview.innerHTML = "";
      const errSpan = document.createElement("span");
      errSpan.className = "preview-error";
      errSpan.textContent = "Lỗi cú pháp LaTeX: " + (err && err.message ? err.message : String(err));
      dom.mathPreview.appendChild(errSpan);
    }
  }

  /* ------------------------------------------------------------------------
   * 3. CHUYỂN ĐỔI MATHML -> WORD OMML & ĐÓNG GÓI FLAT OPC OOXML
   * ------------------------------------------------------------------------ */
  function escapeXml(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function createOmmlRun(text, isNormal = false) {
    if (!text) return "";
    const rPr = isNormal ? "<m:rPr><m:nor/></m:rPr>" : "";
    const wRPr = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
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
    if (tag === "mo") {
      return NARY_CHARS.has(getDirectText(node));
    }
    if (tag === "mrow" && node.children.length === 1) {
      return isNaryOperatorNode(node.children[0]);
    }
    return false;
  }

  function getNaryChar(node) {
    return getDirectText(node);
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

      // Xử lý toán tử N-ary (∫, ∑, ∏...) kèm biểu thức liền sau thành <m:nary>
      const isNaryWrapper =
        (tag === "msubsup" || tag === "munderover" || tag === "msub" || tag === "munder") &&
        node.children.length >= 2 &&
        isNaryOperatorNode(node.children[0]);

      if (isNaryWrapper) {
        const chr = getNaryChar(node.children[0]);
        const isUndOvr = tag === "munderover" || tag === "munder";
        const hasSup = tag === "msubsup" || tag === "munderover";
        const subXml = convertMathNode(node.children[1]);
        const supXml = hasSup && node.children[2] ? convertMathNode(node.children[2]) : "";

        // Gom các phần tử liền sau cho đến khi gặp dấu quan hệ (=, <, >...) để làm <m:e>
        const exprNodes = [];
        let j = i + 1;
        while (j < children.length && !isBoundaryRelation(children[j])) {
          exprNodes.push(children[j]);
          j++;
        }
        const exprXml = exprNodes.length > 0 ? convertChildrenList(exprNodes) : createOmmlRun(" ", true);

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

    const tag = node.localName;
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
        // Nhận diện cặp dấu ngoặc giãn nở (\left(...\right), \begin{cases}, ma trận...)
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
        const isNormal = variant === "normal" || text.length > 1;
        return createOmmlRun(text, isNormal);
      }

      case "mn":
      case "mtext":
        return createOmmlRun(node.textContent || "", true);

      case "mo": {
        const text = node.textContent || "";
        if (!text) return "";
        return createOmmlRun(text, true);
      }

      case "mspace":
        return createOmmlRun(" ", true);

      case "mfrac": {
        if (kids.length < 2) return convertChildrenList(kids);
        const lineThickness = node.getAttribute("linethickness");
        const fPr = lineThickness === "0" || lineThickness === "0px"
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
        return (
          "<m:sSup>" +
            "<m:e>" + convertMathNode(kids[0]) + "</m:e>" +
            "<m:sup>" + convertMathNode(kids[1]) + "</m:sup>" +
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
        return (
          "<m:sSub>" +
            "<m:e>" + convertMathNode(kids[0]) + "</m:e>" +
            "<m:sub>" + convertMathNode(kids[1]) + "</m:sub>" +
          "</m:sSub>"
        );
      }

      case "msubsup": {
        if (kids.length < 3) return convertChildrenList(kids);
        return (
          "<m:sSubSup>" +
            "<m:e>" + convertMathNode(kids[0]) + "</m:e>" +
            "<m:sub>" + convertMathNode(kids[1]) + "</m:sub>" +
            "<m:sup>" + convertMathNode(kids[2]) + "</m:sup>" +
          "</m:sSubSup>"
        );
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
    const parser = new DOMParser();
    const doc = parser.parseFromString(mathmlString, "application/xml");
    const parseErr = doc.querySelector("parsererror");
    if (parseErr) {
      throw new Error("Không thể phân tích cú pháp MathML.");
    }
    const mathEl = doc.documentElement;
    return "<m:oMath>" + convertMathNode(mathEl) + "</m:oMath>";
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

  function latexToWordOoxml(latex, displayMode) {
    const mathml = renderLatexToMathML(latex, displayMode);
    const oMath = mathmlToOmml(mathml);
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
        const tNodes = getChildrenByLocalName(node, "t");
        const raw = tNodes.map((t) => t.textContent || "").join("");
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
        const isNoBar = typeEl && (typeEl.getAttribute("m:val") === "noBar" || typeEl.getAttribute("val") === "noBar");
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
        const e = ommlNodeToLatex(getChildByLocalName(node, "e")).trim();
        const sup = ommlNodeToLatex(getChildByLocalName(node, "sup")).trim();
        return e + "^{" + sup + "}";
      }

      case "sSub": {
        const e = ommlNodeToLatex(getChildByLocalName(node, "e")).trim();
        const sub = ommlNodeToLatex(getChildByLocalName(node, "sub")).trim();
        return e + "_{" + sub + "}";
      }

      case "sSubSup": {
        const e = ommlNodeToLatex(getChildByLocalName(node, "e")).trim();
        const sub = ommlNodeToLatex(getChildByLocalName(node, "sub")).trim();
        const sup = ommlNodeToLatex(getChildByLocalName(node, "sup")).trim();
        return e + "_{" + sub + "}^{" + sup + "}";
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

      case "limLow": {
        const e = ommlNodeToLatex(getChildByLocalName(node, "e")).trim();
        const lim = ommlNodeToLatex(getChildByLocalName(node, "lim")).trim();
        return e + "_{" + lim + "} ";
      }

      case "limUpp": {
        const e = ommlNodeToLatex(getChildByLocalName(node, "e")).trim();
        const lim = ommlNodeToLatex(getChildByLocalName(node, "lim")).trim();
        return e + "^{" + lim + "} ";
      }

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

        // Nhận diện hệ phương trình \begin{cases}...\end{cases}
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

  /* ------------------------------------------------------------------------
   * 5. BỘ TÁCH CÔNG THỨC LATEX TRONG VĂN BẢN ($...$, $$...$$, \(...\), \[...\])
   * ------------------------------------------------------------------------ */
  function extractLatexMatches(text, options) {
    const matches = [];
    if (!text) return matches;

    let i = 0;
    const len = text.length;

    while (i < len) {
      // Bỏ qua ký tự escape \$
      if (text[i] === "\\" && i + 1 < len && text[i + 1] === "$") {
        i += 2;
        continue;
      }

      // 1. Dạng $$...$$
      if (options.useDollar && text[i] === "$" && text[i + 1] === "$") {
        const closeIdx = text.indexOf("$$", i + 2);
        if (closeIdx !== -1) {
          const inner = text.slice(i + 2, closeIdx).trim();
          if (inner.length > 0) {
            matches.push({
              raw: text.slice(i, closeIdx + 2),
              latex: inner,
              displayMode: true
            });
            i = closeIdx + 2;
            continue;
          }
        }
      }

      // 2. Dạng $...$
      if (options.useDollar && text[i] === "$" && text[i + 1] !== "$") {
        if (i + 1 < len && !/\s/.test(text[i + 1])) {
          let closeIdx = -1;
          for (let j = i + 1; j < len; j++) {
            if (text[j] === "\\" && j + 1 < len) {
              j++;
              continue;
            }
            if (text[j] === "$") {
              closeIdx = j;
              break;
            }
          }
          if (closeIdx !== -1 && !/\s/.test(text[closeIdx - 1])) {
            const inner = text.slice(i + 1, closeIdx).trim();
            const looksLikeCurrency =
              options.skipCurrency && /^\d[\d.,]*$/.test(inner);

            if (inner.length > 0 && !looksLikeCurrency) {
              matches.push({
                raw: text.slice(i, closeIdx + 1),
                latex: inner,
                displayMode: false
              });
              i = closeIdx + 1;
              continue;
            }
          }
        }
      }

      // 3. Dạng \[...\]
      if (options.useBracket && text[i] === "\\" && text[i + 1] === "[") {
        const closeIdx = text.indexOf("\\]", i + 2);
        if (closeIdx !== -1) {
          const inner = text.slice(i + 2, closeIdx).trim();
          if (inner.length > 0) {
            matches.push({
              raw: text.slice(i, closeIdx + 2),
              latex: inner,
              displayMode: true
            });
            i = closeIdx + 2;
            continue;
          }
        }
      }

      // 4. Dạng \(...\)
      if (options.useBracket && text[i] === "\\" && text[i + 1] === "(") {
        const closeIdx = text.indexOf("\\)", i + 2);
        if (closeIdx !== -1) {
          const inner = text.slice(i + 2, closeIdx).trim();
          if (inner.length > 0) {
            matches.push({
              raw: text.slice(i, closeIdx + 2),
              latex: inner,
              displayMode: false
            });
            i = closeIdx + 2;
            continue;
          }
        }
      }

      i++;
    }

    return matches;
  }

  function getScanOptions() {
    const scopeEl = document.querySelector('input[name="scanScope"]:checked');
    return {
      scope: scopeEl ? scopeEl.value : "body",
      useDollar: dom.optDollar.checked,
      useBracket: dom.optBracket.checked,
      skipCurrency: dom.optSkipCurrency.checked
    };
  }

  /* ------------------------------------------------------------------------
   * 6. CÁC HÀM TƯƠNG TÁC MICROSOFT WORD (OFFICE.JS - Word.run)
   * ------------------------------------------------------------------------ */
  function ensureWordAvailable() {
    if (typeof window.Word === "undefined" || typeof window.Word.run !== "function") {
      showStatus("Không tìm thấy môi trường Word (Word.run). Hãy mở Add-in trong Microsoft Word.", "warn");
      return false;
    }
    return true;
  }

  async function handleInsertEquation(overrideLatex, overrideDisplay) {
    if (!ensureWordAvailable() || state.isBusy) return;

    const rawLatex = typeof overrideLatex === "string" ? overrideLatex : dom.latexInput.value;
    const latex = stripOuterDelimiters(rawLatex);
    const displayMode = typeof overrideDisplay === "boolean" ? overrideDisplay : state.displayMode;

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
      showStatus("Không thể chèn công thức: " + (err && err.message ? err.message : String(err)), "error");
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
      showStatus("Lỗi khi chèn văn bản: " + (err && err.message ? err.message : String(err)), "error");
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
      showStatus("Không thể đọc công thức từ Word: " + (err && err.message ? err.message : String(err)), "error");
    } finally {
      setBusy(false);
    }
  }

  async function handleScanPreview() {
    if (!ensureWordAvailable() || state.isBusy) return;
    const opts = getScanOptions();
    if (!opts.useDollar && !opts.useBracket) {
      showStatus("Hãy chọn ít nhất một kiểu cú pháp nhận diện LaTeX.", "warn");
      return;
    }

    setBusy(true);
    try {
      const found = [];
      await Word.run(async (context) => {
        const container = opts.scope === "selection"
          ? context.document.getSelection()
          : context.document.body;
        const paragraphs = container.paragraphs;
        paragraphs.load("items/text");
        await context.sync();

        paragraphs.items.forEach((para, idx) => {
          const matches = extractLatexMatches(para.text || "", opts);
          matches.forEach((m) => {
            found.push({
              paragraphIndex: idx + 1,
              raw: m.raw,
              latex: m.latex,
              displayMode: m.displayMode
            });
          });
        });
      });

      state.lastScanItems = found;
      renderScanResults(found);
      if (found.length === 0) {
        showStatus("Không tìm thấy công thức LaTeX nào trong phạm vi đã chọn.", "warn");
      } else {
        showStatus("Tìm thấy " + found.length + " công thức LaTeX sẵn sàng chuyển đổi.", "success");
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
    if (!opts.useDollar && !opts.useBracket) {
      showStatus("Hãy chọn ít nhất một kiểu cú pháp nhận diện LaTeX.", "warn");
      return;
    }

    setBusy(true);
    let convertedCount = 0;
    let errorCount = 0;
    let skippedLong = 0;

    try {
      await Word.run(async (context) => {
        const container = opts.scope === "selection"
          ? context.document.getSelection()
          : context.document.body;
        const paragraphs = container.paragraphs;
        paragraphs.load("items/text");
        await context.sync();

        const tasks = [];
        paragraphs.items.forEach((para) => {
          const matches = extractLatexMatches(para.text || "", opts);
          if (matches.length === 0) return;

          const uniqueMap = new Map();
          matches.forEach((m) => {
            if (!uniqueMap.has(m.raw)) {
              uniqueMap.set(m.raw, m);
            }
          });
          uniqueMap.forEach((match) => {
            tasks.push({ para, match });
          });
        });

        const total = tasks.length;
        if (total === 0) {
          setProgress(0, 0);
          return;
        }

        setProgress(0, total, "Bắt đầu chuyển đổi " + total + " cụm công thức...");

        for (let idx = 0; idx < total; idx++) {
          const { para, match } = tasks[idx];
          if (match.raw.length > WORD_SEARCH_MAX_LEN) {
            skippedLong++;
            setProgress(idx + 1, total);
            continue;
          }

          let ooxml = "";
          try {
            ooxml = latexToWordOoxml(match.latex, match.displayMode);
          } catch {
            errorCount++;
            setProgress(idx + 1, total);
            continue;
          }

          const results = para.search(match.raw, { matchCase: true });
          results.load("items");
          await context.sync();

          for (let k = 0; k < results.items.length; k++) {
            results.items[k].insertOoxml(ooxml, Word.InsertLocation.replace);
            convertedCount++;
          }
          await context.sync();
          setProgress(idx + 1, total, "Đã chuyển " + (idx + 1) + "/" + total + " cụm công thức...");
        }
      });

      setProgress(0, 0);
      if (convertedCount === 0 && errorCount === 0 && skippedLong === 0) {
        showStatus("Không tìm thấy công thức LaTeX nào để chuyển đổi.", "warn");
      } else {
        const parts = ["Đã chuyển thành công " + convertedCount + " công thức sang Word Equation."];
        if (errorCount > 0) parts.push(errorCount + " công thức sai cú pháp.");
        if (skippedLong > 0) parts.push(skippedLong + " công thức quá 250 ký tự (hãy chèn riêng lẻ).");
        showStatus(parts.join(" "), errorCount > 0 || skippedLong > 0 ? "warn" : "success");
      }
      renderScanResults([]);
    } catch (err) {
      setProgress(0, 0);
      showStatus("Lỗi chuyển đổi hàng loạt: " + (err && err.message ? err.message : String(err)), "error");
    } finally {
      setBusy(false);
    }
  }

  async function handleBatchReverse() {
    if (!ensureWordAvailable() || state.isBusy) return;
    const opts = getScanOptions();

    setBusy(true);
    try {
      let replacedCount = 0;

      await Word.run(async (context) => {
        const container = opts.scope === "selection"
          ? context.document.getSelection()
          : context.document.body;
        const ooxmlRes = container.getOoxml();
        await context.sync();

        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(ooxmlRes.value, "application/xml");
        const wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        const allNodes = Array.from(xmlDoc.getElementsByTagName("*"));
        const mathTargets = allNodes.filter(
          (el) =>
            el.localName === "oMathPara" ||
            (el.localName === "oMath" && (!el.parentElement || el.parentElement.localName !== "oMathPara"))
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
          replacedCount++;
        });

        if (replacedCount > 0) {
          const serializer = new XMLSerializer();
          const updatedOoxml = serializer.serializeToString(xmlDoc);
          container.insertOoxml(updatedOoxml, Word.InsertLocation.replace);
          await context.sync();
        }
      });

      if (replacedCount === 0) {
        showStatus("Không tìm thấy Word Equation nào trong phạm vi đã chọn.", "warn");
      } else {
        showStatus("Đã chuyển ngược " + replacedCount + " Word Equation về mã LaTeX.", "success");
      }
    } catch (err) {
      showStatus("Lỗi khi chuyển ngược Equation: " + (err && err.message ? err.message : String(err)), "error");
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
   * 7. HIỂN THỊ KẾT QUẢ QUÉT, MẪU CÔNG THỨC & LỊCH SỬ
   * ------------------------------------------------------------------------ */
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
      card.className = "list-item";

      const head = document.createElement("div");
      head.className = "list-item-head";

      const title = document.createElement("span");
      title.className = "list-item-title";
      title.textContent = "#" + (index + 1) + " · Đoạn " + item.paragraphIndex;

      const meta = document.createElement("span");
      meta.className = "list-item-meta";
      meta.textContent = item.displayMode ? "Dòng riêng" : "Cùng dòng";

      head.appendChild(title);
      head.appendChild(meta);

      const code = document.createElement("div");
      code.className = "list-item-code";
      code.textContent = item.raw;
      code.title = item.raw;

      const actions = document.createElement("div");
      actions.className = "list-item-actions";
      const btnEdit = document.createElement("button");
      btnEdit.type = "button";
      btnEdit.className = "mini-btn";
      btnEdit.textContent = "Mở trong trình soạn";
      btnEdit.addEventListener("click", () => {
        dom.latexInput.value = item.latex;
        state.displayMode = item.displayMode;
        dom.modeBtns.forEach((b) =>
          b.classList.toggle("active", (b.dataset.mode === "display") === state.displayMode)
        );
        updateLivePreview();
        switchTab("editor");
      });
      actions.appendChild(btnEdit);

      card.appendChild(head);
      card.appendChild(code);
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
      meta.textContent = tpl.display ? "Dòng riêng" : "Cùng dòng";

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
      btnLoad.addEventListener("click", () => {
        dom.latexInput.value = tpl.latex;
        state.displayMode = tpl.display;
        dom.modeBtns.forEach((b) =>
          b.classList.toggle("active", (b.dataset.mode === "display") === state.displayMode)
        );
        updateLivePreview();
        switchTab("editor");
      });

      const btnQuickInsert = document.createElement("button");
      btnQuickInsert.type = "button";
      btnQuickInsert.className = "mini-btn accent";
      btnQuickInsert.textContent = "Chèn ngay";
      btnQuickInsert.addEventListener("click", () => {
        handleInsertEquation(tpl.latex, tpl.display);
      });

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
      btnLoad.addEventListener("click", () => {
        dom.latexInput.value = item.latex;
        state.displayMode = Boolean(item.displayMode);
        dom.modeBtns.forEach((b) =>
          b.classList.toggle("active", (b.dataset.mode === "display") === state.displayMode)
        );
        updateLivePreview();
        switchTab("editor");
      });

      const btnInsert = document.createElement("button");
      btnInsert.type = "button";
      btnInsert.className = "mini-btn accent";
      btnInsert.textContent = "Chèn lại";
      btnInsert.addEventListener("click", () => {
        handleInsertEquation(item.latex, Boolean(item.displayMode));
      });

      actions.appendChild(btnLoad);
      actions.appendChild(btnInsert);
      card.appendChild(code);
      card.appendChild(actions);
      dom.historyList.appendChild(card);
    });
  }

  /* ------------------------------------------------------------------------
   * 8. ĐIỂM KHỞI CHẠY OFFICE.ONREADY
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
