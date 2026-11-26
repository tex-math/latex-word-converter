Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        const inputEl = document.getElementById("latex-input");
        inputEl.addEventListener("input", updatePreview);
        updatePreview();

        document.getElementById("btn-insert-latex").onclick = insertSingleLatex;
        document.getElementById("btn-batch-latex-to-word").onclick = convertAllLatexToWord;
        document.getElementById("btn-word-to-latex").onclick = convertWordEquationsToLatex;
    }
});

// Xem trước công thức bằng Temml
function updatePreview() {
    const latex = document.getElementById("latex-input").value.trim();
    const previewBox = document.getElementById("preview-box");
    if (!latex) {
        previewBox.innerHTML = "<span style='color:#999'>Xem trước công thức...</span>";
        return;
    }
    try {
        temml.render(latex, previewBox, { throwOnError: false });
    } catch (e) {
        previewBox.innerText = "Lỗi cú pháp LaTeX";
    }
}

// Tạo gói OOXML từ chuỗi LaTeX
function createEquationOoxml(latexStr) {
    const mathml = temml.renderToString(latexStr, { xml: true });
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
        <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
            <pkg:xmlData>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
            </pkg:xmlData>
        </pkg:part>
        <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
            <pkg:xmlData>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                    <w:body>
                        <w:p>
                            <w:r>
                                ${mathml}
                            </w:r>
                        </w:p>
                    </w:body>
                </w:document>
            </pkg:xmlData>
        </pkg:part>
    </pkg:package>`;
}

// 1. Chèn 1 công thức từ ô nhập liệu
async function insertSingleLatex() {
    const latex = document.getElementById("latex-input").value.trim();
    if (!latex) return;

    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const ooxml = createEquationOoxml(latex);
        selection.insertOoxml(ooxml, Word.InsertLocation.replace);
        await context.sync();
        setStatus("Đã chèn công thức Word Equation!");
    }).catch((err) => setStatus("Lỗi: " + err.message, true));
}

// 2. Quét và chuyển đổi hàng loạt $...$ -> Word Equation
async function convertAllLatexToWord() {
    await Word.run(async (context) => {
        let range = context.document.getSelection();
        range.load("text");
        await context.sync();

        if (!range.text || range.text.trim() === "") {
            range = context.document.body.getRange();
        }

        const searchResults = range.search("\\$*\\$", { matchWildcards: true });
        searchResults.load("items");
        await context.sync();

        let count = 0;
        for (let i = 0; i < searchResults.items.length; i++) {
            const item = searchResults.items[i];
            item.load("text");
            await context.sync();

            let rawText = item.text;
            let cleanLatex = rawText.replace(/^\$\$?|\$\$?$/g, "").trim();

            if (cleanLatex) {
                const ooxml = createEquationOoxml(cleanLatex);
                item.insertOoxml(ooxml, Word.InsertLocation.replace);
                count++;
            }
        }
        await context.sync();
        setStatus(`Hoàn tất! Đã chuyển ${count} chuỗi LaTeX sang Word Equation.`);
    }).catch((err) => setStatus("Lỗi: " + err.message, true));
}

// 3. Chuyển đổi ngược Word Equation (OMML) -> chuỗi $...$
async function convertWordEquationsToLatex() {
    await Word.run(async (context) => {
        let range = context.document.getSelection();
        range.load("text");
        await context.sync();

        if (!range.text || range.text.trim() === "") {
            range = context.document.body.getRange();
        }

        const ooxmlResult = range.getOoxml();
        await context.sync();

        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(ooxmlResult.value, "text/xml");
        const mathNodes = xmlDoc.getElementsByTagName("m:oMath");

        if (mathNodes.length === 0) {
            setStatus("Không tìm thấy công thức Word Equation nào trong vùng chọn.", true);
            return;
        }

        let count = mathNodes.length;
        // Duyệt ngược để thay thế node XML không làm lệch vị trí
        for (let i = mathNodes.length - 1; i >= 0; i--) {
            const oMath = mathNodes[i];
            const latex = ommlToLatex(oMath);
            
            const wRun = xmlDoc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:r");
            const wText = xmlDoc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:t");
            wText.textContent = `$${latex}$`;
            wRun.appendChild(wText);

            // Nếu m:oMath nằm trong m:oMathPara thì thay thế cả m:oMathPara
            const parent = oMath.parentNode;
            if (parent && parent.nodeName === "m:oMathPara") {
                parent.parentNode.replaceChild(wRun, parent);
            } else if (parent) {
                parent.replaceChild(wRun, oMath);
            }
        }

        const serializer = new XMLSerializer();
        const updatedOoxml = serializer.serializeToString(xmlDoc);
        range.insertOoxml(updatedOoxml, Word.InsertLocation.replace);
        await context.sync();

        setStatus(`Đã chuyển ${count} công thức Word Equation sang dạng $...$.`);
    }).catch((err) => setStatus("Lỗi: " + err.message, true));
}

// Bộ dịch cơ bản từ cây XML OMML của Word sang cú pháp LaTeX
function ommlToLatex(node) {
    if (!node) return "";
    let result = "";

    for (let i = 0; i < node.childNodes.length; i++) {
        const child = node.childNodes[i];
        const name = child.nodeName;

        if (name === "m:t") {
            result += child.textContent
                .replace(/×/g, "\\times ")
                .replace(/÷/g, "\\div ")
                .replace(/±/g, "\\pm ")
                .replace(/≤/g, "\\le ")
                .replace(/≥/g, "\\ge ")
                .replace(/≠/g, "\\ne ")
                .replace(/≈/g, "\\approx ")
                .replace(/∞/g, "\\infty ")
                .replace(/π/g, "\\pi ")
                .replace(/α/g, "\\alpha ")
                .replace(/β/g, "\\beta ")
                .replace(/γ/g, "\\gamma ")
                .replace(/Δ/g, "\\Delta ")
                .replace(/θ/g, "\\theta ")
                .replace(/⇔/g, "\\Leftrightarrow ")
                .replace(/⇒/g, "\\Rightarrow ");
        } else if (name === "m:f") {
            // Phân số
            const num = child.getElementsByTagName("m:num")[0];
            const den = child.getElementsByTagName("m:den")[0];
            result += `\\frac{${ommlToLatex(num)}}{${ommlToLatex(den)}}`;
        } else if (name === "m:sSup") {
            // Lũy thừa / chỉ số trên
            const e = child.getElementsByTagName("m:e")[0];
            const sup = child.getElementsByTagName("m:sup")[0];
            result += `{${ommlToLatex(e)}}^{${ommlToLatex(sup)}}`;
        } else if (name === "m:sSub") {
            // Chỉ số dưới
            const e = child.getElementsByTagName("m:e")[0];
            const sub = child.getElementsByTagName("m:sub")[0];
            result += `{${ommlToLatex(e)}}_{${ommlToLatex(sub)}}`;
        } else if (name === "m:sSubSup") {
            // Cả chỉ số dưới và trên
            const e = child.getElementsByTagName("m:e")[0];
            const sub = child.getElementsByTagName("m:sub")[0];
            const sup = child.getElementsByTagName("m:sup")[0];
            result += `{${ommlToLatex(e)}}_{${ommlToLatex(sub)}}^{${ommlToLatex(sup)}}`;
        } else if (name === "m:rad") {
            // Căn thức
            const deg = child.getElementsByTagName("m:deg")[0];
            const e = child.getElementsByTagName("m:e")[0];
            const degText = deg ? ommlToLatex(deg).trim() : "";
            if (degText) {
                result += `\\sqrt[${degText}]{${ommlToLatex(e)}}`;
            } else {
                result += `\\sqrt{${ommlToLatex(e)}}`;
            }
        } else if (name === "m:nary") {
            // Tích phân, tổng sigma
            const chrNode = child.getElementsByTagName("m:chr")[0];
            const chr = chrNode ? chrNode.getAttribute("m:val") : "∫";
            const sub = child.getElementsByTagName("m:sub")[0];
            const sup = child.getElementsByTagName("m:sup")[0];
            const e = child.getElementsByTagName("m:e")[0];
            let op = "\\int";
            if (chr === "∑") op = "\\sum";
            else if (chr === "∏") op = "\\prod";
            const subStr = sub && ommlToLatex(sub).trim() ? `_{${ommlToLatex(sub)}}` : "";
            const supStr = sup && ommlToLatex(sup).trim() ? `^{${ommlToLatex(sup)}}` : "";
            result += `${op}${subStr}${supStr} ${ommlToLatex(e)}`;
        } else if (name === "m:d") {
            // Dấu ngoặc bao quanh
            const e = child.getElementsByTagName("m:e")[0];
            result += `\\left(${ommlToLatex(e)}\\right)`;
        } else {
            result += ommlToLatex(child);
        }
    }
    return result;
}

function setStatus(msg, isError = false) {
    const el = document.getElementById("status");
    if (isError) {
        el.style.background = "#fde7e9";
        el.style.color = "#a4262c";
    } else {
        el.style.background = "#eef6ee";
        el.style.color = "#107c10";
    }
    el.innerText = msg;
}
