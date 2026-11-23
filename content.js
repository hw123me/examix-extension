(() => {
  "use strict";

  const getTargetTable = () =>
    document.getElementById("ctl00_PlaceHolderMain_gvDynamicStudentsMark") ||
    document.querySelector("table.GridClass") ||
    document.querySelector("table[id*='gvDynamicStudentsMark']");

  const TOOLBAR_ID = "examix-noor-toolbar";
  const LAUNCHER_ID = "examix-noor-launcher";
  const PASTE_BUTTON_ID = "examix-noor-paste-btn";
  const MAX_FILE_SIZE = 25 * 1024 * 1024;

  const STORAGE_KEY_DATA = "examix_import_source_list";
  const STORAGE_KEY_ACTIVE = "examix_import_is_active";
  const STORAGE_KEY_STATS = "examix_import_total_filled";
  const STORAGE_KEY_TARGET_PAGE = "examix_import_target_page";
  const STORAGE_KEY_MODE = "examix_import_mode";
  const STORAGE_KEY_TARGET_COLUMN = "examix_import_target_column";
  const STORAGE_KEY_POS = "examix_toolbar_pos";
  const STORAGE_KEY_COLLAPSED = "examix_toolbar_collapsed";
  const STORAGE_KEY_HIDDEN = "examix_toolbar_hidden";
  const STORAGE_KEY_ABSENT_ZERO = "examix_fill_absent_zero";

  const toEnglishDigits = (value) =>
    String(value ?? "").replace(/[٠-٩۰-۹]/g, (digit) => {
      const code = digit.charCodeAt(0);
      return String(code >= 0x6f0 ? code - 0x6f0 : code - 0x660);
    });

  const normalizeText = (value) =>
    toEnglishDigits(String(value ?? ""))
      .normalize("NFKD")
      .replace(/[\u064b-\u065f\u0670\u0640]/g, "")
      .replace(/[أإآٱ]/g, "ا")
      .replace(/ى/g, "ي")
      .replace(/ة/g, "ه")
      .replace(/\s+/g, "")
      .trim();

  const normalizeIdentity = (value) =>
    toEnglishDigits(String(value ?? "")).replace(/[\s\-_]/g, "");

  function parseXml(text, description) {
    const doc = new DOMParser().parseFromString(text, "application/xml");
    if (doc.querySelector("parsererror")) {
      throw new Error(`تعذر قراءة ${description} داخل ملف Excel.`);
    }
    return doc;
  }

  async function readZipEntries(buffer) {
    const bytes = new Uint8Array(buffer);
    const view = new DataView(buffer);
    const minOffset = Math.max(0, bytes.length - 22 - 65535);
    let endOffset = -1;

    for (let offset = bytes.length - 22; offset >= minOffset; offset--) {
      if (view.getUint32(offset, true) === 0x06054b50) {
        endOffset = offset;
        break;
      }
    }
    if (endOffset < 0) {
      throw new Error("الملف ليس مصنف Excel بصيغة ‎.xlsx صالحة.");
    }

    const entryCount = view.getUint16(endOffset + 10, true);
    const directoryOffset = view.getUint32(endOffset + 16, true);
    const decoder = new TextDecoder();
    const entries = new Map();
    let offset = directoryOffset;

    for (let index = 0; index < entryCount; index++) {
      if (view.getUint32(offset, true) !== 0x02014b50) {
        throw new Error("تعذر قراءة بنية ملف Excel.");
      }
      const method = view.getUint16(offset + 10, true);
      const compressedSize = view.getUint32(offset + 20, true);
      const nameLength = view.getUint16(offset + 28, true);
      const extraLength = view.getUint16(offset + 30, true);
      const commentLength = view.getUint16(offset + 32, true);
      const localOffset = view.getUint32(offset + 42, true);
      const name = decoder.decode(
        bytes.subarray(offset + 46, offset + 46 + nameLength),
      );

      if (view.getUint32(localOffset, true) !== 0x04034b50) {
        throw new Error("تعذر قراءة أحد مكونات ملف Excel.");
      }
      const localNameLength = view.getUint16(localOffset + 26, true);
      const localExtraLength = view.getUint16(localOffset + 28, true);
      const dataOffset = localOffset + 30 + localNameLength + localExtraLength;
      const compressedData = bytes.slice(
        dataOffset,
        dataOffset + compressedSize,
      );
      let data;

      if (method === 0) {
        data = compressedData;
      } else if (method === 8 && typeof DecompressionStream !== "undefined") {
        const stream = new Blob([compressedData])
          .stream()
          .pipeThrough(new DecompressionStream("deflate-raw"));
        data = new Uint8Array(await new Response(stream).arrayBuffer());
      } else {
        throw new Error("طريقة ضغط ملف Excel غير مدعومة في هذا المتصفح.");
      }

      entries.set(name, decoder.decode(data));
      offset += 46 + nameLength + extraLength + commentLength;
    }
    return entries;
  }

  function columnIndex(cellReference) {
    const letters = cellReference?.match(/^[A-Z]+/i)?.[0]?.toUpperCase();
    if (!letters) return -1;
    return (
      [...letters].reduce(
        (value, letter) => value * 26 + letter.charCodeAt(0) - 64,
        0,
      ) - 1
    );
  }

  async function readXlsx(file) {
    const entries = await readZipEntries(await file.arrayBuffer());
    const workbookXml = entries.get("xl/workbook.xml");
    if (!workbookXml) {
      throw new Error("لم يتم العثور على ملف بيانات المصنف (xl/workbook.xml).");
    }

    const workbook = parseXml(workbookXml, "بيانات المصنف");
    const relationships = entries.get("xl/_rels/workbook.xml.rels");
    let firstSheetPath = "xl/worksheets/sheet1.xml";
    const firstSheet = workbook.getElementsByTagName("sheet")[0];

    if (firstSheet && relationships) {
      const relationshipId =
        firstSheet.getAttributeNS(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "id",
        ) || firstSheet.getAttribute("r:id");
      const rels = parseXml(relationships, "روابط أوراق العمل");
      const relationship = [...rels.getElementsByTagName("Relationship")].find(
        (item) => item.getAttribute("Id") === relationshipId,
      );
      if (relationship) {
        const target = relationship.getAttribute("Target");
        firstSheetPath = target.startsWith("/")
          ? target.slice(1)
          : `xl/${target}`.replace(/\/(?:\.\/)?/g, "/");
        const parts = [];
        for (const part of firstSheetPath.split("/")) {
          if (part === "..") parts.pop();
          else if (part !== ".") parts.push(part);
        }
        firstSheetPath = parts.join("/");
      }
    }

    let worksheetXml = entries.get(firstSheetPath);
    if (!worksheetXml) {
      for (const [key, val] of entries.entries()) {
        if (key.startsWith("xl/worksheets/sheet") && key.endsWith(".xml")) {
          worksheetXml = val;
          break;
        }
      }
    }
    if (!worksheetXml) {
      throw new Error("تعذر العثور على ورقة العمل الأولى في ملف Excel.");
    }

    const sharedStringsXml = entries.get("xl/sharedStrings.xml");
    const sharedStrings = sharedStringsXml
      ? [
          ...parseXml(sharedStringsXml, "النصوص المشتركة").getElementsByTagName(
            "si",
          ),
        ].map((item) =>
          [...item.getElementsByTagName("t")]
            .map((text) => text.textContent ?? "")
            .join(""),
        )
      : [];
    const worksheet = parseXml(worksheetXml, "ورقة العمل");

    return [...worksheet.getElementsByTagName("row")].map((row) => {
      const values = [];
      let nextIndex = 0;
      for (const cell of row.children) {
        if (cell.localName !== "c") continue;
        const rAttr = cell.getAttribute("r");
        const index = rAttr ? columnIndex(rAttr) : nextIndex;
        if (index >= 0) nextIndex = index + 1;
        const valueNode = cell.getElementsByTagName("v")[0];
        let value = valueNode?.textContent ?? "";
        if (cell.getAttribute("t") === "s") {
          value = sharedStrings[Number(value)] ?? "";
        } else if (cell.getAttribute("t") === "inlineStr") {
          value = [...cell.getElementsByTagName("t")]
            .map((text) => text.textContent ?? "")
            .join("");
        }
        if (index >= 0) values[index] = value;
      }
      return values;
    });
  }

  function readCsv(text) {
    const firstLine = text.split(/\r?\n/, 1)[0] ?? "";
    const delimiter = ["\t", ";", ","].sort(
      (left, right) =>
        firstLine.split(right).length - firstLine.split(left).length,
    )[0];
    const rows = [];
    let row = [];
    let value = "";
    let quoted = false;

    for (let index = 0; index < text.length; index++) {
      const character = text[index];
      if (character === '"' && quoted && text[index + 1] === '"') {
        value += '"';
        index++;
      } else if (character === '"') {
        quoted = !quoted;
      } else if (character === delimiter && !quoted) {
        row.push(value);
        value = "";
      } else if ((character === "\n" || character === "\r") && !quoted) {
        if (character === "\r" && text[index + 1] === "\n") index++;
        row.push(value);
        if (row.some((cell) => String(cell).trim())) rows.push(row);
        row = [];
        value = "";
      } else {
        value += character;
      }
    }
    row.push(value);
    if (row.some((cell) => String(cell).trim())) rows.push(row);
    if (rows.length && rows[0].length && typeof rows[0][0] === "string") {
      rows[0][0] = rows[0][0].replace(/^\uFEFF/, "");
    }
    return rows;
  }

  async function readFileData(file) {
    if (file.name.toLowerCase().endsWith(".csv")) {
      const buffer = await file.arrayBuffer();
      let text = new TextDecoder("utf-8").decode(buffer);
      if (text.includes("\ufffd")) {
        try {
          const decoded1256 = new TextDecoder("windows-1256").decode(buffer);
          if (!decoded1256.includes("\ufffd")) {
            text = decoded1256;
          }
        } catch {}
      }
      return readCsv(text);
    }
    return await readXlsx(file);
  }

  function findIdentityColumn(normalized) {
    let identity = normalized.findIndex(
      (h) => h.includes("هويه") || h.includes("سجل") || h.includes("اقامه"),
    );
    if (identity < 0) {
      identity = normalized.findIndex(
        (h) =>
          (h.includes("رقم") || h.includes("id")) &&
          !h.includes("صف") &&
          !h.includes("فصل") &&
          !h.includes("جلوس") &&
          !h.includes("هاتف") &&
          !h.includes("جوال"),
      );
    }
    return identity;
  }

  function findDualColumns(headers) {
    const normalized = headers.map(normalizeText);
    const identity = findIdentityColumn(normalized);

    let name = normalized.findIndex(
      (h) =>
        (h.includes("اسم") && (h.includes("طالب") || h.includes("طالبه"))) ||
        h === "الاسم" ||
        h === "اسم",
    );
    if (name < 0) {
      name = normalized.findIndex(
        (h) =>
          h.includes("اسم") &&
          !h.includes("ولي") &&
          !h.includes("مدرسه") &&
          !h.includes("ماده"),
      );
    }

    const performance = normalized.findIndex(
      (h) =>
        h.includes("مهام") ||
        h.includes("ادائي") ||
        h.includes("اعمال") ||
        h.includes("السنه") ||
        h.includes("مشاركه") ||
        h.includes("تفاعل"),
    );

    const assessment = normalized.findIndex(
      (h) =>
        h.includes("تقويم") ||
        h.includes("تقييم") ||
        h.includes("اختبار") ||
        h.includes("تحريري") ||
        h.includes("نهائي"),
    );

    if ([identity, name, performance, assessment].some((idx) => idx < 0)) {
      return null;
    }
    return { identity, name, performance, assessment };
  }

  /// يكتشف أعمدة ملف إكزامكس المصدّر: رقم الهوية + الدرجة الكلية (أو الدرجة).
  function findTotalGradeColumns(rows) {
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      const normalized = rows[i].map(normalizeText);
      const identity = findIdentityColumn(normalized);
      if (identity < 0) continue;

      let grade = normalized.findIndex(
        (header) =>
          header.includes("درجه") &&
          (header.includes("كليه") || header.includes("نهائي")),
      );
      if (grade < 0) {
        grade = normalized.findIndex(
          (header) =>
            header === "الدرجه" ||
            header === "درجه" ||
            header === "grade" ||
            header === "score",
        );
      }
      if (grade >= 0) {
        return { headerRowIndex: i, identity, grade };
      }
    }
    return null;
  }

  function scoreValue(value) {
    const text = toEnglishDigits(String(value ?? ""))
      .trim()
      .replace(/٫/g, ".")
      .replace(/٬/g, "")
      .replace(/,/g, ".");
    if (!text || !/^\d+(?:\.\d+)?$/.test(text)) return null;
    const score = Number(text);
    return Number.isFinite(score) ? score : null;
  }

  function maxMark(input, header) {
    const headerText = toEnglishDigits(header?.textContent ?? "");
    const headerMax = headerText.match(/\/\s*(\d+(?:\.\d+)?)/);
    const inputMax = input?.getAttribute
      ? input.getAttribute("CourseMaxMark")
      : null;
    const parsed = headerMax ? Number(headerMax[1]) : Number(inputMax);
    return Number.isFinite(parsed) && parsed >= 0 ? parsed : null;
  }

  function parseTotalGradeRows(rows, columns, { fillAbsentAsZero = true } = {}) {
    if (columns.identity < 0) {
      throw new Error("لم يتم العثور على عمود «رقم الهوية» في ملف Excel.");
    }

    const seenIdentities = new Set();
    const sourceRows = [];
    let absentSkipped = 0;

    for (let i = columns.headerRowIndex + 1; i < rows.length; i++) {
      const cells = rows[i];
      if (!cells || cells.length === 0) continue;

      const rawIdentity = String(cells[columns.identity] ?? "").trim();
      if (!rawIdentity) continue;
      const identity = normalizeIdentity(rawIdentity);
      if (!/^\d{9,10}$/.test(identity)) continue;

      const rawGrade = String(cells[columns.grade] ?? "").trim();
      if (!rawGrade || rawGrade === "-") continue;

      let grade;
      if (normalizeText(rawGrade).startsWith("غائب")) {
        if (!fillAbsentAsZero) {
          absentSkipped++;
          continue;
        }
        grade = 0;
      } else {
        grade = scoreValue(rawGrade);
        if (grade === null) {
          throw new Error(`درجة غير صالحة في الصف ${i + 1}: "${rawGrade}".`);
        }
      }

      if (seenIdentities.has(identity)) continue;
      seenIdentities.add(identity);
      sourceRows.push({ identity, grade });
    }

    if (sourceRows.length === 0) {
      throw new Error(
        absentSkipped > 0
          ? "جميع الصفوف في الملف لطالبات/طلاب غائبين."
          : "لم يتم العثور على درجات صالحة داخل الملف.",
      );
    }
    return sourceRows;
  }

  function parseDualSourceRows(rows) {
    let headerRowIndex = -1;
    let columns = null;
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      columns = findDualColumns(rows[i]);
      if (columns) {
        headerRowIndex = i;
        break;
      }
    }
    if (headerRowIndex < 0 || !columns) {
      throw new Error(
        "لم يتم التعرف على أعمدة ملف Excel. يجب أن يحتوي ملف إكزامكس على عمود «رقم الهوية» وعمود «الدرجة الكلية».",
      );
    }

    const seenIdentities = new Set();
    const sourceRows = [];

    for (let i = headerRowIndex + 1; i < rows.length; i++) {
      const cells = rows[i];
      if (!cells || cells.length === 0) continue;
      const identity = normalizeIdentity(cells[columns.identity]);
      const name = normalizeText(cells[columns.name]);
      if (!/^\d{9,10}$/.test(identity) || !name) continue;
      if (seenIdentities.has(identity)) continue;
      seenIdentities.add(identity);

      sourceRows.push({
        identity,
        performance: scoreValue(cells[columns.performance]),
        assessment: scoreValue(cells[columns.assessment]),
        hasPerformance: String(cells[columns.performance] ?? "").trim() !== "",
        hasAssessment: String(cells[columns.assessment] ?? "").trim() !== "",
      });
    }

    if (sourceRows.length === 0) {
      throw new Error("لم يتم العثور على بيانات طلاب صالحة داخل الملف.");
    }
    return sourceRows;
  }

  function parsePastedGrades(text, { fillAbsentAsZero = true } = {}) {
    const sourceRows = [];
    const seenIdentities = new Set();
    const trimmed = text.trim();

    if (!trimmed) {
      throw new Error("الصق بيانات إكزامكس (TSV أو JSON) أولاً.");
    }

    if (trimmed.startsWith("{")) {
      let grades;
      try {
        grades = JSON.parse(trimmed.replace(/,\s*([}\]])/g, "$1"));
      } catch {
        throw new Error(
          'صيغة JSON غير صحيحة. استخدم كائنًا بالشكل {"رقم الهوية": الدرجة}.',
        );
      }
      if (!grades || Array.isArray(grades) || typeof grades !== "object") {
        throw new Error(
          "يجب أن يحتوي JSON على أرقام الهوية كمفاتيح والدرجات كقيم.",
        );
      }

      for (const [rawIdentity, rawGrade] of Object.entries(grades)) {
        const identity = normalizeIdentity(rawIdentity);
        const grade = scoreValue(rawGrade);
        if (!/^\d{9,10}$/.test(identity) || grade === null) {
          throw new Error(
            `بيانات غير صالحة في JSON عند رقم الهوية: ${rawIdentity}.`,
          );
        }
        sourceRows.push({ identity, grade });
      }
    } else {
      const lines = trimmed.split(/\r?\n/);
      for (let index = 0; index < lines.length; index++) {
        const line = lines[index].trim();
        if (!line) continue;

        const columns = line.includes("\t")
          ? line.split("\t").map((col) => col.trim())
          : line.split(/[;,]\s*/).map((col) => col.trim());
        if (columns.length !== 2) {
          throw new Error(
            `السطر ${index + 1} غير صالح. الصيغة المتوقعة: رقم الهوية [Tab] الدرجة.`,
          );
        }

        const identity = normalizeIdentity(columns[0]);
        if (!/^\d{9,10}$/.test(identity)) {
          throw new Error(`رقم هوية غير صالح في السطر ${index + 1}.`);
        }

        let grade;
        if (normalizeText(columns[1]).startsWith("غائب")) {
          if (!fillAbsentAsZero) continue;
          grade = 0;
        } else {
          grade = scoreValue(columns[1]);
        }

        if (grade === null) {
          throw new Error(
            `درجة غير صالحة في السطر ${index + 1}: "${columns[1]}".`,
          );
        }
        sourceRows.push({ identity, grade });
      }
    }

    const uniqueRows = sourceRows.filter((src) => {
      if (seenIdentities.has(src.identity)) return false;
      seenIdentities.add(src.identity);
      return true;
    });
    if (uniqueRows.length === 0) {
      throw new Error("لم يتم العثور على درجات صالحة لاستيرادها.");
    }
    return uniqueRows;
  }

  function detectTableStructure(table) {
    const headerRow =
      table.querySelector("tr.tableHeaderCell") || table.rows?.[0];
    let performanceColIndex = 2;
    let assessmentColIndex = 3;
    let identityColIndex = 0;
    let nameColIndex = 1;
    let performanceHeader = null;
    let assessmentHeader = null;

    if (headerRow) {
      const ths = [...headerRow.children];
      ths.forEach((th, idx) => {
        const text = normalizeText(th.textContent);
        const html = th.innerHTML;

        if (text.includes("هويه") || text.includes("سجل")) {
          identityColIndex = idx;
        } else if (text.includes("طالب") || text.includes("اسم")) {
          nameColIndex = idx;
        } else if (
          text.includes("مهام") ||
          text.includes("ادائي") ||
          text.includes("مشاركه") ||
          html.includes("Header2") ||
          html.includes("tblHeader2")
        ) {
          performanceColIndex = idx;
          performanceHeader = th;
        } else if (
          text.includes("تقويم") ||
          text.includes("تقييم") ||
          text.includes("اختبار") ||
          html.includes("Header3") ||
          html.includes("tblHeader3")
        ) {
          assessmentColIndex = idx;
          assessmentHeader = th;
        }
      });
    }

    if (!performanceHeader) {
      performanceHeader = table
        .querySelector('th [id*="lblHeader2"], th span[id*="Header2"]')
        ?.closest("th");
    }
    if (!assessmentHeader) {
      assessmentHeader = table
        .querySelector('th [id*="lblHeader3"], th span[id*="Header3"]')
        ?.closest("th");
    }

    return {
      identityColIndex,
      nameColIndex,
      performanceColIndex,
      assessmentColIndex,
      performanceHeader,
      assessmentHeader,
    };
  }

  function updateInputValue(input, mark) {
    if (!input) return;
    input.value = String(mark);
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    input.dispatchEvent(new Event("blur", { bubbles: true }));

    // تحديث حقل hdnDynamic المخفي في نفس الخلية لضمان الحفظ مع ASP.NET WebForms
    const cell = input.closest("td");
    const hiddenInput = cell?.querySelector(
      'input[type="hidden"][id*="hdnDynamic"]',
    );
    if (hiddenInput && hiddenInput.value) {
      const parts = hiddenInput.value.split(",");
      if (parts.length >= 6) {
        parts[parts.length - 1] = String(mark);
        hiddenInput.value = parts.join(",");
      }
    }
  }

  function processCurrentPage(sourceMap, selectedColumn = null) {
    const table = getTargetTable();
    if (!table) return { filledCount: 0, totalPageStudents: 0 };

    const structure = detectTableStructure(table);
    const {
      identityColIndex,
      performanceColIndex,
      assessmentColIndex,
      performanceHeader,
      assessmentHeader,
    } = structure;

    const pageRows = [...table.querySelectorAll("tr")].filter((row) => {
      const cellText = row.cells?.[identityColIndex]?.textContent?.trim() || "";
      return /^\d{9,10}$/.test(normalizeIdentity(cellText));
    });

    let filledCount = 0;

    for (const pageRow of pageRows) {
      const identity = normalizeIdentity(
        pageRow.cells[identityColIndex]?.textContent,
      );
      if (!identity || !sourceMap.has(identity)) continue;

      const source = sourceMap.get(identity);
      if (selectedColumn !== null) {
        const targetCell = pageRow.cells[selectedColumn];
        const targetInput = targetCell?.querySelector(
          'input[type="text"], input[name*="tbMark"]',
        );
        const headerRow =
          table.querySelector("tr.tableHeaderCell") || table.rows?.[0];
        const targetHeader = headerRow?.cells?.[selectedColumn];
        const targetMax = targetInput
          ? maxMark(targetInput, targetHeader)
          : null;

        if (
          targetInput &&
          source.grade >= 0 &&
          (targetMax === null || source.grade <= targetMax)
        ) {
          updateInputValue(targetInput, source.grade);
          filledCount++;
        }
        continue;
      }

      const performanceCell = pageRow.cells[performanceColIndex];
      const assessmentCell = pageRow.cells[assessmentColIndex];

      const performanceInput = performanceCell?.querySelector(
        'input[type="text"], input[name*="tbMark"]',
      );
      const assessmentInput = assessmentCell?.querySelector(
        'input[type="text"], input[name*="tbMark"]',
      );

      const performanceMax =
        performanceInput && performanceHeader
          ? maxMark(performanceInput, performanceHeader)
          : null;
      const assessmentMax =
        assessmentInput && assessmentHeader
          ? maxMark(assessmentInput, assessmentHeader)
          : null;

      let studentUpdated = false;

      if (
        source.hasPerformance &&
        source.performance !== null &&
        performanceInput &&
        (performanceMax === null || source.performance <= performanceMax) &&
        source.performance >= 0
      ) {
        updateInputValue(performanceInput, source.performance);
        studentUpdated = true;
      }

      if (
        source.hasAssessment &&
        source.assessment !== null &&
        assessmentInput &&
        (assessmentMax === null || source.assessment <= assessmentMax) &&
        source.assessment >= 0
      ) {
        updateInputValue(assessmentInput, source.assessment);
        studentUpdated = true;
      }

      if (studentUpdated) {
        filledCount++;
      }
    }

    return { filledCount, totalPageStudents: pageRows.length };
  }

  function clearSelectedColumnCurrentPage(columnIndex) {
    const table = getTargetTable();
    if (!table) {
      return {
        success: false,
        message: "لم يتم العثور على جدول رصد الدرجات في هذه الصفحة.",
      };
    }

    const structure = detectTableStructure(table);
    const { identityColIndex } = structure;

    const pageRows = [...table.querySelectorAll("tr")].filter((row) => {
      const cellText = row.cells?.[identityColIndex]?.textContent?.trim() || "";
      return /^\d{9,10}$/.test(normalizeIdentity(cellText));
    });

    if (pageRows.length === 0) {
      return {
        success: false,
        message: "لا توجد صفوف طلاب في الجدول لتصفيرها.",
      };
    }

    let count = 0;
    for (const pageRow of pageRows) {
      const targetCell = pageRow.cells[columnIndex];
      const targetInput = targetCell?.querySelector(
        'input[type="text"], input[name*="tbMark"]',
      );
      if (targetInput) {
        updateInputValue(targetInput, "");
        count++;
      }
    }

    return {
      success: true,
      count,
      message: `تم تفريغ العمود (${columnIndex + 1}) لـ ${count} طالب في هذه الصفحة. تذكر الضغط على حفظ في نور.`,
    };
  }

  function findNextPageButton(targetPageNum) {
    const pagerRow = document.querySelector("tr.MyPagerStyle");
    if (!pagerRow) return null;

    const links = [...pagerRow.querySelectorAll("a")];

    for (const link of links) {
      const text = toEnglishDigits(link.textContent.trim());
      if (text === String(targetPageNum)) {
        return link;
      }
    }

    for (const link of links) {
      const match = link.href?.match(/Page\$(\d+)/);
      if (match && Number(match[1]) === targetPageNum) {
        return link;
      }
    }

    for (const link of links) {
      const text = link.textContent.trim();
      if (text === ">" || text.includes("التالي") || text === ">>") {
        return link;
      }
    }

    return null;
  }

  function clearImportSession() {
    for (const key of [
      STORAGE_KEY_ACTIVE,
      STORAGE_KEY_DATA,
      STORAGE_KEY_STATS,
      STORAGE_KEY_TARGET_PAGE,
      STORAGE_KEY_MODE,
      STORAGE_KEY_TARGET_COLUMN,
    ]) {
      sessionStorage.removeItem(key);
      localStorage.removeItem(key);
    }
  }

  function restorePosition(element) {
    try {
      const savedPos = localStorage.getItem(STORAGE_KEY_POS);
      if (savedPos) {
        const { top, left } = JSON.parse(savedPos);
        const pTop = parseFloat(top);
        const pLeft = parseFloat(left);
        if (Number.isFinite(pTop) && Number.isFinite(pLeft)) {
          const maxLeft = Math.max(10, window.innerWidth - 100);
          const maxTop = Math.max(10, window.innerHeight - 60);
          element.style.top = `${Math.max(10, Math.min(maxTop, pTop))}px`;
          element.style.left = `${Math.max(10, Math.min(maxLeft, pLeft))}px`;
          element.style.right = "auto";
          element.style.bottom = "auto";
          return true;
        }
      }
    } catch {}
    return false;
  }

  function makeDraggable(element, handle) {
    let isDragging = false;
    let startX = 0;
    let startY = 0;
    let startLeft = 0;
    let startTop = 0;

    handle.style.cursor = "grab";

    const onPointerDown = (e) => {
      if (
        e.target.closest("button") ||
        e.target.closest("input") ||
        e.target.closest("a") ||
        e.target.closest("select")
      ) {
        return;
      }
      isDragging = true;
      handle.style.cursor = "grabbing";

      const rect = element.getBoundingClientRect();
      startX = e.clientX;
      startY = e.clientY;
      startLeft = rect.left;
      startTop = rect.top;

      element.style.right = "auto";
      element.style.bottom = "auto";
      element.style.left = `${startLeft}px`;
      element.style.top = `${startTop}px`;
    };

    const onPointerMove = (e) => {
      if (!isDragging) return;
      const deltaX = e.clientX - startX;
      const deltaY = e.clientY - startY;

      const maxLeft = Math.max(5, window.innerWidth - element.offsetWidth - 5);
      const maxTop = Math.max(
        5,
        window.innerHeight - element.offsetHeight - 5,
      );

      const newLeft = Math.max(5, Math.min(maxLeft, startLeft + deltaX));
      const newTop = Math.max(5, Math.min(maxTop, startTop + deltaY));

      element.style.left = `${newLeft}px`;
      element.style.top = `${newTop}px`;
    };

    const onPointerUp = () => {
      if (!isDragging) return;
      isDragging = false;
      handle.style.cursor = "grab";
      try {
        localStorage.setItem(
          STORAGE_KEY_POS,
          JSON.stringify({
            top: element.style.top,
            left: element.style.left,
          }),
        );
      } catch {}
    };

    handle.addEventListener("mousedown", onPointerDown);
    window.addEventListener("mousemove", onPointerMove);
    window.addEventListener("mouseup", onPointerUp);

    handle.addEventListener(
      "touchstart",
      (e) => {
        if (e.touches.length === 1) {
          const touch = e.touches[0];
          onPointerDown({
            target: e.target,
            clientX: touch.clientX,
            clientY: touch.clientY,
          });
        }
      },
      { passive: true },
    );

    window.addEventListener(
      "touchmove",
      (e) => {
        if (isDragging && e.touches.length === 1) {
          const touch = e.touches[0];
          onPointerMove({
            clientX: touch.clientX,
            clientY: touch.clientY,
          });
        }
      },
      { passive: true },
    );

    window.addEventListener("touchend", onPointerUp);
  }

  function addToolbar() {
    if (document.getElementById(TOOLBAR_ID)) return;

    const toolbar = document.createElement("section");
    toolbar.id = TOOLBAR_ID;
    toolbar.dir = "rtl";
    toolbar.setAttribute("aria-label", "إكزامكس — رصد الدرجات في نظام نور");
    toolbar.style.cssText =
      "position:fixed;top:16px;left:16px;z-index:2147483647;background:#ffffff;color:#172044;border:1px solid #cbd5e1;border-radius:12px;overflow:hidden;box-shadow:0 12px 32px rgba(16,22,93,0.22);font:13px Tahoma,Arial,sans-serif;width:365px;max-width:calc(100vw - 24px);box-sizing:border-box;";

    restorePosition(toolbar);

    function showLauncher() {
      if (document.getElementById(LAUNCHER_ID)) return;
      const launcher = document.createElement("button");
      launcher.id = LAUNCHER_ID;
      launcher.dir = "rtl";
      launcher.type = "button";
      launcher.setAttribute("aria-label", "إظهار لوحة إكزامكس لرصد الدرجات");
      launcher.style.cssText =
        "position:fixed;bottom:16px;left:16px;z-index:2147483647;background:#10165d;color:#fff;border:1px solid #20c9e8;border-radius:999px;padding:9px 16px;box-shadow:0 6px 20px rgba(16,22,93,0.3);font:bold 12px Tahoma,Arial,sans-serif;cursor:pointer;display:flex;align-items:center;gap:8px;";
      launcher.innerHTML = `<span style="display:inline-block;width:9px;height:9px;border-radius:50%;background:#08c995;"></span><span>إكزامكس — رصد نور</span>`;

      launcher.addEventListener("click", () => {
        launcher.remove();
        toolbar.style.display = "block";
        try {
          localStorage.setItem(STORAGE_KEY_HIDDEN, "false");
        } catch {}
      });
      document.body.append(launcher);
    }

    // شريط العنوان
    const header = document.createElement("div");
    header.style.cssText =
      "display:flex;justify-content:space-between;align-items:center;padding:10px 12px;background:linear-gradient(120deg,#10165d,#1a237e);color:#fff;gap:8px;user-select:none;";
    header.title = "اسحب من هنا لتحريك اللوحة (انقر نقراً مزدوجاً للطي/التوسيع)";

    const titleGroup = document.createElement("div");
    titleGroup.style.cssText = "display:flex;align-items:center;gap:7px;";

    const dragHandle = document.createElement("span");
    dragHandle.textContent = "⠿";
    dragHandle.style.cssText =
      "cursor:grab;color:#8df0d8;font-size:16px;line-height:1;";

    const label = document.createElement("strong");
    label.textContent = "إكزامكس | رصد الدرجات في نور";
    label.style.cssText = "font-size:13px;white-space:nowrap;color:#fff;";

    titleGroup.append(dragHandle, label);

    const controlsGroup = document.createElement("div");
    controlsGroup.style.cssText = "display:flex;align-items:center;gap:4px;";

    const resetButton = document.createElement("button");
    resetButton.type = "button";
    resetButton.textContent = "إلغاء";
    resetButton.title = "إلغاء عملية الرصد الحالية";
    resetButton.style.cssText =
      "background:#e74c3c;color:#fff;border:none;border-radius:5px;padding:3px 8px;cursor:pointer;font-size:11px;font-weight:bold;";

    const collapseBtn = document.createElement("button");
    collapseBtn.type = "button";
    collapseBtn.textContent = "−";
    collapseBtn.title = "طي / توسيع اللوحة";
    collapseBtn.style.cssText =
      "background:rgba(255,255,255,0.16);color:#fff;border:none;border-radius:5px;width:22px;height:22px;cursor:pointer;font-size:14px;font-weight:bold;line-height:1;";

    const closeBtn = document.createElement("button");
    closeBtn.type = "button";
    closeBtn.textContent = "✕";
    closeBtn.title = "إخفاء اللوحة مؤقتاً";
    closeBtn.style.cssText =
      "background:rgba(255,255,255,0.16);color:#fff;border:none;border-radius:5px;width:22px;height:22px;cursor:pointer;font-size:12px;font-weight:bold;line-height:1;";

    controlsGroup.append(resetButton, collapseBtn, closeBtn);
    header.append(titleGroup, controlsGroup);

    // جسم اللوحة
    const bodyContainer = document.createElement("div");
    bodyContainer.style.cssText = "padding:12px;";

    // خيار ملف Excel
    const fileLabel = document.createElement("label");
    fileLabel.style.cssText =
      "display:block;font-size:12px;font-weight:bold;color:#10165d;margin-bottom:4px;";
    fileLabel.textContent = "1. استيراد ملف إكسل المصدّر من إكزامكس (.xlsx / .csv):";

    const input = document.createElement("input");
    input.type = "file";
    input.accept =
      ".xlsx,.csv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    input.style.cssText =
      "display:block;width:100%;padding:6px;border:1px dashed #94a3b8;border-radius:8px;background:#f8fafc;font-size:12px;box-sizing:border-box;cursor:pointer;";

    // خيار اللصق
    const pasteButton = document.createElement("button");
    pasteButton.id = PASTE_BUTTON_ID;
    pasteButton.type = "button";
    pasteButton.textContent = "📋 2. لصق درجات «نسخ لنور» من إكزامكس";
    pasteButton.title =
      "لصق البيانات المنسوخة من زر (نسخ لنور) في شاشة تقارير إكزامكس";
    pasteButton.style.cssText =
      "display:block;width:100%;margin-top:10px;background:linear-gradient(115deg,#08c995,#0ea5e9);color:#fff;border:none;border-radius:8px;padding:9px 10px;cursor:pointer;font-size:12.5px;font-weight:bold;box-shadow:0 4px 12px rgba(8,201,149,0.25);";

    // خيار الغائب (0)
    const absentRow = document.createElement("label");
    absentRow.style.cssText =
      "display:flex;align-items:center;gap:7px;margin-top:10px;font-size:12px;color:#334155;cursor:pointer;user-select:none;";
    const absentCheckbox = document.createElement("input");
    absentCheckbox.type = "checkbox";
    absentCheckbox.checked =
      localStorage.getItem(STORAGE_KEY_ABSENT_ZERO) !== "false";
    absentCheckbox.addEventListener("change", () => {
      try {
        localStorage.setItem(
          STORAGE_KEY_ABSENT_ZERO,
          absentCheckbox.checked ? "true" : "false",
        );
      } catch {}
    });
    const absentText = document.createElement("span");
    absentText.textContent = "رصد الطالب الغائب «غائب (0)» بدرجة 0";
    absentRow.append(absentCheckbox, absentText);

    // زر تصفير عمود محدد
    const clearColButton = document.createElement("button");
    clearColButton.type = "button";
    clearColButton.textContent = "🧹 تصفير عمود درجات في هذه الصفحة";
    clearColButton.title = "اختر عمود درجات لتفريغه لجميع طلاب الصفحة الحالية";
    clearColButton.style.cssText =
      "display:block;width:100%;margin-top:8px;background:#f1f5f9;color:#475569;border:1px solid #cbd5e1;border-radius:7px;padding:6px 8px;cursor:pointer;font-size:11.5px;font-weight:bold;";

    const status = document.createElement("p");
    status.setAttribute("aria-live", "polite");
    status.style.cssText =
      "margin:10px 0 0;padding:8px 10px;border-radius:7px;background:#f8fafc;border:1px solid #e2e8f0;color:#1e293b;font-size:12px;line-height:1.55;";
    status.textContent =
      "اختر ملف Excel المصدّر من إكزامكس أو اضغط «لصق درجات نسخ لنور»، ثم حدد مربع الدرجة المطلوب في جدول نور.";

    bodyContainer.append(
      fileLabel,
      input,
      pasteButton,
      absentRow,
      clearColButton,
      status,
    );

    // نافذة اللصق
    const pasteDialog = document.createElement("dialog");
    pasteDialog.dir = "rtl";
    pasteDialog.setAttribute("aria-label", "لصق درجات إكزامكس");
    pasteDialog.style.cssText =
      "width:min(540px,calc(100vw - 32px));border:1px solid #cbd5e1;border-radius:14px;padding:18px;box-shadow:0 16px 40px rgba(16,22,93,0.28);font:13.5px Tahoma,Arial,sans-serif;color:#172044;";

    const pasteTitle = document.createElement("h2");
    pasteTitle.textContent = "لصق درجات إكزامكس (نسخ لنور)";
    pasteTitle.style.cssText = "font-size:16px;margin:0 0 6px;color:#10165d;";

    const pasteHelp = document.createElement("p");
    pasteHelp.textContent =
      "اضغط «لصق من الحافظة» أو الصق البيانات (رقم الهوية [Tab] الدرجة أو JSON) ثم حدد مربع الدرجة في جدول نور.";
    pasteHelp.style.cssText =
      "line-height:1.6;margin:0 0 10px;color:#475569;font-size:12.5px;";

    const pasteFromClipboardBtn = document.createElement("button");
    pasteFromClipboardBtn.type = "button";
    pasteFromClipboardBtn.textContent = "📋 لصق تلقائي من الحافظة";
    pasteFromClipboardBtn.style.cssText =
      "margin-bottom:8px;background:#eef2ff;color:#10165d;border:1px solid #c7d2fe;border-radius:7px;padding:6px 12px;cursor:pointer;font-size:12px;font-weight:bold;";

    const pasteArea = document.createElement("textarea");
    pasteArea.rows = 9;
    pasteArea.placeholder = "1166104156\t25\n1163590076\t18\nأو JSON من إكزامكس...";
    pasteArea.setAttribute("aria-label", "بيانات الهوية والدرجات");
    pasteArea.style.cssText =
      "box-sizing:border-box;width:100%;resize:vertical;font:13px Consolas,monospace;padding:9px;border:1px solid #cbd5e1;border-radius:8px;direction:ltr;text-align:left;";

    const pasteFeedback = document.createElement("p");
    pasteFeedback.setAttribute("aria-live", "polite");
    pasteFeedback.style.cssText =
      "margin:8px 0 0;color:#b91c1c;font-size:12px;line-height:1.5;";

    const dialogActions = document.createElement("div");
    dialogActions.style.cssText =
      "display:flex;flex-wrap:wrap;gap:8px;justify-content:flex-start;margin-top:12px;";

    const chooseFieldButton = document.createElement("button");
    chooseFieldButton.type = "button";
    chooseFieldButton.textContent = "🎯 تحديد مربع الدرجة في نور";
    chooseFieldButton.style.cssText =
      "background:#10165d;color:#fff;border:none;border-radius:7px;padding:8px 13px;cursor:pointer;font-weight:bold;font-size:12.5px;";

    const startPasteButton = document.createElement("button");
    startPasteButton.type = "button";
    startPasteButton.textContent = "▶ بدء الرصد التلقائي";
    startPasteButton.style.cssText =
      "background:#08c995;color:#fff;border:none;border-radius:7px;padding:8px 13px;cursor:pointer;font-weight:bold;font-size:12.5px;";

    const closeDialogButton = document.createElement("button");
    closeDialogButton.type = "button";
    closeDialogButton.textContent = "إغلاق";
    closeDialogButton.style.cssText =
      "background:#e2e8f0;color:#1e293b;border:none;border-radius:7px;padding:8px 13px;cursor:pointer;font-size:12.5px;";

    dialogActions.append(
      chooseFieldButton,
      startPasteButton,
      closeDialogButton,
    );
    pasteDialog.append(
      pasteTitle,
      pasteHelp,
      pasteFromClipboardBtn,
      pasteArea,
      pasteFeedback,
      dialogActions,
    );

    // شريط إشعار اختيار العمود
    const pickerNotice = document.createElement("div");
    pickerNotice.dir = "rtl";
    pickerNotice.style.cssText =
      "position:fixed;bottom:18px;left:50%;transform:translateX(-50%);z-index:2147483647;background:#10165d;color:#fff;border:2px solid #08c995;border-radius:10px;padding:11px 16px;font:bold 13px Tahoma,Arial,sans-serif;box-shadow:0 8px 26px rgba(0,0,0,0.35);display:flex;align-items:center;gap:12px;";
    const pickerText = document.createElement("span");
    pickerText.textContent =
      "انقر داخل مربع الدرجة المطلوب في جدول نور، أو اضغط Esc للإلغاء.";
    const cancelPickButton = document.createElement("button");
    cancelPickButton.type = "button";
    cancelPickButton.textContent = "إلغاء";
    cancelPickButton.style.cssText =
      "background:#e74c3c;color:#fff;border:0;border-radius:6px;padding:4px 10px;cursor:pointer;font-size:12px;";
    pickerNotice.append(pickerText, cancelPickButton);

    let isCollapsed = false;
    const setCollapsed = (collapsed) => {
      isCollapsed = collapsed;
      if (isCollapsed) {
        bodyContainer.style.display = "none";
        collapseBtn.textContent = "＋";
        toolbar.style.width = "260px";
      } else {
        bodyContainer.style.display = "block";
        collapseBtn.textContent = "−";
        toolbar.style.width = "365px";
      }
      try {
        localStorage.setItem(
          STORAGE_KEY_COLLAPSED,
          isCollapsed ? "true" : "false",
        );
      } catch {}
    };

    collapseBtn.addEventListener("click", () => setCollapsed(!isCollapsed));
    header.addEventListener("dblclick", (e) => {
      if (e.target.closest("button")) return;
      setCollapsed(!isCollapsed);
    });
    closeBtn.addEventListener("click", () => {
      toolbar.style.display = "none";
      showLauncher();
      try {
        localStorage.setItem(STORAGE_KEY_HIDDEN, "true");
      } catch {}
    });

    toolbar.append(header, bodyContainer);
    document.body.append(toolbar, pasteDialog);
    makeDraggable(toolbar, header);

    window.addEventListener("resize", () => {
      const rect = toolbar.getBoundingClientRect();
      const maxLeft = Math.max(5, window.innerWidth - toolbar.offsetWidth - 5);
      const maxTop = Math.max(5, window.innerHeight - toolbar.offsetHeight - 5);
      if (rect.left > maxLeft || rect.top > maxTop) {
        toolbar.style.left = `${Math.max(5, Math.min(maxLeft, rect.left))}px`;
        toolbar.style.top = `${Math.max(5, Math.min(maxTop, rect.top))}px`;
        toolbar.style.right = "auto";
      }
    });

    const isImportActive =
      sessionStorage.getItem(STORAGE_KEY_ACTIVE) === "true" ||
      localStorage.getItem(STORAGE_KEY_ACTIVE) === "true";

    const isHidden =
      !isImportActive && localStorage.getItem(STORAGE_KEY_HIDDEN) === "true";

    if (isHidden) {
      toolbar.style.display = "none";
      showLauncher();
    } else {
      toolbar.style.display = "block";
      if (
        localStorage.getItem(STORAGE_KEY_COLLAPSED) === "true" &&
        !isImportActive
      ) {
        setCollapsed(true);
      }
    }

    let pickingField = false;
    let selectedColumn = null;
    let onFieldPicked = null;

    const stopPickingField = () => {
      pickingField = false;
      toolbar.style.pointerEvents = "";
      pickerNotice.remove();
      document.removeEventListener("click", handleFieldPick, true);
      document.removeEventListener("keydown", handlePickerKeydown, true);
    };

    const cancelFieldPicking = () => {
      const wasCallback = Boolean(onFieldPicked);
      onFieldPicked = null;
      stopPickingField();
      if (wasCallback) {
        input.disabled = false;
        input.value = "";
        pasteButton.disabled = false;
      }
      status.textContent = "تم إلغاء تحديد مربع الدرجة.";
    };

    const handlePickerKeydown = (event) => {
      if (event.key === "Escape" && pickingField) {
        cancelFieldPicking();
      }
    };

    const handleFieldPick = (event) => {
      if (!pickingField || !(event.target instanceof Element)) return;
      if (pickerNotice.contains(event.target)) return;
      if (toolbar.contains(event.target)) return;

      const targetInput = event.target.closest(
        'input[type="text"], input[name*="tbMark"]',
      );
      const table = getTargetTable();
      const structure = table ? detectTableStructure(table) : null;
      const studentRow =
        table && structure && targetInput
          ? [...table.querySelectorAll("tr")].find((candidate) => {
              const identity = normalizeIdentity(
                candidate.cells?.[structure.identityColIndex]?.textContent,
              );
              return (
                candidate.contains(targetInput) && /^\d{9,10}$/.test(identity)
              );
            })
          : null;
      const colIdx = studentRow
        ? [...studentRow.cells].findIndex((cell) => cell.contains(targetInput))
        : -1;

      if (
        !table ||
        !targetInput ||
        !table.contains(targetInput) ||
        !studentRow ||
        colIdx < 0
      ) {
        event.preventDefault();
        event.stopImmediatePropagation();
        status.textContent = "انقر داخل مربع نص الدرجة في أحد صفوف الطلاب.";
        return;
      }

      event.preventDefault();
      event.stopImmediatePropagation();
      selectedColumn = colIdx;
      const callback = onFieldPicked;
      onFieldPicked = null;
      stopPickingField();
      status.textContent = `تم تحديد عمود الدرجة رقم (${selectedColumn + 1}).`;
      if (callback) {
        callback(selectedColumn);
        return;
      }

      pasteFeedback.style.color = "#06896b";
      pasteFeedback.textContent = `تم تحديد عمود الدرجة رقم (${selectedColumn + 1}) بنجاح. اضغط «بدء الرصد التلقائي».`;
      pasteDialog.showModal();
      startPasteButton.focus();
    };

    const startFieldPicking = (callback = null) => {
      if (!getTargetTable()) {
        throw new Error(
          "اعرض جدول الطلاب في نظام نور أولاً، ثم اختر مربع الدرجة.",
        );
      }
      onFieldPicked = callback;
      if (pasteDialog.open) pasteDialog.close();
      pickingField = true;
      toolbar.style.pointerEvents = "none";
      document.body.append(pickerNotice);
      document.addEventListener("click", handleFieldPick, true);
      document.addEventListener("keydown", handlePickerKeydown, true);
    };

    const beginImport = async (sourceList, targetColumn = null) => {
      const table = getTargetTable();
      if (!table) {
        throw new Error(
          "لم يتم العثور على جدول رصد الدرجات. اعرض جدول الطلاب في نور أولاً.",
        );
      }
      if (
        targetColumn !== null &&
        (!Number.isInteger(targetColumn) || targetColumn < 0)
      ) {
        throw new Error("حدد مربع الدرجة المطلوب من جدول نور أولاً.");
      }

      input.disabled = true;
      input.value = "";
      pasteButton.disabled = true;
      try {
        sessionStorage.setItem(STORAGE_KEY_DATA, JSON.stringify(sourceList));
        sessionStorage.setItem(
          STORAGE_KEY_MODE,
          targetColumn === null ? "dual" : "single",
        );
        if (targetColumn !== null) {
          sessionStorage.setItem(
            STORAGE_KEY_TARGET_COLUMN,
            String(targetColumn),
          );
        } else {
          sessionStorage.removeItem(STORAGE_KEY_TARGET_COLUMN);
        }
        sessionStorage.setItem(STORAGE_KEY_ACTIVE, "true");
        sessionStorage.setItem(STORAGE_KEY_STATS, "0");
        sessionStorage.setItem(STORAGE_KEY_TARGET_PAGE, "2");

        status.textContent = "جارٍ رصد الدرجات في الصفحة الحالية...";
        const sourceMap = new Map(
          sourceList.map((source) => [source.identity, source]),
        );
        const result = processCurrentPage(sourceMap, targetColumn);
        sessionStorage.setItem(STORAGE_KEY_STATS, String(result.filledCount));

        const nextButton = findNextPageButton(2);
        if (nextButton) {
          status.textContent = `تم رصد (${result.filledCount} طالب) في الصفحة الأولى، جاري الانتقال للصفحة 2...`;
          sessionStorage.setItem(STORAGE_KEY_TARGET_PAGE, "3");
          await new Promise((resolve) => setTimeout(resolve, 800));
          nextButton.click();
        } else {
          status.textContent = `✓ اكتمل الرصد! تم رصد درجات (${result.filledCount}) طالب. تذكر الضغط على «حفظ» في نظام نور.`;
          clearImportSession();
          input.disabled = false;
          pasteButton.disabled = false;
        }
      } catch (error) {
        clearImportSession();
        input.disabled = false;
        pasteButton.disabled = false;
        throw error;
      }
    };

    pasteButton.addEventListener("click", () => {
      setCollapsed(false);
      pasteFeedback.textContent = "";
      pasteDialog.showModal();
      pasteArea.focus();
    });

    pasteFromClipboardBtn.addEventListener("click", async () => {
      try {
        const clipText = await navigator.clipboard.readText();
        if (clipText && clipText.trim()) {
          pasteArea.value = clipText;
          pasteFeedback.style.color = "#06896b";
          pasteFeedback.textContent =
            "تم لصق البيانات من الحافظة بنجاح. حدد مربع الدرجة أو ابدأ الرصد.";
        } else {
          pasteFeedback.style.color = "#b91c1c";
          pasteFeedback.textContent = "الحافظة فارغة حالياً.";
        }
      } catch {
        pasteFeedback.style.color = "#b45309";
        pasteFeedback.textContent =
          "يرجى لصق البيانات يدوياً داخل المربع باستخدام Ctrl+V.";
        pasteArea.focus();
      }
    });

    closeDialogButton.addEventListener("click", () => pasteDialog.close());

    chooseFieldButton.addEventListener("click", () => {
      try {
        startFieldPicking();
      } catch (error) {
        pasteFeedback.style.color = "#b91c1c";
        pasteFeedback.textContent =
          error instanceof Error ? error.message : "تعذر اختيار مربع الدرجة.";
      }
    });

    cancelPickButton.addEventListener("click", cancelFieldPicking);

    startPasteButton.addEventListener("click", async () => {
      try {
        const sourceList = parsePastedGrades(pasteArea.value, {
          fillAbsentAsZero: absentCheckbox.checked,
        });
        if (selectedColumn === null) {
          throw new Error(
            "اضغط «تحديد مربع الدرجة في نور» واختر خانة الدرجة المطلوبة أولاً.",
          );
        }
        pasteDialog.close();
        await beginImport(sourceList, selectedColumn);
      } catch (error) {
        const message =
          error instanceof Error ? error.message : "حدث خطأ أثناء الرصد.";
        if (pasteDialog.open) {
          pasteFeedback.style.color = "#b91c1c";
          pasteFeedback.textContent = message;
        } else {
          status.textContent = message;
        }
      }
    });

    clearColButton.addEventListener("click", () => {
      try {
        status.textContent = "انقر على مربع الدرجة في العمود المراد تصفيره...";
        startFieldPicking((colIdx) => {
          if (
            !confirm(
              `هل أنت متأكد من تصفير (تفريغ) العمود رقم (${colIdx + 1}) لجميع طلاب هذه الصفحة؟`,
            )
          ) {
            status.textContent = "تم إلغاء التصفير.";
            return;
          }
          const res = clearSelectedColumnCurrentPage(colIdx);
          status.textContent = res.message;
        });
      } catch (error) {
        status.textContent =
          error instanceof Error ? error.message : "تعذر تصفير العمود.";
      }
    });

    resetButton.addEventListener("click", () => {
      clearImportSession();
      input.disabled = false;
      pasteButton.disabled = false;
      input.value = "";
      status.textContent = "تم إلغاء عملية الرصد وتفريغ الجلسة.";
    });

    if (isImportActive) {
      input.disabled = true;
      pasteButton.disabled = true;
      status.textContent = "جارٍ انتظار تحميل صفحة نور واستكمال الرصد...";
      waitForTableAndResume();
    }

    input.addEventListener("change", async () => {
      const file = input.files?.[0];
      if (!file) return;

      setCollapsed(false);

      const table = getTargetTable();
      if (!table) {
        status.textContent =
          "تنبيه: لم يتم العثور على جدول الدرجات في الصفحة. اختر الصف والمادة واعرض جدول الطلاب أولاً.";
        input.value = "";
        return;
      }

      input.disabled = true;
      pasteButton.disabled = true;
      status.textContent = "جارٍ قراءة ملف إكزامكس وتحليل الدرجات...";

      try {
        if (file.size > MAX_FILE_SIZE) {
          throw new Error("حجم الملف أكبر من المسموح (25 ميغابايت).");
        }

        const rows = await readFileData(file);
        const totalGradeColumns = findTotalGradeColumns(rows);
        if (totalGradeColumns) {
          const sourceList = parseTotalGradeRows(rows, totalGradeColumns, {
            fillAbsentAsZero: absentCheckbox.checked,
          });
          status.textContent = `تمت قراءة (${sourceList.length} طالب). انقر الآن على مربع الدرجة المطلوب في جدول نور.`;
          startFieldPicking((targetColumn) => {
            beginImport(sourceList, targetColumn).catch((error) => {
              status.textContent =
                error instanceof Error
                  ? error.message
                  : "حدث خطأ أثناء رصد الدرجات.";
              input.disabled = false;
              pasteButton.disabled = false;
            });
          });
          return;
        }

        const sourceList = parseDualSourceRows(rows);
        await beginImport(sourceList);
      } catch (error) {
        status.textContent =
          error instanceof Error ? error.message : "حدث خطأ أثناء قراءة الملف.";
        clearImportSession();
        input.disabled = false;
        pasteButton.disabled = false;
      }
    });
  }

  function waitForTableAndResume() {
    let attempts = 0;
    const interval = setInterval(() => {
      attempts++;
      const table = getTargetTable();
      const hasRows = table && table.querySelectorAll("tr").length > 2;

      if (hasRows || attempts > 25) {
        clearInterval(interval);
        resumeImportProcess();
      }
    }, 300);
  }

  async function resumeImportProcess() {
    const toolbar = document.getElementById(TOOLBAR_ID);
    const status = toolbar?.querySelector("p");
    const input = toolbar?.querySelector("input[type='file']");
    const pasteButton = document.getElementById(PASTE_BUTTON_ID);

    try {
      const rawData =
        sessionStorage.getItem(STORAGE_KEY_DATA) ||
        localStorage.getItem(STORAGE_KEY_DATA);
      if (!rawData) {
        clearImportSession();
        if (input) input.disabled = false;
        if (pasteButton) pasteButton.disabled = false;
        return;
      }

      const sourceList = JSON.parse(rawData);
      const sourceMap = new Map(sourceList.map((s) => [s.identity, s]));
      const mode =
        sessionStorage.getItem(STORAGE_KEY_MODE) ||
        localStorage.getItem(STORAGE_KEY_MODE) ||
        "single";
      const selectedColumn =
        mode === "single"
          ? Number(
              sessionStorage.getItem(STORAGE_KEY_TARGET_COLUMN) ||
                localStorage.getItem(STORAGE_KEY_TARGET_COLUMN),
            )
          : null;
      if (mode === "single" && !Number.isInteger(selectedColumn)) {
        throw new Error(
          "تعذر استعادة عمود الدرجة المحدد بعد الانتقال بين الصفحات.",
        );
      }
      let totalFilled = Number(
        sessionStorage.getItem(STORAGE_KEY_STATS) ||
          localStorage.getItem(STORAGE_KEY_STATS) ||
          "0",
      );
      let targetPageNum = Number(
        sessionStorage.getItem(STORAGE_KEY_TARGET_PAGE) ||
          localStorage.getItem(STORAGE_KEY_TARGET_PAGE) ||
          "2",
      );

      const result = processCurrentPage(sourceMap, selectedColumn);
      totalFilled += result.filledCount;
      sessionStorage.setItem(STORAGE_KEY_STATS, String(totalFilled));

      const nextButton = findNextPageButton(targetPageNum);
      if (nextButton) {
        if (status) {
          status.textContent = `تم رصد (${result.filledCount}) في هذه الصفحة. الانتقال للصفحة ${targetPageNum}... (الإجمالي: ${totalFilled})`;
        }
        sessionStorage.setItem(
          STORAGE_KEY_TARGET_PAGE,
          String(targetPageNum + 1),
        );
        await new Promise((resolve) => setTimeout(resolve, 1000));
        nextButton.click();
      } else {
        if (status) {
          status.textContent = `✓ اكتمل الرصد بنجاح! تم رصد درجات (${totalFilled}) طالب عبر جميع الصفحات. تذكر الضغط على «حفظ» في نظام نور.`;
        }
        clearImportSession();
        if (input) input.disabled = false;
        if (pasteButton) pasteButton.disabled = false;
      }
    } catch (error) {
      if (status) {
        status.textContent =
          error instanceof Error
            ? `توقفت المتابعة: ${error.message}`
            : "حدث خطأ أثناء المتابعة التلقائية.";
      }
      clearImportSession();
      if (input) input.disabled = false;
      if (pasteButton) pasteButton.disabled = false;
    }
  }

  addToolbar();
})();
