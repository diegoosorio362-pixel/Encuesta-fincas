const form = document.querySelector("#survey");
const progressText = document.querySelector(".progress-text");
const progressBar = document.querySelector(".progress-bar");
const thanks = document.querySelector(".thanks");
const resetBtn = document.querySelector("#resetBtn");
const downloadBtn = document.querySelector("#downloadBtn");
const downloadPdfBtn = document.querySelector("#downloadPdfBtn");
const autofillBtn = document.querySelector("#autofillBtn");
const sectionNavList = document.querySelector("#sectionNavList");
const prevSectionBtn = document.querySelector("#prevSectionBtn");
const nextSectionBtn = document.querySelector("#nextSectionBtn");
const mobileSectionState = document.querySelector("#mobileSectionState");
const sections = Array.from(document.querySelectorAll(".form-section"));
const brandLogoImg = document.querySelector(".hero-brand img");
let lastResponses = null;
let lastExportMeta = null;
const phoneRegex = /^[0-9]{10}$/;
let cachedLogoDataUrl = null;
const embeddedLogoDataUrl = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAASwAAAEsCAMAAABOo35HAAAA+VBMVEVHcEz////////////////////////8/Pv5+Pf////////////////l8Ob////////////////////////////BwcH///////////////////////////////////8KdQ0iIiLoExHw8vDn6echISFvb28sLCzW2Nbf4t/Ozs7vY2M3Nzcbfh6ZmZn96emRkZHpHh2ioqLrMC/GxsZISEjtSEeqqqr1nZz60tJ2dnZOTk6JiYlQnFJVVVVAQEC5ubmwsLCDg4O/v798fHzygoH4uLhqampaWloziza1tbXB28JlZWVmqGddXV1iYmKozahfX194snqRwJJoiNnnAAAAHnRSTlMAMya+4RfM/v0Kdz5exrFplKiLgqL2nERWUU5MSki5PfnFAAAncUlEQVR42uya227iOhSGOYQsBxAUigrtaCz11heWZcmyH8B+/1faGDs1jk0y7lAKnf3BcEg8lfJrHf4sM7oXmtlkvNj/OhyWy/WR5fJweNsvxpNZM/qfVqLZZLGcP63qujqCEABgbJ9HELLH6nr1Ml/uJ/+2bM14+byZ1ggwdvrgPlA13WzX439QsMnv9a4G/Bmqzfx18q9INpscttMKAf4ccAqy+mU5no1+OLO3p2mFrwI6CvZz9Zot5jXg61Jvf/08wZrZ722NMJxlE2RTDNyZ8CUQPkdH693hR5WwZv9U4xwQPbPkT8dHqt3hh8jV7LdV5zqHgZ7DkDuCXl4fPh+bybpGSYpdDCXovp6cKbUoC7UgBJARFtXPj23CFrvKXkwQZjgPIbwgqozUmhAhBPcIIYjWxigKOAFtXh9Vrtm67kkx6Mk+OKlEBH+/CONEG0VTyzqfjB6OZrytwcdLmnyQaYkhnCQRnLH3QRgX2lDo5HW1e7RsHD9VAyUbcKZ+ISUJP8rklRqGMaEVJNm4Hz0MzWKVeKLhTketUPb6Ez24OEIcQvBu1DEhKXT+9PRBvESz31Tdau3fIT4SQNRowVkqEtHSKBqjlNHEapbKFQRD07cHsBLjHcpGEFxshnBSqlONvEreP+Ti0LhA9AiJcAysfo/um/EGhhWK3JFPvgAXts35vtAPVVqwD7kMwh1W+ztOxsk2TsCkgHfPUlvPI6GIjBxUqlns2oCaD62ZpmGB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";

const conditionalRules = [
  { controller: "hay_viviendas", value: "si", targets: ["cuantas_viviendas"] },
  { controller: "fuentes_agua_cerca", value: "si", targets: ["cuantas_fuentes_agua"] },
  {
    controller: "estudios_suelo_fruto_agua",
    value: "si",
    targets: ["estudios_cuales", "estudios_parametro"]
  },
  { controller: "sombrio_transitorio", value: "si", targets: ["sombrio_transitorio_porcentaje"] },
  { controller: "sombrio_permanente", value: "si", targets: ["sombrio_permanente_porcentaje"] },
  { controller: "cultivo_transitorio", value: "si", targets: ["cultivo_transitorio_detalle"] },
  { controller: "vegetacion_alta", value: "si", targets: ["vegetacion_alta_detalle"] },
  { controller: "vegetacion_baja", value: "si", targets: ["vegetacion_baja_detalle"] },
  {
    controller: "otros_cultivos",
    value: "si",
    targets: ["otros_cultivos_cual", "otros_cultivos_porcentaje", "otros_cultivos_cercanos"]
  },
  { controller: "enfermedades_fruto", value: "otro", type: "checkbox", targets: ["enfermedades_fruto_otro"] }
];

const getCategory = (element) => {
  const section = element.closest(".form-section");
  return section ? section.dataset.category : "General";
};

const getFieldContainer = (element) =>
  element.closest(".field") || element.closest("fieldset") || element.closest("label");

const getTargetNodesByName = (name) => {
  const elements = Array.from(form.querySelectorAll(`[name='${name}']`));
  const nodes = [];

  elements.forEach((element) => {
    if (element.type === "radio") {
      const container = getFieldContainer(element);
      if (container) {
        nodes.push(container);
      }
      return;
    }

    if (element.type === "checkbox") {
      nodes.push(element.closest("label") || element);
      return;
    }

    nodes.push(element.closest("label") || element);
  });

  return [...new Set(nodes)];
};

const clearInputValue = (input) => {
  if (input.type === "radio" || input.type === "checkbox") {
    input.checked = false;
    return;
  }
  input.value = "";
};

const clearErrors = () => {
  form.querySelectorAll(".invalid").forEach((el) => el.classList.remove("invalid"));
  form.querySelectorAll(".error-text").forEach((el) => el.remove());
};

const clearSectionErrors = (section) => {
  if (!section) {
    return;
  }
  section.querySelectorAll(".invalid").forEach((el) => el.classList.remove("invalid"));
  section.querySelectorAll(".error-text").forEach((el) => el.remove());
};

const setFieldError = (container, message) => {
  if (!container) {
    return;
  }
  container.classList.add("invalid");
  if (!container.querySelector(".error-text")) {
    const error = document.createElement("p");
    error.className = "error-text";
    error.textContent = message;
    container.appendChild(error);
  }
};

const toggleTargets = (targetNames, show) => {
  targetNames.forEach((targetName) => {
    const nodes = getTargetNodesByName(targetName);
    nodes.forEach((node) => {
      node.hidden = !show;
      const elements = node.matches("input, select, textarea")
        ? [node]
        : Array.from(node.querySelectorAll("input, select, textarea"));

      elements.forEach((element) => {
        element.disabled = !show;
        if (!show) {
          clearInputValue(element);
        }
      });

      if (!show) {
        node.classList.remove("invalid");
        node.querySelectorAll(".error-text").forEach((el) => el.remove());
      }
    });
  });
};

const updateConditionalVisibility = () => {
  conditionalRules.forEach((rule) => {
    let show = false;
    if (rule.type === "checkbox") {
      show = !!form.querySelector(`input[name='${rule.controller}'][value='${rule.value}']:checked`);
    } else {
      const checked = form.querySelector(`input[name='${rule.controller}']:checked`);
      show = !!checked && checked.value === rule.value;
    }
    toggleTargets(rule.targets, show);
  });
};

const getQuestionLabel = (element) => {
  const label = element.closest("label");
  if (label) {
    const span = label.querySelector("span");
    if (span) {
      return span.textContent.trim();
    }
  }

  const fieldset = element.closest("fieldset");
  if (fieldset) {
    const legend = fieldset.querySelector("legend");
    if (legend) {
      return legend.textContent.trim();
    }
  }

  if (element.placeholder) {
    return element.placeholder.trim();
  }

  return element.name;
};

const collectResponses = () => {
  const elements = Array.from(form.elements).filter((el) => el.name && !el.disabled);
  const handled = new Set();
  const rows = [];

  elements.forEach((element) => {
    if (handled.has(element.name)) {
      return;
    }

    let value = "";
    if (element.type === "radio") {
      const checked = form.querySelector(`input[name='${element.name}']:checked`);
      value = checked ? checked.value : "";
    } else if (element.type === "checkbox") {
      const checked = Array.from(
        form.querySelectorAll(`input[name='${element.name}']:checked`)
      ).map((input) => input.value);
      value = checked.join(", ");
    } else {
      value = element.value.trim();
    }

    handled.add(element.name);
    rows.push([getCategory(element), getQuestionLabel(element), value]);
  });

  return rows;
};

const sanitizeFilePart = (value) => {
  if (!value) {
    return "sin-dato";
  }
  return value
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 40);
};

const getDateStamp = () => {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};

const buildExportMeta = () => {
  const finca = form.querySelector("[name='unidad_productora']")?.value?.trim() || "sin-finca";
  return {
    finca,
    dateStamp: getDateStamp()
  };
};

const buildPdfFileName = (meta) => {
  const fincaPart = sanitizeFilePart(meta?.finca);
  const datePart = meta?.dateStamp || getDateStamp();
  return `encuesta-${fincaPart}-${datePart}.pdf`;
};

const buildExcelFileName = (meta) => {
  const fincaPart = sanitizeFilePart(meta?.finca);
  const datePart = meta?.dateStamp || getDateStamp();
  return `encuesta-${fincaPart}-${datePart}.xlsx`;
};

const loadLogoDataUrl = async (cropHeader = false) => {
  if (cachedLogoDataUrl) {
    return cachedLogoDataUrl;
  }
  if (embeddedLogoDataUrl) {
    cachedLogoDataUrl = embeddedLogoDataUrl;
    return cachedLogoDataUrl;
  }
  try {
    if (brandLogoImg && brandLogoImg.src) {
      const img = brandLogoImg;
      if (!img.complete) {
        await new Promise((resolve) => {
          img.addEventListener("load", resolve, { once: true });
          img.addEventListener("error", resolve, { once: true });
        });
      }
      if (img.naturalWidth > 0 && img.naturalHeight > 0) {
        const canvas = document.createElement("canvas");
        if (cropHeader) {
          canvas.width = Math.floor(img.naturalWidth * 0.82);
          canvas.height = Math.floor(img.naturalHeight * 0.55);
        } else {
          canvas.width = img.naturalWidth;
          canvas.height = img.naturalHeight;
        }
        const ctx = canvas.getContext("2d");
        if (ctx) {
          if (cropHeader) {
            const sx = Math.floor(img.naturalWidth * 0.09);
            const sy = Math.floor(img.naturalHeight * 0.06);
            const sw = Math.floor(img.naturalWidth * 0.82);
            const sh = Math.floor(img.naturalHeight * 0.55);
            ctx.drawImage(img, sx, sy, sw, sh, 0, 0, canvas.width, canvas.height);
            const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            const data = imageData.data;
            for (let i = 0; i < data.length; i += 4) {
              const r = data[i];
              const g = data[i + 1];
              const b = data[i + 2];
              if (r > 242 && g > 242 && b > 242) {
                data[i + 3] = 0;
              }
            }
            ctx.putImageData(imageData, 0, 0);
          } else {
            ctx.drawImage(img, 0, 0);
          }
          cachedLogoDataUrl = canvas.toDataURL("image/png");
          return cachedLogoDataUrl;
        }
      }
    }

    const candidates = [
      "Universidad-de-la-amazonia-300x300.png",
      "./Universidad-de-la-amazonia-300x300.png"
    ];

    for (const path of candidates) {
      const response = await fetch(path);
      if (!response.ok) {
        continue;
      }
      const blob = await response.blob();
      cachedLogoDataUrl = await new Promise((resolve) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.readAsDataURL(blob);
      });
      if (cachedLogoDataUrl) {
        return cachedLogoDataUrl;
      }
    }
    return null;
  } catch {
    return null;
  }
};

const downloadPdf = async (rows, meta) => {
  if (!rows || rows.length === 0 || !window.jspdf?.jsPDF) {
    return;
  }

  const doc = new window.jspdf.jsPDF({ unit: "mm", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 12;
  const contentWidth = pageWidth - marginX * 2;
  const leftColWidth = (contentWidth - 6) / 2;
  const rightColX = marginX + leftColWidth + 6;
  const fileName = buildPdfFileName(meta || lastExportMeta || buildExportMeta());
  const logoData = await loadLogoDataUrl(false);
  const headerMain = [34, 64, 45];
  const headerDark = [219, 229, 221];
  const borderColor = [191, 201, 194];
  const categoryFill = [238, 243, 239];
  const rowAltFill = [247, 250, 247];

  const drawHeader = () => {
    doc.setFillColor(255, 255, 255);
    doc.rect(0, 0, pageWidth, 40, "F");
    doc.setFillColor(...headerDark);
    doc.rect(0, 0, pageWidth, 3.2, "F");

    if (logoData) {
      doc.addImage(logoData, "PNG", marginX + 1, 5.9, 24, 24);
    } else {
      doc.setDrawColor(...headerMain);
      doc.setLineWidth(0.5);
      doc.circle(marginX + 15, 18, 10, "S");
      doc.setTextColor(...headerMain);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(10);
      doc.text("UA", marginX + 15, 19, { align: "center" });
    }

    const panelX = pageWidth - marginX - 58;
    const textCenterX = (44 + panelX - 4) / 2;

    doc.setTextColor(...headerMain);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.text("UNIVERSIDAD DE LA AMAZONIA", textCenterX, 13.2, { align: "center" });
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9.2);
    doc.text("Formato institucional de encuesta de unidad productora", textCenterX, 18.9, {
      align: "center"
    });
    doc.setFontSize(8);
    doc.text("Vigilada Ministerio de Educacion Nacional", textCenterX, 23.9, {
      align: "center"
    });

    doc.setDrawColor(203, 214, 205);
    doc.setLineWidth(0.35);
    doc.line(panelX, 9.8, panelX, 21.6);

    doc.setTextColor(96, 113, 102);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(6.9);
    doc.text("Fecha de registro", panelX + 4, 11.2);
    doc.text("Unidad productora", panelX + 4, 14.9);

    doc.setTextColor(...headerMain);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.4);
    doc.text(meta?.dateStamp || getDateStamp(), pageWidth - marginX, 11.2, {
      align: "right"
    });
    doc.text(meta?.finca || "Sin finca", pageWidth - marginX, 14.9, {
      align: "right"
    });

    doc.setDrawColor(199, 209, 201);
    doc.setLineWidth(0.35);
    doc.line(marginX, 29.2, pageWidth - marginX, 29.2);
  };

  let y = 33;
  drawHeader();
  doc.setDrawColor(...borderColor);
  doc.setLineWidth(0.22);

  let currentCategory = "";
  let questionNumber = 1;

  const ensureSpace = (heightNeeded) => {
    if (y + heightNeeded <= pageHeight - 15) {
      return;
    }
    doc.addPage();
    drawHeader();
    y = 33;
  };

  rows.forEach(([category, question, answer]) => {
    if (category !== currentCategory) {
      currentCategory = category;
      ensureSpace(18);
      doc.setFillColor(...categoryFill);
      doc.rect(marginX, y, contentWidth, 6.8, "F");
      doc.setFont("helvetica", "bold");
      doc.setFontSize(8.8);
      doc.setTextColor(...headerMain);
      doc.text(category.toUpperCase(), marginX + 2, y + 4.5);
      y += 8.2;

      doc.setFillColor(242, 246, 243);
      doc.rect(marginX, y, leftColWidth, 5.6, "F");
      doc.rect(rightColX, y, leftColWidth, 5.6, "F");
      doc.setTextColor(48, 63, 52);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(8);
      doc.text("Pregunta", marginX + 2, y + 3.8);
      doc.text("Respuesta", rightColX + 2, y + 3.8);
      y += 6.6;
    }

    const cleanedQuestion = question.replace(/^\d+\.\s*/, "");
    const questionText = `${questionNumber}. ${cleanedQuestion}`;
    const answerText = answer && answer.trim() ? answer : "-";
    const questionLines = doc.splitTextToSize(questionText, leftColWidth - 4);
    const answerLines = doc.splitTextToSize(answerText, leftColWidth - 4);
    const lineCount = Math.max(questionLines.length, answerLines.length);
    const rowHeight = Math.max(7.5, lineCount * 4.3 + 2.8);

    ensureSpace(rowHeight + 0.7);

    if (questionNumber % 2 === 0) {
      doc.setFillColor(...rowAltFill);
      doc.rect(marginX, y, leftColWidth, rowHeight, "F");
      doc.rect(rightColX, y, leftColWidth, rowHeight, "F");
    }

    doc.setDrawColor(...borderColor);
    doc.rect(marginX, y, leftColWidth, rowHeight);
    doc.rect(rightColX, y, leftColWidth, rowHeight);

    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.2);
    doc.setTextColor(36, 47, 40);
    doc.text(questionLines, marginX + 2, y + 4.5);

    doc.setFont("helvetica", "normal");
    doc.setTextColor(48, 58, 50);
    doc.text(answerLines, rightColX + 2, y + 4.5);

    y += rowHeight + 0.7;
    questionNumber += 1;
  });

  const totalPages = doc.getNumberOfPages();
  for (let page = 1; page <= totalPages; page += 1) {
    doc.setPage(page);
    doc.setTextColor(110, 122, 113);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.text(`Pagina ${page} de ${totalPages}`, pageWidth - marginX, pageHeight - 6, {
      align: "right"
    });
  }

  doc.save(fileName);
};

const buildStyledSheet = (rows) => {
  const sheet = XLSX.utils.aoa_to_sheet([]);
  sheet["!cols"] = [{ wch: 6 }, { wch: 68 }, { wch: 68 }, { wch: 20 }];
  sheet["!merges"] = [];

  const headerFill = { patternType: "solid", fgColor: { rgb: "1F2D3D" } };
  const titleFill = { patternType: "solid", fgColor: { rgb: "E8EEF3" } };
  const categoryFill = { patternType: "solid", fgColor: { rgb: "D4DEE8" } };
  const zebraFill = { patternType: "solid", fgColor: { rgb: "F7F9FB" } };
  const borderAll = {
    top: { style: "thin", color: { rgb: "6B7C93" } },
    bottom: { style: "thin", color: { rgb: "6B7C93" } },
    left: { style: "thin", color: { rgb: "6B7C93" } },
    right: { style: "thin", color: { rgb: "6B7C93" } }
  };

  let rowIndex = 0;
  const addRow = (values, style) => {
    XLSX.utils.sheet_add_aoa(sheet, [values], { origin: { r: rowIndex, c: 0 } });
    values.forEach((_, colIndex) => {
      const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      const cell = sheet[cellAddress];
      if (!cell) {
        return;
      }
      cell.s = { ...style, border: borderAll };
    });
    rowIndex += 1;
  };

  const mergeRow = (startCol, endCol) => {
    sheet["!merges"].push({
      s: { r: rowIndex, c: startCol },
      e: { r: rowIndex, c: endCol }
    });
  };

  addRow(["", "FORMATO DE ENCUESTA - CACAO", "", ""], {
    font: { bold: true, color: { rgb: "FFFFFF" }, sz: 14 },
    fill: headerFill,
    alignment: { horizontal: "center", vertical: "center" }
  });
  mergeRow(1, 2);
  rowIndex += 1;

  addRow(["", "IDENTIFICACION", "", ""], {
    font: { bold: true, color: { rgb: "FFFFFF" } },
    fill: headerFill,
    alignment: { horizontal: "center", vertical: "center" }
  });
  mergeRow(1, 2);
  rowIndex += 1;

  addRow(["Fecha:", "", "Encuestador:", ""], {
    font: { bold: true },
    alignment: { vertical: "center" }
  });
  addRow(["Finca:", "", "Municipio:", ""], {
    font: { bold: true },
    alignment: { vertical: "center" }
  });
  addRow(["Contacto:", "", "Vereda:", ""], {
    font: { bold: true },
    alignment: { vertical: "center" }
  });
  rowIndex += 1;

  addRow(["No.", "Pregunta", "Respuesta", "Observaciones"], {
    font: { bold: true },
    fill: titleFill,
    alignment: { horizontal: "center", vertical: "center", wrapText: true }
  });

  let questionNumber = 1;
  let currentCategory = "";
  rows.forEach(([category, question, answer]) => {
    if (category !== currentCategory) {
      currentCategory = category;
      addRow([currentCategory, "", "", ""], {
        font: { bold: true },
        fill: categoryFill,
        alignment: { vertical: "center" }
      });
      mergeRow(0, 3);
      rowIndex += 1;
    }

    const rowStyle = {
      alignment: { vertical: "top", wrapText: true }
    };
    if (questionNumber % 2 === 0) {
      rowStyle.fill = zebraFill;
    }
    addRow([questionNumber, question, answer || "", ""], rowStyle);
    questionNumber += 1;
  });

  rowIndex += 1;
  addRow(["OBSERVACIONES DEL ENTREVISTADOR", "", "", ""], {
    font: { bold: true },
    fill: titleFill,
    alignment: { vertical: "center" }
  });
  mergeRow(0, 3);
  rowIndex += 1;
  addRow(["", "", "", ""], { alignment: { vertical: "top" } });
  addRow(["", "", "", ""], { alignment: { vertical: "top" } });

  rowIndex += 1;
  addRow(["FIRMA DEL ENTREVISTADOR", "", "FIRMA DEL PRODUCTOR", ""], {
    font: { bold: true },
    alignment: { vertical: "center" }
  });

  return sheet;
};

const downloadExcel = async (rows, meta) => {
  if (!rows || rows.length === 0) {
    return;
  }

  const exportMeta = meta || lastExportMeta || buildExportMeta();
  const fileName = buildExcelFileName(exportMeta);

  if (window.ExcelJS) {
    const workbook = new window.ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Encuesta");
    const logoData = await loadLogoDataUrl(false);

    sheet.columns = [
      { key: "a", width: 6 },
      { key: "b", width: 56 },
      { key: "c", width: 56 },
      { key: "d", width: 30 }
    ];

    const headerMain = "22402D";
    const headerSoft = "DBE5DD";
    const border = { style: "thin", color: { argb: "FFC4CEC5" } };

    const paintRow = (idx, fill) => {
      [1, 2, 3, 4].forEach((col) => {
        const cell = sheet.getCell(idx, col);
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${fill}` } };
      });
    };

    paintRow(1, headerSoft);
    paintRow(2, "FFFFFF");
    paintRow(3, "FFFFFF");
    paintRow(4, "FFFFFF");

    if (logoData) {
      const logoId = workbook.addImage({ base64: logoData, extension: "png" });
      sheet.addImage(logoId, {
        tl: { col: 0.1, row: 0.8 },
        ext: { width: 88, height: 88 }
      });
    }

    sheet.mergeCells("B2:C2");
    sheet.getCell("B2").value = "UNIVERSIDAD DE LA AMAZONIA";
    sheet.getCell("B2").font = { name: "Calibri", bold: true, size: 18, color: { argb: `FF${headerMain}` } };
    sheet.getCell("B2").alignment = { horizontal: "center", vertical: "middle" };

    sheet.mergeCells("B3:C3");
    sheet.getCell("B3").value = "Formato institucional de encuesta de unidad productora";
    sheet.getCell("B3").font = { name: "Calibri", size: 12, color: { argb: `FF${headerMain}` } };
    sheet.getCell("B3").alignment = { horizontal: "center", vertical: "middle" };

    sheet.mergeCells("B4:C4");
    sheet.getCell("B4").value = "Vigilada Ministerio de Educacion Nacional";
    sheet.getCell("B4").font = { name: "Calibri", size: 10, color: { argb: "FF4F6354" } };
    sheet.getCell("B4").alignment = { horizontal: "center", vertical: "middle" };

    sheet.getCell("D2").value = `Fecha: ${exportMeta.dateStamp}`;
    sheet.getCell("D3").value = `Finca: ${exportMeta.finca}`;
    sheet.getCell("D2").font = { name: "Calibri", bold: true, size: 11, color: { argb: `FF${headerMain}` } };
    sheet.getCell("D3").font = { name: "Calibri", bold: true, size: 11, color: { argb: `FF${headerMain}` } };
    sheet.getCell("D2").alignment = { horizontal: "right", vertical: "middle" };
    sheet.getCell("D3").alignment = { horizontal: "right", vertical: "middle" };

    sheet.getRow(6).values = ["No.", "Pregunta", "Respuesta", "Observaciones"];
    sheet.getRow(6).font = { name: "Calibri", bold: true, size: 11, color: { argb: `FF${headerMain}` } };
    sheet.getRow(6).alignment = { horizontal: "center", vertical: "middle" };
    [1, 2, 3, 4].forEach((col) => {
      sheet.getCell(6, col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF4F0" } };
      sheet.getCell(6, col).border = { top: border, left: border, right: border, bottom: border };
    });

    let row = 7;
    let currentCategory = "";
    let n = 1;
    rows.forEach(([category, question, answer]) => {
      if (category !== currentCategory) {
        currentCategory = category;
        sheet.mergeCells(`A${row}:D${row}`);
        const c = sheet.getCell(`A${row}`);
        c.value = category;
        c.font = { name: "Calibri", bold: true, size: 11, color: { argb: `FF${headerMain}` } };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE9F1EA" } };
        c.alignment = { vertical: "middle" };
        c.border = { top: border, left: border, right: border, bottom: border };
        row += 1;
      }

      const cleanedQuestion = (question || "").replace(/^\d+\.\s*/, "");
      sheet.getCell(`A${row}`).value = n;
      sheet.getCell(`B${row}`).value = cleanedQuestion;
      sheet.getCell(`C${row}`).value = answer || "-";
      sheet.getCell(`D${row}`).value = "";
      [1, 2, 3, 4].forEach((col) => {
        const cell = sheet.getCell(row, col);
        cell.alignment = { vertical: "top", wrapText: true };
        cell.border = { top: border, left: border, right: border, bottom: border };
        if (n % 2 === 0) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FBF8" } };
        }
      });
      row += 1;
      n += 1;
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    return;
  }

  if (window.XLSX) {
    const worksheet = buildStyledSheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Encuesta");
    const arrayBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([arrayBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }
};

const updateProgress = () => {
  const inputs = Array.from(form.querySelectorAll("input, select, textarea")).filter(
    (input) => !input.disabled
  );

  if (inputs.length === 0) {
    progressText.textContent = "0% completado";
    progressBar.style.setProperty("--progress", "0%");
    return;
  }

  const filled = Array.from(inputs).filter((input) => {
    if (input.type === "radio") {
      return inputs.some((item) => item.name === input.name && item.checked);
    }
    if (input.type === "checkbox") {
      return inputs.some((item) => item.name === input.name && item.checked);
    }
    return input.value.trim().length > 0;
  });

  const percent = Math.round((filled.length / inputs.length) * 100);
  progressText.textContent = `${percent}% completado`;
  progressBar.style.setProperty("--progress", `${percent}%`);
  updateSectionProgress();
};

const sectionStats = (section) => {
  const sectionInputs = Array.from(section.querySelectorAll("input, select, textarea")).filter(
    (input) => !input.disabled
  );

  if (sectionInputs.length === 0) {
    return { percent: 0, filled: 0, total: 0 };
  }

  const filled = sectionInputs.filter((input) => {
    if (input.type === "radio" || input.type === "checkbox") {
      return sectionInputs.some((item) => item.name === input.name && item.checked);
    }
    return input.value.trim().length > 0;
  });

  const percent = Math.round((filled.length / sectionInputs.length) * 100);
  return { percent, filled: filled.length, total: sectionInputs.length };
};

const updateSectionProgress = () => {
  sections.forEach((section) => {
    const id = section.dataset.sectionId;
    const item = sectionNavList?.querySelector(`[data-nav='${id}']`);
    if (!item) {
      return;
    }

    const stats = sectionStats(section);
    const badge = item.querySelector(".section-nav-progress");
    if (badge) {
      badge.textContent = `${stats.percent}%`;
    }
    item.classList.toggle("done", stats.percent === 100);
  });
};

const setActiveSection = (sectionId) => {
  if (!sectionNavList) {
    return;
  }
  sectionNavList.querySelectorAll(".section-nav-item").forEach((node) => {
    node.classList.toggle("active", node.dataset.nav === sectionId);
  });
  updateMobileSectionNav(sectionId);
};

const updateMobileSectionNav = (sectionId) => {
  if (!sections.length) {
    return;
  }
  const index = sections.findIndex((section) => section.dataset.sectionId === sectionId);
  const current = index >= 0 ? index : 0;

  if (mobileSectionState) {
    mobileSectionState.textContent = `Seccion ${current + 1} de ${sections.length}`;
  }
  if (prevSectionBtn) {
    prevSectionBtn.disabled = current === 0;
  }
  if (nextSectionBtn) {
    nextSectionBtn.disabled = current === sections.length - 1;
  }
};

const activateSection = (sectionId, shouldScroll = false) => {
  const targetSection = sections.find((section) => section.dataset.sectionId === sectionId);
  if (!targetSection) {
    return;
  }

  sections.forEach((section) => {
    section.hidden = section.dataset.sectionId !== sectionId;
  });

  setActiveSection(sectionId);
  if (shouldScroll) {
    targetSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }
};

const initSectionNav = () => {
  if (!sectionNavList || sections.length === 0) {
    return;
  }

  sections.forEach((section, index) => {
    const safeId = `section-${index + 1}`;
    section.dataset.sectionId = safeId;
    section.id = safeId;

    const li = document.createElement("li");
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "section-nav-item";
    btn.dataset.nav = safeId;
    btn.innerHTML = `<span class="section-nav-label">${section.dataset.category}</span><span class="section-nav-progress">0%</span>`;
    btn.addEventListener("click", () => {
      activateSection(safeId, true);
    });
    li.appendChild(btn);
    sectionNavList.appendChild(li);
  });
  activateSection(sections[0].dataset.sectionId);
  updateSectionProgress();
};

const validateForm = () => {
  clearErrors();
  let isValid = true;
  const visibleElements = Array.from(form.elements).filter(
    (element) => element.name && !element.disabled
  );
  const handledRadio = new Set();

  visibleElements.forEach((element) => {
    if (element.type === "radio") {
      if (handledRadio.has(element.name)) {
        return;
      }
      handledRadio.add(element.name);
      const group = visibleElements.filter((item) => item.type === "radio" && item.name === element.name);
      const groupRequired = group.some((item) => item.required);
      const checked = group.some((item) => item.checked);
      if (groupRequired && !checked) {
        setFieldError(getFieldContainer(element), "Selecciona una opcion.");
        isValid = false;
      }
      return;
    }

    if (!element.required && !element.value.trim()) {
      return;
    }

    const container = getFieldContainer(element);
    if (element.required && !element.value.trim()) {
      setFieldError(container, "Este campo es obligatorio.");
      isValid = false;
      return;
    }

    if (element.type === "tel") {
      const normalized = element.value.replace(/\D/g, "");
      if (!phoneRegex.test(normalized)) {
        setFieldError(container, "Ingresa un celular valido de 10 digitos.");
        isValid = false;
      }
      return;
    }

    if (element.type === "email" && !element.checkValidity()) {
      setFieldError(container, "Ingresa un correo electronico valido.");
      isValid = false;
    }
  });

  return isValid;
};

const validateSection = (section) => {
  if (!section) {
    return true;
  }

  clearSectionErrors(section);
  let isValid = true;
  const elements = Array.from(section.querySelectorAll("input, select, textarea")).filter(
    (element) => element.name && !element.disabled
  );
  const handledRadio = new Set();

  elements.forEach((element) => {
    if (element.type === "radio") {
      if (handledRadio.has(element.name)) {
        return;
      }
      handledRadio.add(element.name);
      const group = elements.filter((item) => item.type === "radio" && item.name === element.name);
      const groupRequired = group.some((item) => item.required);
      const checked = group.some((item) => item.checked);
      if (groupRequired && !checked) {
        setFieldError(getFieldContainer(element), "Selecciona una opcion.");
        isValid = false;
      }
      return;
    }

    if (!element.required && !element.value.trim()) {
      return;
    }

    const container = getFieldContainer(element);
    if (element.required && !element.value.trim()) {
      setFieldError(container, "Este campo es obligatorio.");
      isValid = false;
      return;
    }

    if (element.type === "tel") {
      const normalized = element.value.replace(/\D/g, "");
      if (!phoneRegex.test(normalized)) {
        setFieldError(container, "Ingresa un celular valido de 10 digitos.");
        isValid = false;
      }
      return;
    }

    if (element.type === "email" && !element.checkValidity()) {
      setFieldError(container, "Ingresa un correo electronico valido.");
      isValid = false;
    }
  });

  return isValid;
};

const sampleValueFor = (element) => {
  if (element.type === "email") {
    return "prueba@udla.edu.co";
  }
  if (element.type === "tel") {
    return "3001234567";
  }
  if (element.type === "number") {
    return "10";
  }

  const name = element.name || "";
  if (name.includes("coordenadas")) {
    return "1.2345, -75.6789";
  }
  if (name.includes("temperatura")) {
    return "24 C";
  }
  if (name.includes("presion")) {
    return "760 mmHg";
  }
  if (name.includes("altitud")) {
    return "450 msnm";
  }
  if (name.includes("distancia")) {
    return "5 km";
  }
  if (name.includes("unidad_productora")) {
    return "Finca Demo";
  }
  if (name.includes("municipio")) {
    return "Florencia";
  }
  if (name.includes("departamento")) {
    return "Caqueta";
  }
  if (name.includes("vereda")) {
    return "La Esperanza";
  }
  if (name.includes("observaciones")) {
    return "Dato de prueba para verificar exportacion PDF.";
  }
  return "Dato temporal de prueba";
};

const fillBasicInputs = () => {
  const textLikeInputs = Array.from(form.querySelectorAll("input, textarea, select")).filter(
    (el) => !el.disabled && el.type !== "radio" && el.type !== "checkbox"
  );

  textLikeInputs.forEach((element) => {
    if (element.tagName === "SELECT") {
      if (element.options.length > 1) {
        element.selectedIndex = 1;
      }
      return;
    }
    element.value = sampleValueFor(element);
  });
};

const fillRadioGroups = () => {
  const radioNames = [...new Set(Array.from(form.querySelectorAll("input[type='radio']")).map((r) => r.name))];
  radioNames.forEach((name) => {
    const group = Array.from(form.querySelectorAll(`input[type='radio'][name='${name}']`)).filter(
      (item) => !item.disabled
    );
    if (!group.length) {
      return;
    }
    const preferred = group.find((item) => item.value === "si") || group[0];
    preferred.checked = true;
  });
};

const fillCheckboxGroups = () => {
  const checkboxNames = [...new Set(Array.from(form.querySelectorAll("input[type='checkbox']")).map((c) => c.name))];
  checkboxNames.forEach((name) => {
    const group = Array.from(form.querySelectorAll(`input[type='checkbox'][name='${name}']`)).filter(
      (item) => !item.disabled
    );
    if (!group.length) {
      return;
    }
    group[0].checked = true;
    const other = group.find((item) => item.value === "otro");
    if (other) {
      other.checked = true;
    }
  });
};

const autofillSurvey = () => {
  form.reset();
  clearErrors();

  fillRadioGroups();
  fillCheckboxGroups();
  updateConditionalVisibility();

  fillBasicInputs();
  fillRadioGroups();
  fillCheckboxGroups();
  updateConditionalVisibility();
  fillBasicInputs();

  if (sections[0] && sections[0].dataset.sectionId) {
    activateSection(sections[0].dataset.sectionId);
  }
  updateProgress();
};

form.addEventListener("input", (event) => {
  if (event.target.type === "tel") {
    event.target.value = event.target.value.replace(/\D/g, "").slice(0, 10);
  }
  const container = getFieldContainer(event.target);
  if (container) {
    container.classList.remove("invalid");
    container.querySelectorAll(".error-text").forEach((el) => el.remove());
  }
  updateConditionalVisibility();
  updateProgress();
});

form.addEventListener("change", () => {
  updateConditionalVisibility();
  updateProgress();
});

form.addEventListener("submit", (event) => {
  event.preventDefault();
  updateConditionalVisibility();
  if (!validateForm()) {
    const firstError = form.querySelector(".invalid");
    if (firstError) {
      const targetSection = firstError.closest(".form-section");
      if (targetSection && targetSection.dataset.sectionId) {
        activateSection(targetSection.dataset.sectionId);
      }
      firstError.scrollIntoView({ behavior: "smooth", block: "center" });
    }
    return;
  }
  lastResponses = collectResponses();
  lastExportMeta = buildExportMeta();
  downloadExcel(lastResponses);
  form.reset();
  clearErrors();
  updateConditionalVisibility();
  updateProgress();
  form.parentElement.hidden = true;
  thanks.hidden = false;
});

resetBtn.addEventListener("click", () => {
  thanks.hidden = true;
  form.parentElement.hidden = false;
  clearErrors();
  updateConditionalVisibility();
  if (sections[0] && sections[0].dataset.sectionId) {
    activateSection(sections[0].dataset.sectionId);
  }
  updateProgress();
});

downloadBtn.addEventListener("click", () => {
  downloadExcel(lastResponses);
});

if (downloadPdfBtn) {
  downloadPdfBtn.addEventListener("click", async () => {
    await downloadPdf(lastResponses, lastExportMeta);
  });
}

if (prevSectionBtn) {
  prevSectionBtn.addEventListener("click", () => {
    const activeSection = sections.find((section) => !section.hidden);
    const index = activeSection ? sections.indexOf(activeSection) : 0;
    const prevIndex = Math.max(0, index - 1);
    activateSection(sections[prevIndex].dataset.sectionId, true);
  });
}

if (nextSectionBtn) {
  nextSectionBtn.addEventListener("click", () => {
    updateConditionalVisibility();
    const activeSection = sections.find((section) => !section.hidden);
    const index = activeSection ? sections.indexOf(activeSection) : 0;
    if (!validateSection(activeSection)) {
      const firstError = activeSection?.querySelector(".invalid");
      if (firstError) {
        firstError.scrollIntoView({ behavior: "smooth", block: "center" });
      }
      return;
    }
    const nextIndex = Math.min(sections.length - 1, index + 1);
    activateSection(sections[nextIndex].dataset.sectionId, true);
  });
}

if (autofillBtn) {
  autofillBtn.addEventListener("click", () => {
    autofillSurvey();
  });
}

updateConditionalVisibility();
initSectionNav();
updateProgress();
