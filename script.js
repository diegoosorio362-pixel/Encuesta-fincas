const form = document.querySelector("#survey");
const progressText = document.querySelector(".progress-text");
const progressBar = document.querySelector(".progress-bar");
const thanks = document.querySelector(".thanks");
const resetBtn = document.querySelector("#resetBtn");
const downloadBtn = document.querySelector("#downloadBtn");
let lastResponses = null;

const getCategory = (element) => {
  const section = element.closest(".form-section");
  return section ? section.dataset.category : "General";
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
  const elements = Array.from(form.elements).filter((el) => el.name);
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

const downloadExcel = (rows) => {
  if (!rows || rows.length === 0 || !window.XLSX) {
    return;
  }

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
  link.download = "cuestionario.xlsx";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
};

const updateProgress = () => {
  const inputs = form.querySelectorAll("input, select, textarea");
  const filled = Array.from(inputs).filter((input) => {
    if (input.type === "radio") {
      return form.querySelector(`input[name='${input.name}']:checked`);
    }
    if (input.type === "checkbox") {
      return form.querySelectorAll(`input[name='${input.name}']:checked`).length > 0;
    }
    return input.value.trim().length > 0;
  });

  const percent = Math.round((filled.length / inputs.length) * 100);
  progressText.textContent = `${percent}% completado`;
  progressBar.style.setProperty("--progress", `${percent}%`);
};

form.addEventListener("input", () => {
  updateProgress();
});

form.addEventListener("submit", (event) => {
  event.preventDefault();
  lastResponses = collectResponses();
  downloadExcel(lastResponses);
  form.reset();
  updateProgress();
  form.parentElement.hidden = true;
  thanks.hidden = false;
});

resetBtn.addEventListener("click", () => {
  thanks.hidden = true;
  form.parentElement.hidden = false;
  updateProgress();
});

downloadBtn.addEventListener("click", () => {
  downloadExcel(lastResponses);
});

updateProgress();
