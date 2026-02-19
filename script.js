const form = document.querySelector("#survey");
const progressText = document.querySelector(".progress-text");
const progressBar = document.querySelector(".progress-bar");
const thanks = document.querySelector(".thanks");
const resetBtn = document.querySelector("#resetBtn");
const downloadBtn = document.querySelector("#downloadBtn");
const sectionNavList = document.querySelector("#sectionNavList");
const prevSectionBtn = document.querySelector("#prevSectionBtn");
const nextSectionBtn = document.querySelector("#nextSectionBtn");
const mobileSectionState = document.querySelector("#mobileSectionState");
const sections = Array.from(document.querySelectorAll(".form-section"));
let lastResponses = null;
const phoneRegex = /^[0-9]{10}$/;

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

updateConditionalVisibility();
initSectionNav();
updateProgress();
