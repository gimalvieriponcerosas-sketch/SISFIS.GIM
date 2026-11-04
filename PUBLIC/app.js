const ids = [
  "expediente",
  "empresa",
  "placa",
  "supervisor",
  "registroHidrocarburos",
  "ruc",
  "fecha1",
  "fecha2",
  "fecha3",
  "fechaAsignacion",
  "fechaVisita",
  "fechaConsultaRegistro",
  "fechaInforme",
  "antecedentes",
  "descripcionGeneral",
  "conclusiones",
  "recomendaciones",
];

let editId = null;
let previewTimer = null;
let previewRequest = 0;
let imageRequest = 0;
let imagenDescripcionGeneral = "";
let descripcionGeneralEditada = false;
let registrosModelo = [];
let registrosFiscalizacion = [];
let tablaFiscalizacionLista = false;
const textoEditado = { conclusiones: false, recomendaciones: false };
const meses = [
  "enero",
  "febrero",
  "marzo",
  "abril",
  "mayo",
  "junio",
  "julio",
  "agosto",
  "septiembre",
  "octubre",
  "noviembre",
  "diciembre",
];
const fechasIniciales = {
  fecha1: "2026-09-10",
  fecha2: "2024-08-29",
  fecha3: "2026-09-11",
  fechaAsignacion: "2026-09-18",
  fechaVisita: "2026-09-24",
  fechaConsultaRegistro: "2026-10-22",
};
const valoresAntecedentes = {
  expediente: "[Número de expediente]",
  empresa: "[Agente fiscalizado]",
  placa: "[placa principal]",
  supervisor: "GIM ALVIERI PONCE ROSAS",
  registroHidrocarburos: "[N.° de Registro de Hidrocarburos]",
  ruc: "[R.U.C.]",
};
const valoresAnteriores = {};

const $ = (id) => document.getElementById(id);

function fechaEnEspanol(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value || "");
  if (!match) return "[Fecha pendiente]";
  const [, year, month, day] = match;
  return `${Number(day)} de ${meses[Number(month) - 1]} de ${year}`;
}

function hoyISO() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
}

function plantillaDescripcionGeneral() {
  const empresa = $("empresa").value.trim() || "[Agente fiscalizado]";
  return `El medio de transporte terrestre de combustibles líquidos y OPDH, operado por ${empresa}; cuenta con la siguiente información:`;
}

function mostrarImagenDescripcion(dataUrl) {
  const image = $("imagenDescripcionPreview");
  image.hidden = !dataUrl;
  image.removeAttribute("src");
  if (dataUrl) image.src = dataUrl;
  $("restaurarImagenDescripcion").hidden = !dataUrl;
}

const camposTablaFiscalizacion = [
  { id: "numero", label: "N.º / Observación", rows: 2 },
  {
    id: "incumplimiento",
    label: "Incumplimiento normativo – base legal",
    rows: 6,
  },
  {
    id: "descripcion",
    label: "Descripción del incumplimiento verificado",
    rows: 5,
  },
  { id: "medioProbatorio", label: "Medio probatorio", rows: 3 },
];

function registroFiscalizacionVacio() {
  return {
    numero: "",
    incumplimiento: "",
    descripcion: "",
    medioProbatorio: "",
  };
}

function renderizarTablaFiscalizacion() {
  const editor = $("tablaFiscalizacionEditor");
  editor.replaceChildren();
  registrosFiscalizacion.forEach((registro, index) => {
    const card = document.createElement("section");
    card.className = "observacion-editor";
    const heading = document.createElement("h4");
    heading.textContent = `Registro ${index + 1}`;
    card.append(heading);
    camposTablaFiscalizacion.forEach((campo) => {
      const label = document.createElement("label");
      label.htmlFor = `observacion-${index}-${campo.id}`;
      label.textContent = campo.label;
      const input = document.createElement(
        campo.id === "numero" ? "input" : "textarea",
      );
      input.id = label.htmlFor;
      input.name = `tablaFiscalizacion[${index}][${campo.id}]`;
      input.value = registro[campo.id] || "";
      input.required = true;
      input.maxLength = 10000;
      if (campo.id === "numero") input.type = "text";
      else input.rows = campo.rows;
      input.addEventListener("input", () => {
        registrosFiscalizacion[index][campo.id] = input.value;
      });
      card.append(label, input);
    });
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "btn btn-secundario";
    remove.textContent = "Quitar registro";
    remove.disabled = registrosFiscalizacion.length === 1;
    remove.addEventListener("click", () => {
      registrosFiscalizacion.splice(index, 1);
      renderizarTablaFiscalizacion();
      programarVistaPrevia();
    });
    card.append(remove);
    editor.append(card);
  });
  $("agregarObservacion").disabled =
    !tablaFiscalizacionLista || registrosFiscalizacion.length >= 20;
  if (tablaFiscalizacionLista) {
    $("tablaFiscalizacionEstado").textContent =
      `${registrosFiscalizacion.length} registro(s) en la tabla. Puedes editarlos, agregar filas o quitar las que no necesites.`;
  }
}

function configurarTablaFiscalizacion(registros) {
  registrosFiscalizacion = (registros.length ? registros : registrosModelo).map(
    (registro) => ({
      numero: registro.numero || "",
      incumplimiento: registro.incumplimiento || "",
      descripcion: registro.descripcion || "",
      medioProbatorio: registro.medioProbatorio || "",
    }),
  );
  renderizarTablaFiscalizacion();
}

async function cargarModeloTablaFiscalizacion() {
  const response = await fetch("/api/plantilla/tabla-fiscalizacion");
  if (!response.ok)
    throw new Error(
      "No se pudieron cargar los registros de la tabla de fiscalización.",
    );
  const result = await response.json();
  if (!Array.isArray(result.registros) || !result.registros.length) {
    throw new Error(
      "La plantilla no contiene registros para la tabla de fiscalización.",
    );
  }
  registrosModelo = result.registros;
  tablaFiscalizacionLista = true;
  configurarTablaFiscalizacion(registrosModelo);
  $("tablaFiscalizacionEstado").textContent =
    "Puedes editar los registros, agregar filas o quitar las que no necesites.";
  $("agregarObservacion").disabled = registrosFiscalizacion.length >= 20;
  ["descargarWord", "vistaPreviaWord"].forEach(
    (id) => ($(id).disabled = false),
  );
  $("form").querySelector('[type="submit"]').disabled = false;
}

function archivoComoDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () =>
      reject(new Error("No se pudo leer la imagen seleccionada."));
    reader.readAsDataURL(blob);
  });
}

async function convertirImagenDescripcion(file) {
  if (!["image/jpeg", "image/png", "image/webp"].includes(file.type)) {
    throw new Error("Selecciona una imagen JPG, PNG o WebP.");
  }
  if (file.size > 8 * 1024 * 1024) {
    throw new Error("La imagen original debe pesar menos de 8 MB.");
  }

  const objectUrl = URL.createObjectURL(file);
  try {
    const image = new Image();
    image.src = objectUrl;
    await image.decode();
    if (
      !image.naturalWidth ||
      !image.naturalHeight ||
      image.naturalWidth * image.naturalHeight > 25000000
    ) {
      throw new Error("La imagen tiene dimensiones no compatibles.");
    }
    const scale = Math.min(
      1,
      1800 / Math.max(image.naturalWidth, image.naturalHeight),
    );
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(image.naturalWidth * scale);
    canvas.height = Math.round(image.naturalHeight * scale);
    canvas.getContext("2d").drawImage(image, 0, 0, canvas.width, canvas.height);
    const jpeg = await new Promise((resolve) =>
      canvas.toBlob(resolve, "image/jpeg", 0.88),
    );
    if (!jpeg) throw new Error("No se pudo procesar la imagen seleccionada.");
    if (jpeg.size > 2 * 1024 * 1024) {
      throw new Error("La imagen procesada supera el límite de 2 MB.");
    }
    return archivoComoDataUrl(jpeg);
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function actualizarValoresAnteriores() {
  Object.keys(valoresAntecedentes).forEach((id) => {
    valoresAnteriores[id] = $(id).value.trim() || valoresAntecedentes[id];
  });
  Object.keys(fechasIniciales)
    .concat("fechaInforme")
    .forEach((id) => {
      valoresAnteriores[id] = fechaEnEspanol($(id).value);
    });
}

function plantillaAntecedentes() {
  const valor = (id) => $(id).value.trim() || valoresAntecedentes[id];
  return [
    `Con fecha ${fechaEnEspanol($("fecha2").value)}, ${valor("empresa")} (en adelante, el Agente Fiscalizado) obtuvo el Registro de Hidrocarburos N.° ${valor("registroHidrocarburos")}, para operar la unidad vehicular de placa principal N.° ${valor("placa")} y tracto XYZ-000, en la actividad de transporte de combustibles líquidos y OPDH, con capacidad autorizada de 9000 galones y seis compartimientos, con dirección legal en DIRECCIÓN FICTICIA PARA CAPACITACIÓN, HUÁNUCO, distrito de CHURUBAMBA, provincia y departamento de HUÁNUCO, conforme a la consulta del Registro de Hidrocarburos.`,
    `Mediante la Resolución de Autoridad de Fiscalización – Oficina Regional Huánuco N.° 000-2026-OS-GSE/DSR-OR HUÁNUCO, de fecha ${fechaEnEspanol($("fecha1").value)}, se dispuso, en calidad de MANDATO, que ${valor("empresa")} ponga a disposición la unidad de placa principal N.° ${valor("placa")}, con Registro de Hidrocarburos N.° ${valor("registroHidrocarburos")}.`,
    `Con fecha ${fechaEnEspanol($("fecha3").value)}, mediante carta dirigida al Especialista Regional en Hidrocarburos de la Oficina Regional Huánuco, ${valor("empresa")} comunicó que la fiscalización de la unidad ${valor("placa")} se realizaría en la Carretera Central Huánuco–Tingo María, Km 23+270, Centro Poblado de Chullqui, distrito de Churubamba, Huánuco.`,
    `Con fecha ${fechaEnEspanol($("fechaAsignacion").value)}, el Especialista Regional en Hidrocarburos, ${valor("supervisor")}, mediante Asignación N.° 00000-2026, asignó la Carta Línea N.° 0000000-0 a la empresa supervisora correspondiente, a fin de efectuar la fiscalización operativa a medios de transporte terrestre de CL, con visita, a la unidad vehicular de placa ${valor("placa")}.`,
    `Con fecha ${fechaEnEspanol($("fechaVisita").value)}, ${valor("supervisor")} realizó la visita de fiscalización al medio de transporte de placa principal N.° ${valor("placa")}, operado por ${valor("empresa")}, y suscribió el Acta de Fiscalización correspondiente al Expediente N.° ${valor("expediente")}.`,
    `Con fecha ${fechaEnEspanol($("fechaInforme").value)}, se elabora el presente informe de fiscalización, correspondiente a la Carta Línea N.° 0000000-0 y al Expediente N.° ${valor("expediente")}, para su evaluación.`,
  ].join("\n");
}

function iniciarDatosAntecedentes() {
  Object.entries(fechasIniciales).forEach(([id, value]) => {
    $(id).value = value;
  });
  $("fechaInforme").value = hoyISO();
  $("antecedentes").value = plantillaAntecedentes();
  $("descripcionGeneral").value = plantillaDescripcionGeneral();
  descripcionGeneralEditada = false;
  if (tablaFiscalizacionLista) configurarTablaFiscalizacion(registrosModelo);
  actualizarValoresAnteriores();
}

function actualizarAntecedentesPorCampo(id) {
  const actual = Object.prototype.hasOwnProperty.call(valoresAntecedentes, id)
    ? $(id).value.trim() || valoresAntecedentes[id]
    : fechaEnEspanol($(id).value);
  const anterior = valoresAnteriores[id];
  if (anterior && anterior !== actual) {
    $("antecedentes").value = $("antecedentes")
      .value.split(anterior)
      .join(actual);
  }
  valoresAnteriores[id] = actual;
}

function liberarRecursosVistaPrevia(container) {
  const urls = new Set(container.innerHTML.match(/blob:[^)"'\s]+/g) || []);
  urls.forEach((url) => URL.revokeObjectURL(url));
}

function ajustarEscalaWord(container) {
  const availableWidth = container.clientWidth - 24;
  if (availableWidth <= 0) return;
  container.querySelectorAll(".word-document").forEach((page) => {
    const pageWidth = page.offsetWidth;
    if (pageWidth > 0)
      page.style.zoom = String(Math.min(1, availableWidth / pageWidth));
  });
}

function centrarFirmaVistaPrevia(container) {
  const tables = container.querySelectorAll("table");
  const signatureTable = tables[tables.length - 1];
  const fiscalizerCell = signatureTable?.rows[0]?.cells[1];
  const image = fiscalizerCell?.querySelector("img");
  const anchoredImage = image?.parentElement;
  if (anchoredImage && image) {
    const widthInPoints =
      Number.parseFloat(image.style.width) ||
      (image.getBoundingClientRect().width * 72) / 96;
    anchoredImage.style.left = `${-widthInPoints / 2}pt`;
    anchoredImage.style.transform = "none";
  }
}

function plantillasTexto() {
  const empresa = $("empresa").value.trim() || "[Agente fiscalizado]";
  const registro =
    $("registroHidrocarburos").value.trim() ||
    "[N.° de Registro de Hidrocarburos]";
  const placa = $("placa").value.trim() || "[placa principal]";
  return {
    conclusiones: `6.1. ${empresa}, la Ficha de Registro de Hidrocarburos N.° ${registro}, NO CUMPLIÓ con lo establecido de las condiciones de seguridad de Medio De Transporte De Combustible Líquidos Y OPDH.`,
    recomendaciones: `7.1. Comunicar a ${empresa}, con Ficha de Registro de Hidrocarburos N.° ${registro}, las observaciones consignadas en el apartado II «Otros» del Acta de Fiscalización, a fin de que acredite su subsanación respecto de la unidad vehicular de placa principal ${placa}.\n7.2. Asimismo, se otorgue el plazo establecido de 30 días calendario para que la empresa pueda subsanar los incumplimientos de orden técnico y que no constituyen un alto riesgo para la seguridad (TIPO B), bajo apercibimiento de iniciar Procedimiento Administrativo Sancionador – PAS.`,
  };
}

function actualizarPlantillasTexto() {
  const plantillas = plantillasTexto();
  if (!textoEditado.conclusiones)
    $("conclusiones").value = plantillas.conclusiones;
  if (!textoEditado.recomendaciones)
    $("recomendaciones").value = plantillas.recomendaciones;
}

Object.keys(valoresAntecedentes)
  .concat(Object.keys(fechasIniciales), "fechaInforme")
  .forEach((id) => {
    const cambiar = () => {
      actualizarAntecedentesPorCampo(id);
      actualizarPlantillasTexto();
    };
    $(id).addEventListener("input", cambiar);
    $(id).addEventListener("change", cambiar);
  });

["conclusiones", "recomendaciones"].forEach((id) => {
  $(id).addEventListener("input", () => {
    textoEditado[id] = true;
  });
});

$("form").addEventListener("reset", () => {
  imageRequest++;
  imagenDescripcionGeneral = "";
  descripcionGeneralEditada = false;
  textoEditado.conclusiones = false;
  textoEditado.recomendaciones = false;
  setTimeout(() => {
    iniciarDatosAntecedentes();
    actualizarPlantillasTexto();
    mostrarImagenDescripcion("");
    $("imagenDescripcionEstado").textContent =
      "Se usará la imagen incluida en la plantilla.";
    programarVistaPrevia();
  });
});

iniciarDatosAntecedentes();
actualizarPlantillasTexto();

function datos() {
  const x = {};
  ids.forEach((k) => (x[k] = $(k).value));
  x.imagenDescripcionGeneral = imagenDescripcionGeneral;
  x.tablaFiscalizacion = registrosFiscalizacion.map((registro) => ({
    ...registro,
  }));
  return x;
}

function cargar(x) {
  ids.forEach((k) => ($(k).value = x[k] || ""));

  Object.entries(fechasIniciales).forEach(([id, value]) => {
    if (!x[id]) $(id).value = value;
  });
  if (!x.supervisor) $("supervisor").value = "GIM ALVIERI PONCE ROSAS";
  if (!x.fechaInforme) $("fechaInforme").value = hoyISO();
  if (!x.antecedentes) $("antecedentes").value = plantillaAntecedentes();
  if (!x.descripcionGeneral)
    $("descripcionGeneral").value = plantillaDescripcionGeneral();
  descripcionGeneralEditada =
    !!x.descripcionGeneral &&
    x.descripcionGeneral !== plantillaDescripcionGeneral();
  configurarTablaFiscalizacion(
    Array.isArray(x.tablaFiscalizacion) && x.tablaFiscalizacion.length
      ? x.tablaFiscalizacion
      : registrosModelo,
  );
  imagenDescripcionGeneral = x.imagenDescripcionGeneral || "";
  $("imagenDescripcionGeneral").value = "";
  mostrarImagenDescripcion(imagenDescripcionGeneral);
  $("imagenDescripcionEstado").textContent = imagenDescripcionGeneral
    ? "Imagen personalizada guardada con este informe."
    : "Se usará la imagen incluida en la plantilla.";
  actualizarValoresAnteriores();
  const plantillas = plantillasTexto();
  textoEditado.conclusiones =
    !!x.conclusiones && x.conclusiones !== plantillas.conclusiones;
  textoEditado.recomendaciones =
    !!x.recomendaciones && x.recomendaciones !== plantillas.recomendaciones;
  actualizarPlantillasTexto();
  editId = x.id;
  actualizarVistaPrevia(true);
}

async function listar() {
  const r = await fetch("/api/informes");
  const a = await r.json();

  $("lista").replaceChildren();
  a.forEach((x) => {
    const p = document.createElement("p");

    p.textContent = `${x.expediente} – ${x.empresa} `;

    const e = document.createElement("button");
    e.textContent = "Editar";
    e.onclick = () => cargar(x);

    const b = document.createElement("button");
    b.textContent = "Eliminar";

    b.onclick = async () => {
      if (confirm("¿Eliminar?")) {
        await fetch("/api/informes/" + x.id, { method: "DELETE" });
        listar();
      }
    };

    const v = document.createElement("button");
    v.textContent = "Vista previa Word";
    v.onclick = () =>
      mostrarWord(
        "/api/informes/" + encodeURIComponent(x.id) + "/word",
        {},
        true,
      );

    const d = document.createElement("button");
    d.textContent = "Descargar Word";
    d.onclick = () =>
      descargarWord(
        "/api/informes/" + encodeURIComponent(x.id) + "/word",
        {},
        x.expediente,
      );

    p.append(e, b, v, d);
    $("lista").append(p);
  });
}

async function cargarWord(url, options = {}) {
  const r = await fetch(url, options);
  if (!r.ok) {
    let message = "No se pudo generar la vista previa del documento Word.";
    try {
      const error = await r.json();
      if (error.error) message = error.error;
    } catch {}
    throw new Error(message);
  }
  return r.arrayBuffer();
}

async function prepararWordVistaPrevia(documentData) {
  const zip = await window.JSZip.loadAsync(documentData);
  const documentFile = zip.file("word/document.xml");
  if (!documentFile) {
    throw new Error("El documento Word no contiene el archivo principal.");
  }
  const xml = await documentFile.async("text");
  // docx-preview renders superscript runs twice, including their footnote references.
  const previewXml = xml.replace(/<w:r\b[\s\S]*?<\/w:r>/g, (run) =>
    run.includes("<w:footnoteReference")
      ? run.replace(/<w:vertAlign\b[^>]*\/>/g, "")
      : run,
  );
  if (previewXml === xml) return documentData;
  zip.file("word/document.xml", previewXml);
  return zip.generateAsync({ type: "arraybuffer" });
}

async function mostrarWord(
  url,
  options = {},
  scroll = false,
  readyMessage = "Vista previa actualizada.",
) {
  const target = $("wordPreview");
  const status = $("wordPreviewStatus");
  const request = ++previewRequest;
  status.textContent = "Actualizando vista previa...";
  let staging;
  try {
    const responseData = await cargarWord(url, options);
    const documentData = await prepararWordVistaPrevia(responseData);
    if (request !== previewRequest) return;
    if (!window.docx?.renderAsync) {
      throw new Error(
        "No se pudo cargar el visor de Word. Recargue la página.",
      );
    }
    staging = document.createElement("div");
    await window.docx.renderAsync(documentData, staging, staging, {
      className: "word-document",
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      breakPages: true,
      renderHeaders: true,
      renderFooters: true,
    });
    if (request !== previewRequest) {
      liberarRecursosVistaPrevia(staging);
      return;
    }
    centrarFirmaVistaPrevia(staging);
    liberarRecursosVistaPrevia(target);
    target.replaceChildren(...staging.childNodes);
    ajustarEscalaWord(target);
    status.textContent = readyMessage;
    if (scroll) $("vista-previa").scrollIntoView({ behavior: "smooth" });
  } catch (error) {
    if (staging) liberarRecursosVistaPrevia(staging);
    if (request === previewRequest) status.textContent = error.message;
  }
}

function actualizarVistaPrevia(scroll = false) {
  actualizarPlantillasTexto();
  const informe = datos();
  const expedienteCompleto = !!informe.expediente.trim();
  const empresaCompleta = !!informe.empresa.trim();
  if (!expedienteCompleto) informe.expediente = "[Número de expediente]";
  if (!empresaCompleta) informe.empresa = "[Agente fiscalizado]";
  const readyMessage =
    expedienteCompleto && empresaCompleta
      ? "Vista previa actualizada."
      : "Vista previa lista. Completa expediente y empresa en el formulario para reemplazar los datos de ejemplo.";
  mostrarWord(
    "/api/informes/word",
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(informe),
    },
    scroll,
    readyMessage,
  );
}

function programarVistaPrevia() {
  clearTimeout(previewTimer);
  previewTimer = setTimeout(() => actualizarVistaPrevia(), 500);
}

$("form").addEventListener("input", programarVistaPrevia);
$("form").addEventListener("change", programarVistaPrevia);

$("empresa").addEventListener("input", () => {
  if (!descripcionGeneralEditada)
    $("descripcionGeneral").value = plantillaDescripcionGeneral();
});

$("descripcionGeneral").addEventListener("input", () => {
  descripcionGeneralEditada = true;
});

$("imagenDescripcionGeneral").addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;
  const request = ++imageRequest;
  $("imagenDescripcionEstado").textContent = "Procesando imagen...";
  try {
    const dataUrl = await convertirImagenDescripcion(file);
    if (request !== imageRequest) return;
    imagenDescripcionGeneral = dataUrl;
    mostrarImagenDescripcion(dataUrl);
    $("imagenDescripcionEstado").textContent =
      "Imagen lista para la vista previa y el Word.";
    programarVistaPrevia();
  } catch (error) {
    if (request === imageRequest)
      $("imagenDescripcionEstado").textContent = error.message;
  }
});

$("restaurarImagenDescripcion").addEventListener("click", () => {
  imageRequest++;
  imagenDescripcionGeneral = "";
  $("imagenDescripcionGeneral").value = "";
  mostrarImagenDescripcion("");
  $("imagenDescripcionEstado").textContent =
    "Se usará la imagen incluida en la plantilla.";
  programarVistaPrevia();
});

$("agregarObservacion").addEventListener("click", () => {
  if (registrosFiscalizacion.length >= 20) return;
  registrosFiscalizacion.push(registroFiscalizacionVacio());
  renderizarTablaFiscalizacion();
  $("observacion-" + (registrosFiscalizacion.length - 1) + "-numero").focus();
});

new ResizeObserver(() => ajustarEscalaWord($("wordPreview"))).observe(
  $("wordPreview"),
);

async function descargarWord(url, options, expediente) {
  try {
    const r = await fetch(url, options);
    if (!r.ok) {
      let message = "No se pudo descargar el documento Word.";
      try {
        const error = await r.json();
        if (error.error) message = error.error;
      } catch {}
      throw new Error(message);
    }
    const blob = await r.blob();
    const link = document.createElement("a");
    const nombre = String(expediente || "informe").replace(
      /[^a-zA-Z0-9._-]/g,
      "_",
    );
    const objectUrl = URL.createObjectURL(blob);
    link.href = objectUrl;
    link.download = `Informe_${nombre}.docx`;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
  } catch (error) {
    alert(error.message);
  }
}

$("descargarWord").onclick = () => {
  if (!$("form").reportValidity()) return;
  actualizarPlantillasTexto();
  descargarWord(
    "/api/informes/word",
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(datos()),
    },
    $("expediente").value,
  );
};

$("vistaPreviaWord").onclick = () => {
  clearTimeout(previewTimer);
  actualizarVistaPrevia(true);
};

document
  .querySelector('a[href="#vista-previa"]')
  .addEventListener("click", (event) => {
    event.preventDefault();
    clearTimeout(previewTimer);
    actualizarVistaPrevia(true);
  });

$("form").onsubmit = async (ev) => {
  ev.preventDefault();
  const x = datos();
  if (!tablaFiscalizacionLista) {
    $("tablaFiscalizacionEstado").textContent =
      "La tabla todavía no está lista; vuelva a cargar la página para intentarlo de nuevo.";
    return;
  }

  const r = await fetch("/api/informes" + (editId ? "/" + editId : ""), {
    method: editId ? "PUT" : "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(x),
  });

  const result = await r.json();
  if (!r.ok) {
    alert(result.error || "No se pudo guardar el informe.");
    return;
  }

  editId = null;
  $("form").reset();

  listar();
};

cargarModeloTablaFiscalizacion()
  .then(() => {
    listar();
    actualizarVistaPrevia();
  })
  .catch((error) => {
    $("tablaFiscalizacionEstado").textContent = error.message;
  });
