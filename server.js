const express = require('express');

const fs = require('fs');

const path = require('path');

const PizZip = require('pizzip');

const app = express();

const file = path.join(__dirname, 'informes.json');

const wordTemplate = path.join(__dirname, 'templates', 'informe-supervision.docx');

app.use(express.json({ limit: '4mb' }));

app.use(express.static(path.join(__dirname, 'public')));

app.use('/vendor/docx-preview', express.static(path.join(__dirname, 'node_modules', 'docx-preview', 'dist')));

app.use('/vendor/jszip', express.static(path.join(__dirname, 'node_modules', 'jszip', 'dist')));

const read = () => fs.existsSync(file) ? JSON.parse(fs.readFileSync(file,'utf8')) : [];

const save = data => fs.writeFileSync(file, JSON.stringify(data,null,2));

const valid = x => x && typeof x.expediente === 'string' && x.expediente.trim()

 && typeof x.empresa === 'string' && x.empresa.trim();

const meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

const escapeXml = value => String(value || '').replace(/[<>&"']/g, character => ({
 '<': '&lt;',
 '>': '&gt;',
 '&': '&amp;',
 '"': '&quot;',
 "'": '&apos;'
})[character]);

function formatSpanishDate(value) {
 const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value || '');
 if (!match) return 'Sin fecha';
 const [, year, month, day] = match;
 const monthIndex = Number(month) - 1;
 if (monthIndex < 0 || monthIndex >= meses.length) return 'Sin fecha';
 return `${Number(day)} de ${meses[monthIndex]} de ${year}`;
}

function paragraphText(paragraph) {
 return [...paragraph.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
  .map(([, text]) => text)
  .join('')
  .replace(/&amp;/g, '&')
  .replace(/&lt;/g, '<')
  .replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'");
}

function replaceParagraphText(paragraph, value) {
 let inserted = false;
 return paragraph.replace(/(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g,
  (match, open, content, close) => {
   if (!inserted) {
    inserted = true;
    return `${open}${escapeXml(value)}${close}`;
   }
   return `${open}${close}`;
  });
}

function contentLines(value) {
 const lines = String(value || '').split(/\r?\n/)
  .map(line => line.trim())
  .filter(Boolean)
  .map(line => line.replace(/^\d+(?:\.\d+)*(?:[.)])?\s*/, ''))
  .filter(Boolean);
 return lines.length ? lines : ['Sin información.'];
}

function extraerTextoCelda(cell) {
 return [...cell.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)]
  .map(([paragraph]) => paragraphText(paragraph).trim())
  .filter(Boolean)
  .join('\n');
}

function registrosTablaFiscalizacionModelo() {
 const zip = new PizZip(fs.readFileSync(wordTemplate));
 const xml = zip.file('word/document.xml').asText();
 const table = [...xml.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>/g)]
  .find(([content]) => paragraphText(content).includes('MEDIO PROBATORIO'));
 if (!table) throw new Error('No se encontró la tabla de fiscalización en la plantilla Word.');
 const rows = [...table[0].matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)];
 return rows.slice(1).map(([row]) => {
  const cells = [...row.matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)]
   .map(([cell]) => extraerTextoCelda(cell));
  if (cells.length !== 4) {
   throw new Error('La tabla de fiscalización debe tener cuatro columnas.');
  }
  return {
   numero: cells[0],
   incumplimiento: cells[1],
   descripcion: cells[2],
   medioProbatorio: cells[3]
  };
 });
}

function validarRegistrosTablaFiscalizacion(value) {
 if (value === undefined) return null;
 if (!Array.isArray(value) || value.length < 1 || value.length > 20) {
  throw new Error('La tabla debe tener entre 1 y 20 registros.');
 }
 return value.map((record, index) => {
  const fields = ['numero', 'incumplimiento', 'descripcion', 'medioProbatorio'];
  if (!record || fields.some(field => typeof record[field] !== 'string' ||
   !record[field].trim() || record[field].length > 10000)) {
   throw new Error(`Completa los cuatro campos del registro ${index + 1} (máximo 10 000 caracteres por campo).`);
  }
  return Object.fromEntries(fields.map(field => [field, record[field].trim()]));
 });
}

function reemplazarTextoCelda(cell, value) {
 const paragraphs = [...cell.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)];
 if (!paragraphs.length) throw new Error('No se encontró texto editable en una celda de la tabla.');
 const first = paragraphs[0];
 const last = paragraphs[paragraphs.length - 1];
 const lines = String(value).split(/\r?\n/);
 const replacement = lines.map(line => replaceParagraphText(first[0], line)).join('');
 const endOffset = last.index + last[0].length;
 return cell.slice(0, first.index) + replacement + cell.slice(endOffset);
}

function reemplazarTablaFiscalizacion(xml, records) {
 if (!records) return xml;
 let found = false;
 const result = xml.replace(/<w:tbl\b[\s\S]*?<\/w:tbl>/g, table => {
  if (!paragraphText(table).includes('MEDIO PROBATORIO')) return table;
  found = true;
  const rows = [...table.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)];
  if (rows.length < 2) throw new Error('La tabla de fiscalización no tiene filas editables.');
  const templateRow = rows[1][0];
  const replacementRows = records.map(record => {
   let cellIndex = 0;
   const row = templateRow.replace(/<w:tc\b[\s\S]*?<\/w:tc>/g, cell => {
    const field = ['numero', 'incumplimiento', 'descripcion', 'medioProbatorio'][cellIndex++];
    if (!field) throw new Error('La fila de fiscalización contiene más de cuatro columnas.');
    return reemplazarTextoCelda(cell, record[field]);
   });
   if (cellIndex !== 4) throw new Error('La fila de fiscalización debe tener cuatro columnas.');
   return row;
  });
  const first = rows[1];
  const last = rows[rows.length - 1];
  const endOffset = last.index + last[0].length;
  return table.slice(0, first.index) + replacementRows.join('') + table.slice(endOffset);
 });
 if (!found) throw new Error('No se encontró la tabla de fiscalización en la plantilla Word.');
 return result;
}

function getJpegDimensions(buffer) {
 if (buffer.length < 4 || buffer[0] !== 0xff || buffer[1] !== 0xd8) {
  throw new Error('La imagen de descripción general no es un JPEG válido.');
 }

 const frameMarkers = new Set([
  0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7,
  0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf
 ]);
 let offset = 2;
 while (offset < buffer.length - 9) {
  if (buffer[offset] !== 0xff) {
   offset++;
   continue;
  }
  const marker = buffer[offset + 1];
  if (marker === 0xd8 || marker === 0xd9 || marker === 0x01 ||
   (marker >= 0xd0 && marker <= 0xd7)) {
   offset += 2;
   continue;
  }
  const segmentLength = buffer.readUInt16BE(offset + 2);
  if (segmentLength < 2 || offset + 2 + segmentLength > buffer.length) break;
  if (frameMarkers.has(marker)) {
   return {
    width: buffer.readUInt16BE(offset + 7),
    height: buffer.readUInt16BE(offset + 5)
   };
  }
  offset += 2 + segmentLength;
 }
 throw new Error('No se pudieron leer las dimensiones de la imagen de descripción general.');
}

function imagenDescripcion(informe) {
 if (!informe.imagenDescripcionGeneral) return null;
 const match = /^data:image\/jpeg;base64,([A-Za-z0-9+/]+={0,2})$/
  .exec(informe.imagenDescripcionGeneral);
 if (!match) throw new Error('La imagen de descripción general no tiene un formato válido.');
 const buffer = Buffer.from(match[1], 'base64');
 if (buffer.length > 2 * 1024 * 1024) {
  throw new Error('La imagen de descripción general supera el límite de 2 MB.');
 }
 const dimensions = getJpegDimensions(buffer);
 return { buffer, ...dimensions };
}

function ajustarImagenDescripcion(xml, relationId, image) {
 let found = false;
 const maxWidth = 5303520;
 const maxHeight = 6000000;
 const scale = Math.min(maxWidth / image.width, maxHeight / image.height);
 const width = Math.round(image.width * scale);
 const height = Math.round(image.height * scale);
 const resized = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, paragraph => {
  if (!paragraph.includes(`r:embed="${relationId}"`)) return paragraph;
  found = true;
  return paragraph.replace(/<(?:wp:extent|a:ext)\b[^>]*\/>/g, extent =>
   extent.replace(/\bcx="\d+"/, `cx="${width}"`)
    .replace(/\bcy="\d+"/, `cy="${height}"`));
 });
 if (!found) throw new Error('No se encontró la imagen de la sección Descripción general.');
 return resized;
}

function replaceParagraphRange(xml, paragraphs, firstIndex, endIndex, replacement) {
 const first = paragraphs[firstIndex];
 const end = paragraphs[endIndex];
 const startOffset = first.index;
 const endOffset = end ? end.index : xml.length;
 return xml.slice(0, startOffset) + replacement + xml.slice(endOffset);
}

function replaceSectionBody(xml, heading, nextHeading, textLines) {
 const paragraphs = [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)];
 const headingIndex = paragraphs.findIndex(match =>
  paragraphText(match[0]).toUpperCase().includes(heading));
 const nextIndex = paragraphs.findIndex((match, index) =>
  index > headingIndex && paragraphText(match[0]).toUpperCase().includes(nextHeading));
 if (headingIndex < 0 || nextIndex < 0 || nextIndex <= headingIndex + 1) {
  throw new Error(`No se encontró la sección ${heading} en la plantilla Word.`);
 }

 const template = paragraphs[headingIndex + 1][0];
 const replacements = textLines.map(line => replaceParagraphText(template, line));
 return replaceParagraphRange(xml, paragraphs, headingIndex + 1, nextIndex, replacements.join(''));
}

function replaceNumberedBody(xml, heading, introText, nextHeading, textLines) {
 const paragraphs = [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)];
 const headingIndex = paragraphs.findIndex(match =>
  paragraphText(match[0]).toUpperCase().includes(heading));
 const introIndex = paragraphs.findIndex((match, index) =>
  index > headingIndex && paragraphText(match[0]).toLowerCase().includes(introText));
 const nextIndex = paragraphs.findIndex((match, index) =>
  index > introIndex && paragraphText(match[0]).toUpperCase().includes(nextHeading));
 if (headingIndex < 0 || introIndex < 0 || nextIndex <= introIndex + 1) {
  throw new Error(`No se encontró el contenido de ${heading} en la plantilla Word.`);
 }

 const template = paragraphs[introIndex + 1][0];
 const replacements = textLines.map(line => replaceParagraphText(template, line));
 const footnoteRuns = paragraphs.slice(introIndex + 1, nextIndex)
  .flatMap(match => [...match[0].matchAll(/<w:r\b[\s\S]*?<\/w:r>/g)]
   .map(([run]) => run)
   .filter(run => run.includes('<w:footnoteReference')));
 if (footnoteRuns.length) {
  const lastReplacement = replacements.length - 1;
  replacements[lastReplacement] = replacements[lastReplacement].replace('</w:p>',
   `${footnoteRuns.join('')}</w:p>`);
 }
 return replaceParagraphRange(xml, paragraphs, introIndex + 1, nextIndex, replacements.join(''));
}

function updateMergeField(xml, fieldName, value) {
 const fieldPattern = /(<w:fldChar\b[^>]*w:fldCharType=["']begin["'][^>]*>)([\s\S]*?)(<w:fldChar\b[^>]*w:fldCharType=["']separate["'][^>]*>)([\s\S]*?)(<w:fldChar\b[^>]*w:fldCharType=["']end["'][^>]*>)/g;

 return xml.replace(fieldPattern, (match, begin, instruction, separate, result, end) => {
  const field = new RegExp(`MERGEFIELD\\s+${fieldName}\\b`, 'i');
  if (!field.test(instruction)) return match;

  let inserted = false;
  result = result.replace(/(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g,
   (text, open, content, close) => {
    if (!inserted) {
     inserted = true;
     return `${open}${escapeXml(value)}${close}`;
    }
    return `${open}${close}`;
   });

  return `${begin}${instruction}${separate}${result}${end}`;
 });
}

function generarWord(informe) {
 if (!fs.existsSync(wordTemplate)) {
  throw new Error(`No se encontró la plantilla Word: ${wordTemplate}`);
 }

 const zip = new PizZip(fs.readFileSync(wordTemplate));
 const tablaFiscalizacion = validarRegistrosTablaFiscalizacion(informe.tablaFiscalizacion);
 const customImage = imagenDescripcion(informe);
 if (customImage) zip.file('word/media/image1.jpg', customImage.buffer, { binary: true });
 const descriptionText = String(informe.descripcionGeneral || '').trim() ||
  `El medio de transporte terrestre de combustibles líquidos y OPDH, operado por ${informe.empresa}; cuenta con la siguiente información:`;
 const replacements = [
  ['202600000000', informe.expediente],
  ['TRANSPORTES MODELO DEMOSTRATIVO S.A.C.', informe.empresa],
  ['ABC-000', informe.placa || 'No especificada'],
  ['Carlos Ejemplo Rivera', informe.supervisor || 'No especificado'],
  ['00000-060-000000', informe.registroHidrocarburos || 'No especificado'],
  ['20000000001', informe.ruc || 'No especificado'],
  ['29 de agosto de 2024', formatSpanishDate(informe.fecha2 || '2024-08-29')],
  ['10 de septiembre de 2026', formatSpanishDate(informe.fecha1 || '2026-09-10')],
  ['11 de septiembre de 2026', formatSpanishDate(informe.fecha3 || '2026-09-11')],
  ['18 de septiembre de 2026', formatSpanishDate(informe.fechaAsignacion || '2026-09-18')],
  ['24 de septiembre de 2026', formatSpanishDate(informe.fechaVisita || '2026-09-24')],
  ['22 de octubre de 2026', formatSpanishDate(informe.fechaConsultaRegistro || '2026-10-22')],
  ['25 de octubre de 2026', formatSpanishDate(informe.fechaInforme || '2026-10-25')]
 ];

 zip.file(/^word\/.*\.xml$/).forEach(entry => {
  let xml = entry.asText();
  replacements.forEach(([sample, value]) => {
   xml = xml.split(sample).join(escapeXml(value));
  });
  xml = updateMergeField(xml, 'osinumero', informe.expediente);
  if (entry.name === 'word/document.xml') {
   xml = reemplazarTablaFiscalizacion(xml, tablaFiscalizacion);
   if (informe.antecedentes) {
    xml = replaceSectionBody(xml, 'ANTECEDENTES:', 'BASE LEGAL:',
     contentLines(informe.antecedentes));
   }
   xml = replaceSectionBody(xml, 'DESCRIPCION GENERAL:', 'IMAGEN N° 01',
    contentLines(descriptionText));
   if (customImage) xml = ajustarImagenDescripcion(xml, 'rId8', customImage);
   xml = replaceNumberedBody(xml, 'CONCLUSIONES:', 'se concluye lo siguiente',
    'RECOMENDACIONES:', contentLines(informe.conclusiones));
   xml = replaceNumberedBody(xml, 'RECOMENDACIONES:', 'se recomienda:',
    'ATENTAMENTE.', contentLines(informe.recomendaciones));
   xml = xml.replace(/(?:<w:p\b[^>]*\/>\s*)+(?=<w:sectPr\b)/, '');
  }
  zip.file(entry.name, xml);
 });

 return zip.generate({
  type: 'nodebuffer',
  compression: 'DEFLATE',
  compressionOptions: { level: 6 }
 });
}

function responderWord(res, informe) {
 try {
  const document = generarWord(informe);
  const expediente = String(informe.expediente).replace(/[^a-zA-Z0-9._-]/g, '_');
  res.type('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.attachment(`Informe_${expediente}.docx`);
  res.send(document);
 } catch (error) {
  console.error('No se pudo generar el documento Word:', error);
  res.status(500).json({ error: 'No se pudo generar el documento Word.' });
 }
}

app.get('/api/informes', (req,res) => res.json(read()));

app.get('/api/plantilla/tabla-fiscalizacion', (req, res) => {
 try {
  res.json({ registros: registrosTablaFiscalizacionModelo() });
 } catch (error) {
  console.error('No se pudieron cargar los registros de la tabla de fiscalización:', error);
  res.status(500).json({ error: 'No se pudieron cargar los registros de la tabla de fiscalización.' });
 }
});

app.post('/api/informes/word', (req, res) => {
 if (!valid(req.body)) return res.status(400).json({ error: 'Expediente y empresa obligatorios' });
 try {
  imagenDescripcion(req.body);
  validarRegistrosTablaFiscalizacion(req.body.tablaFiscalizacion);
 } catch (error) {
  return res.status(400).json({ error: error.message });
 }
 responderWord(res, req.body);
});

app.get('/api/informes/:id/word', (req, res) => {
 const informe = read().find(x => x.id === req.params.id);
 if (!informe) return res.status(404).json({ error: 'No encontrado' });
 responderWord(res, informe);
});

app.get('/api/informes/:id', (req,res) => {

 const x=read().find(a=>a.id===req.params.id);

 return x ? res.json(x) : res.status(404).json({error:'No encontrado'});

});

app.post('/api/informes', (req,res) => {

 if(!valid(req.body)) return res.status(400).json({error:'Expediente y empresa obligatorios'});
 try {
  imagenDescripcion(req.body);
  validarRegistrosTablaFiscalizacion(req.body.tablaFiscalizacion);
 } catch (error) {
  return res.status(400).json({error:error.message});
 }

 const data=read(); const x={...req.body,id:String(Date.now())};

 data.push(x); save(data); res.status(201).json(x);

});

app.put('/api/informes/:id', (req,res) => {

 if(!valid(req.body)) return res.status(400).json({error:'Datos incompletos'});
 try {
  imagenDescripcion(req.body);
  validarRegistrosTablaFiscalizacion(req.body.tablaFiscalizacion);
 } catch (error) {
  return res.status(400).json({error:error.message});
 }

 const data=read(); const i=data.findIndex(a=>a.id===req.params.id);

 if(i<0) return res.status(404).json({error:'No encontrado'});

 data[i]={...req.body,id:req.params.id};save(data);res.json(data[i]);

});

app.delete('/api/informes/:id', (req,res) => {

 const data=read(); const next=data.filter(a=>a.id!==req.params.id);

 if(next.length===data.length) return res.status(404).json({error:'No encontrado'});

 save(next);res.status(204).end();

});

const port = process.env.PORT || 3000;

app.listen(port,()=>console.log(`Servidor: http://localhost:${port}`));