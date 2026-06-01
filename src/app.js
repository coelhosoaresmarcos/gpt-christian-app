const REQUIRED_SHEETS = [
  'HOSPITAL_ORIGINAL',
  'CONTROLE DE MEDICAMENTOS',
  'ASSOCIACOES',
  'BAIXAR',
  'RELATORIO',
  'VALIDACAO',
];

const EXCLUDED_ITEMS = [
  'INFUSOR - 100 ML - 2 ML/H',
  'INFUSOR - 200ML - 4ML/H',
  'MATERIAL',
  'SOL. DE GLICOSE 5%',
  'EQUIPO MACRO SIMPLES',
  'SOL. FISIOLOGICA 0,9%',
  'EQUIPO DE BOMBA - SIMPLES',
  'EQUIPO DE BOMBA - FOTOSSENSIVEL',
  'INFUSOR - 270 ML - 5 ML/H',
  'FILTRO',
  'EQUIPO DE BOMBA - ISENTO DE PVC',
  'MANITOL - 20%',
  'EQUIPO DE BOMBA FREE FOTO',
  'EQUIPO MACRO FOTOSSENSIVEL',
  'INFUSOR - 250 ML - 5 ML/H',
  'EQUIPO MACRO ISENTO DE PVC COM FILTRO - 0,2 MICRAS',
  'INFUSOR - 150 ML - 3 ML/H',
  'INFUSOR - 250 ML - 2,5ML/H',
  'FRETE',
  'INFUSOR - 250 ML - 10 ML/H',
];

const hospitalInput = document.querySelector('#hospitalFile');
const controlInput = document.querySelector('#controlFile');
const hospitalFileName = document.querySelector('#hospitalFileName');
const controlFileName = document.querySelector('#controlFileName');
const hospitalNameInput = document.querySelector('#hospitalName');
const processButton = document.querySelector('#processButton');
const downloadButton = document.querySelector('#downloadButton');
const statusElement = document.querySelector('#status');
const resultCard = document.querySelector('#resultCard');
const summaryElement = document.querySelector('#summary');

let generatedBlob = null;
let generatedFileName = '';

hospitalInput.addEventListener('change', () => updateFileName(hospitalInput, hospitalFileName));
controlInput.addEventListener('change', () => updateFileName(controlInput, controlFileName));
processButton.addEventListener('click', processFiles);
downloadButton.addEventListener('click', downloadGeneratedFile);

function updateFileName(input, target) {
  target.textContent = input.files?.[0]?.name ? `📄 ${input.files[0].name}` : '📄 Selecionar Excel';
}

async function processFiles() {
  const hospitalFile = hospitalInput.files?.[0];
  const controlFile = controlInput.files?.[0];
  if (!hospitalFile || !controlFile) {
    setStatus('Envie o arquivo do hospital e a planilha SOBRAS ALTO CUSTO REDE AMÉRICAS.', 'error');
    return;
  }

  if (!window.ExcelJS) {
    setStatus('ExcelJS não foi carregado. Verifique a conexão com a CDN ou instale a dependência localmente.', 'error');
    return;
  }

  processButton.disabled = true;
  resultCard.classList.add('hidden');
  setStatus('Processando planilhas e aplicando as regras Christian...', 'info');

  try {
    const result = await processChristianWorkbook({
      hospitalFile,
      controlFile,
      hospitalName: hospitalNameInput.value.trim(),
    });
    generatedBlob = result.blob;
    generatedFileName = result.fileName;
    summaryElement.textContent = `${result.summary.associations} associações, ${result.summary.downloadRows} itens baixáveis e ${result.summary.validationRows} registros de validação.`;
    resultCard.classList.remove('hidden');
    setStatus('Arquivo consolidado gerado com sucesso.', 'success');
  } catch (error) {
    setStatus(error instanceof Error ? error.message : 'Erro inesperado ao processar as planilhas.', 'error');
  } finally {
    processButton.disabled = false;
  }
}

function downloadGeneratedFile() {
  if (!generatedBlob) return;
  const url = URL.createObjectURL(generatedBlob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = generatedFileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

async function processChristianWorkbook({ hospitalFile, controlFile, hospitalName }) {
  const hospitalWorkbook = await loadWorkbook(hospitalFile);
  const controlWorkbook = await loadWorkbook(controlFile);
  const hospitalSheet = hospitalWorkbook.worksheets[0];
  const controlSheet = controlWorkbook.getWorksheet('CONTROLE DE MEDICAMENTOS');

  if (!hospitalSheet) throw new Error('O arquivo do hospital não possui abas legíveis.');
  if (!controlSheet) throw new Error('A planilha de controle precisa conter a aba CONTROLE DE MEDICAMENTOS.');

  const validations = [];
  const hospitalRows = readHospitalRows(hospitalSheet, validations);
  const controlRows = readControlRows(controlSheet, validations);
  const associations = associateRows(hospitalRows, controlRows, validations, hospitalName);
  const outputWorkbook = buildOutputWorkbook(hospitalRows, controlRows, associations, validations, hospitalSheet);
  const buffer = await outputWorkbook.xlsx.writeBuffer();

  return {
    blob: new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    fileName: `christian_consolidado_${new Date().toISOString().slice(0, 10)}.xlsx`,
    summary: {
      associations: associations.length,
      downloadRows: hospitalRows.filter((row) => !row.__isExcluded).length,
      validationRows: validations.length,
    },
  };
}

async function loadWorkbook(file) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  return workbook;
}

function readHospitalRows(sheet, validations) {
  const { headers, startRow } = extractHeaders(sheet);
  const rows = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber <= startRow) return;
    const object = rowToObject(row, headers);
    if (isEmptyObject(object)) return;

    const hospitalRow = {
      ...object,
      __rowId: rowNumber,
      __data: getByCandidates(object, ['Data', 'Dt Atendimento', 'Data Atendimento'], 2),
      __cliente: String(getByCandidates(object, ['Cliente', 'Unidade', 'Hospital'], 3) ?? ''),
      __paciente: String(getByCandidates(object, ['Paciente', 'Nome Paciente'], 4) ?? ''),
      __medicamento: String(getByCandidates(object, ['Medicamento', 'Produto', 'Descrição', 'Descricao'], 10) ?? ''),
      __principioAtivo: String(getByCandidates(object, ['PrincipioAtivo', 'Princípio Ativo', 'Principio Ativo'], 16) ?? ''),
      __codBarra: onlyDigits(getByCandidates(object, ['CodBarra', 'Código de Barras', 'Codigo de Barras', 'EAN'], 22)),
      __qtde: toNumber(getByCandidates(object, ['Qtde', 'Quantidade', 'Qtd'], 11)),
      __lote: String(getByCandidates(object, ['Lote'], 19) ?? ''),
      __os: String(getByCandidates(object, ['OS', 'Ordem de serviço', 'Ordem de Servico', 'Ordem Serviço']) ?? ''),
    };

    hospitalRow.__osNormalizada = normalizeOS(hospitalRow.__os);
    hospitalRow.__key = buildKey(hospitalRow.__osNormalizada, hospitalRow.__codBarra);
    hospitalRow.__isExcluded = isExcludedItem(hospitalRow.__medicamento);

    if (hospitalRow.__isExcluded) {
      validations.push(validation('Itens excluídos da avaliação', 'INFO', 'HOSPITAL_ORIGINAL', rowNumber, hospitalRow, `Item excluído: ${hospitalRow.__medicamento}`));
    }
    rows.push(hospitalRow);
  });
  return rows;
}

function readControlRows(sheet, validations) {
  const { headers, startRow } = extractHeaders(sheet);
  const rows = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber <= startRow) return;
    const object = rowToObject(row, headers);
    if (isEmptyObject(object)) return;

    const controlRow = {
      ...object,
      __rowId: rowNumber,
      __data: getByCandidates(object, ['Data', 'Dt', 'Data Otimização', 'Data Otimizacao']),
      __os: String(getByCandidates(object, ['Ordem de serviço', 'Ordem de Servico', 'OS', 'Ordem Serviço']) ?? ''),
      __medicamento: String(getByCandidates(object, ['Medicamento', 'Produto']) ?? ''),
      __qtde: toNumber(getByCandidates(object, ['Quantidade (mg)', 'Quantidade', 'Qtde', 'Qtd'])),
      __lote: String(getByCandidates(object, ['Lote']) ?? ''),
      __motivo: String(getByCandidates(object, ['Motivo']) ?? ''),
      __unidadeDestino: String(getByCandidates(object, ['Unidade de Destino', 'Unidade Destino', 'Destino']) ?? ''),
      __codBarra: onlyDigits(getByCandidates(object, ['CodBarra', 'Código de Barras', 'Codigo de Barras', 'EAN'])),
      __status: 'Não avaliado como otimização',
      __observacao: 'Linha mantida no controle original; motivo/unidade/data serão avaliados no processamento.',
      __tipoMatch: '',
      __confianca: '',
      __hospitalAssociado: '',
      __dataAssociacao: '',
      __used: false,
    };

    controlRow.__osNormalizada = normalizeOS(controlRow.__os);
    controlRow.__medicamentoBase = normalizeMedicineBase(controlRow.__medicamento);
    controlRow.__key = buildKey(controlRow.__osNormalizada, controlRow.__codBarra);

    if (!controlRow.__osNormalizada) {
      validations.push(validation('Divergência de OS', 'BLOQUEIO', 'CONTROLE DE MEDICAMENTOS', rowNumber, controlRow, 'OS ausente ou inválida no controle.'));
    }
    rows.push(controlRow);
  });
  return rows;
}

function associateRows(hospitalRows, controlRows, validations, hospitalName) {
  const associations = [];
  const usableHospitalRows = hospitalRows.filter((row) => !row.__isExcluded);
  const hospitalByOS = groupBy(usableHospitalRows, (row) => row.__osNormalizada);
  const hospitalByKey = groupBy(usableHospitalRows, (row) => row.__key);
  const consumedKeys = new Set();

  if (!hospitalName) {
    validations.push({ Tipo: 'Filtro de otimização', Severidade: 'ALERTA', Aba: 'CONTROLE DE MEDICAMENTOS', Linha: 'todas', OS_Normalizada: '', CodBarra: '', Mensagem: 'Nome do hospital não informado; o filtro por Unidade de Destino não foi aplicado.' });
  }

  for (const control of controlRows) {
    const isOptimization = normalizeText(control.__motivo).includes('OTIMIZA');
    const destinationOk = !hospitalName || normalizeText(control.__unidadeDestino).includes(normalizeText(hospitalName));

    if (!isOptimization) {
      control.__status = 'Fora do filtro de otimização';
      control.__observacao = 'Motivo não contém OTIMIZA.';
      continue;
    }
    if (!destinationOk) {
      control.__status = 'Fora do filtro de otimização';
      control.__observacao = 'Unidade de Destino não contém o nome do hospital informado.';
      continue;
    }

    const candidatesByOS = hospitalByOS.get(control.__osNormalizada) ?? [];
    if (candidatesByOS.length === 0) {
      refuseControl(control, validations, 'Otimização sem correspondência segura', 'Nenhum item do hospital com a mesma OS normalizada.');
      continue;
    }

    const dateMatches = candidatesByOS.filter((hospital) => sameDayMonth(hospital.__data, control.__data));
    if (dateMatches.length === 0) {
      refuseControl(control, validations, 'Associação recusada', 'OS encontrada, mas a data não coincide em dia e mês.');
      continue;
    }

    const productCandidates = dateMatches.filter((hospital) => medicineCompatible(control.__medicamentoBase, hospital));
    if (productCandidates.length === 0) {
      refuseControl(control, validations, 'Tentativa de associação apenas por OS', 'Apenas a OS coincidiu; medicamento/princípio ativo incompatível.');
      continue;
    }

    const possibleBarcodes = unique(productCandidates.map((row) => row.__codBarra).filter(Boolean));
    const originalControlBarcode = control.__codBarra;
    if (!control.__codBarra) {
      if (possibleBarcodes.length === 1) {
        control.__codBarra = possibleBarcodes[0];
        control.__key = buildKey(control.__osNormalizada, control.__codBarra);
      } else {
        const message = possibleBarcodes.length === 0 ? 'Ausência de CodBarra compatível no hospital.' : 'Múltiplos CodBarra possíveis para a mesma OS/medicamento.';
        refuseControl(control, validations, possibleBarcodes.length === 0 ? 'Ausência de CodBarra compatível' : 'Múltiplos CodBarra possíveis na mesma OS', message);
        continue;
      }
    }

    const keyMatches = (hospitalByKey.get(control.__key) ?? []).filter((hospital) => medicineCompatible(control.__medicamentoBase, hospital) && sameDayMonth(hospital.__data, control.__data));
    if (keyMatches.length === 0) {
      refuseControl(control, validations, 'Divergência de CodBarra', 'CodBarra + OS não encontrou produto compatível no hospital.');
      continue;
    }
    if (consumedKeys.has(control.__key)) {
      refuseControl(control, validations, 'Possível duplicidade de uso da otimização', 'Possível duplicidade: já existe otimização consumida para esta chave OS+CodBarra.');
      continue;
    }

    const hospital = keyMatches[0];
    const type = originalControlBarcode ? 'MATCH FORTE' : 'MATCH MÉDIO';
    const notes = [];
    if (!sameYear(hospital.__data, control.__data) && sameDayMonth(hospital.__data, control.__data)) {
      notes.push('Ano do controle divergente; considerada a data operacional do hospital.');
      validations.push(validation('Correção de data por divergência de ano', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, `Controle ${formatDate(control.__data)} ajustado pela data do hospital ${formatDate(hospital.__data)}.`));
    }
    if (normalizeText(hospital.__lote) !== normalizeText(control.__lote) && control.__lote) {
      notes.push('Lote divergente não bloqueante.');
      validations.push(validation('Divergência de lote, como alerta não bloqueante', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, `Hospital: ${hospital.__lote || 'vazio'} | Otimização: ${control.__lote || 'vazio'}.`));
    }

    control.__status = 'Associado';
    control.__observacao = notes.join(' ') || 'Associação segura por OS normalizada + CodBarra.';
    control.__tipoMatch = type;
    control.__confianca = type === 'MATCH FORTE' ? 'Alta' : 'Média';
    control.__hospitalAssociado = hospital.__os;
    control.__dataAssociacao = formatDate(hospital.__data);
    control.__used = true;
    consumedKeys.add(control.__key);

    associations.push({
      osNormalizada: control.__osNormalizada,
      codBarra: control.__codBarra,
      key: control.__key,
      dataHospital: hospital.__data,
      dataControle: control.__data,
      medicamentoHospital: hospital.__medicamento,
      medicamentoControle: control.__medicamento,
      principioAtivo: hospital.__principioAtivo,
      qtdePrescrita: hospital.__qtde,
      qtdeOtimizada: control.__qtde,
      loteHospital: hospital.__lote,
      loteOtimizacao: control.__lote,
      tipoMatch: control.__tipoMatch,
      confianca: control.__confianca,
      observacao: control.__observacao,
      hospitalRowId: hospital.__rowId,
      controlRowId: control.__rowId,
    });
  }

  for (const control of controlRows) {
    if (normalizeText(control.__motivo).includes('OTIMIZA') && !control.__used && control.__status !== 'Fora do filtro de otimização') {
      validations.push(validation('Quantidade otimizada não utilizada', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, `Quantidade ${control.__qtde} não aplicada em baixa.`));
    }
  }
  return associations;
}

function buildOutputWorkbook(hospitalRows, controlRows, associations, validations, originalHospitalSheet) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Christian Web App';
  workbook.created = new Date();

  copySheetValues(originalHospitalSheet, workbook.addWorksheet(REQUIRED_SHEETS[0]));

  const controlColumns = Object.keys(stripInternal(controlRows[0] ?? {})).concat(['OS Normalizada', 'Medicamento Base', 'CodBarra Produto Hospital', 'Chave OS+CodBarra', 'OS Hospital Associada', 'Data Associação', 'Tipo de Match', 'Confiança da Associação', 'Status Associação', 'Observação']);
  addJsonSheet(workbook, REQUIRED_SHEETS[1], controlRows.map((row) => ({
    ...stripInternal(row),
    'OS Normalizada': row.__osNormalizada,
    'Medicamento Base': row.__medicamentoBase,
    'CodBarra Produto Hospital': row.__codBarra,
    'Chave OS+CodBarra': row.__key,
    'OS Hospital Associada': row.__hospitalAssociado,
    'Data Associação': row.__dataAssociacao,
    'Tipo de Match': row.__tipoMatch,
    'Confiança da Associação': row.__confianca,
    'Status Associação': row.__status,
    Observação: row.__observacao,
  })), controlColumns);

  addJsonSheet(workbook, REQUIRED_SHEETS[2], associations.map((item) => ({
    'OS Normalizada': item.osNormalizada,
    CodBarra: item.codBarra,
    'Chave OS+CodBarra': item.key,
    'Data Hospital': item.dataHospital,
    'Data Controle': item.dataControle,
    'Medicamento Hospital': item.medicamentoHospital,
    'Medicamento Controle': item.medicamentoControle,
    PrincipioAtivo: item.principioAtivo,
    'Qtde Prescrita': item.qtdePrescrita,
    'Qtde Otimizada': item.qtdeOtimizada,
    'Lote Hospital': item.loteHospital,
    'Lote Otimização': item.loteOtimizacao,
    'Tipo de Match': item.tipoMatch,
    'Confiança da Associação': item.confianca,
    Observação: item.observacao,
  })));

  const associationByHospitalRow = new Map(associations.map((item) => [item.hospitalRowId, item]));
  const baixarRows = hospitalRows.filter((row) => !row.__isExcluded).map((row) => downloadRecord(row, associationByHospitalRow.get(row.__rowId)));
  addJsonSheet(workbook, REQUIRED_SHEETS[3], baixarRows);
  addJsonSheet(workbook, REQUIRED_SHEETS[4], baixarRows.map((row) => ({
    Data: row.Data,
    Cliente: row.Cliente,
    Paciente: row.Paciente,
    'Medicamento Hospital': row['Medicamento Hospital'],
    'Qtde Prescrita': row['Qtde Prescrita'],
    'Qtde Otimizada': row['Qtde Otimizada'],
    'Lote Otimização': row['Lote Otimização'],
    'Qtde Baixa': row['Qtde Baixa'],
    'Lote da Baixa': row['Lote da Baixa'],
    'Status Otimização': row['Status Otimização'],
  })));
  addJsonSheet(workbook, REQUIRED_SHEETS[5], validations);

  workbook.worksheets.forEach(formatWorksheet);
  return workbook;
}

function downloadRecord(row, association) {
  const optimizedQuantity = association?.qtdeOtimizada ?? 0;
  return {
    Data: row.__data,
    Cliente: row.__cliente,
    Paciente: row.__paciente,
    'Medicamento Hospital': row.__medicamento,
    PrincipioAtivo: row.__principioAtivo,
    OS: row.__os,
    'OS Normalizada': row.__osNormalizada,
    CodBarra: row.__codBarra,
    'Chave OS+CodBarra': row.__key,
    'Qtde Prescrita': row.__qtde,
    'Qtde Otimizada': optimizedQuantity,
    'Lote Otimização': association?.loteOtimizacao ?? 'Sem otimização',
    'Qtde Baixa': row.__qtde - optimizedQuantity,
    'Lote da Baixa': row.__lote,
    'Status Otimização': association ? 'Com otimização' : 'Sem otimização',
  };
}

function refuseControl(control, validations, type, message) {
  control.__status = 'Associação recusada';
  control.__observacao = message;
  validations.push(validation(type, 'BLOQUEIO', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, message));
}

function extractHeaders(sheet) {
  let startRow = 1;
  let bestScore = -1;
  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber > 10) return;
    const values = rowValues(row).map((value) => normalizeText(value));
    const score = ['DATA', 'MEDICAMENTO', 'PRODUTO', 'QTDE', 'QUANTIDADE', 'LOTE', 'OS', 'ORDEM'].filter((term) => values.some((value) => value.includes(term))).length;
    if (score > bestScore) {
      bestScore = score;
      startRow = rowNumber;
    }
  });
  const headers = rowValues(sheet.getRow(startRow)).map((value, index) => String(value || `Coluna ${index + 1}`).trim());
  return { headers, startRow };
}

function rowValues(row) {
  const values = [];
  row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    values[colNumber - 1] = normalizeCell(cell.value);
  });
  return values;
}

function rowToObject(row, headers) {
  const object = {};
  headers.forEach((header, index) => {
    object[header] = normalizeCell(row.getCell(index + 1).value);
  });
  return object;
}

function normalizeCell(value) {
  if (value == null) return null;
  if (value instanceof Date) return value;
  if (typeof value === 'object') {
    if ('result' in value) return normalizeCell(value.result);
    if ('text' in value) return String(value.text);
    if ('richText' in value) return value.richText.map((part) => part.text).join('');
  }
  return value;
}

function getByCandidates(object, candidates, fallbackIndex) {
  const entries = Object.entries(object);
  for (const candidate of candidates) {
    const direct = entries.find(([key]) => normalizeText(key) === normalizeText(candidate));
    if (direct) return direct[1];
    const partial = entries.find(([key]) => normalizeText(key).includes(normalizeText(candidate)) || normalizeText(candidate).includes(normalizeText(key)));
    if (partial) return partial[1];
  }
  if (fallbackIndex) return entries[fallbackIndex - 1]?.[1] ?? null;
  return null;
}

function addJsonSheet(workbook, name, rows, forcedColumns) {
  const sheet = workbook.addWorksheet(name);
  const columns = forcedColumns?.length ? forcedColumns : Object.keys(rows[0] ?? { Mensagem: '' });
  sheet.addRow(columns);
  rows.forEach((row) => sheet.addRow(columns.map((column) => row[column] ?? '')));
}

function copySheetValues(source, target) {
  source.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const targetRow = target.getRow(rowNumber);
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      targetRow.getCell(colNumber).value = normalizeCell(cell.value);
    });
    targetRow.commit();
  });
}

function formatWorksheet(sheet) {
  if (sheet.rowCount === 0) return;
  const columnCount = Math.max(sheet.columnCount, 1);
  const header = sheet.getRow(1);
  header.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1D4ED8' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  sheet.views = [{ state: 'frozen', ySplit: 1 }];
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columnCount } };

  if (sheet.rowCount > 1 && sheet.name !== 'HOSPITAL_ORIGINAL') {
    sheet.addTable({
      name: sanitizeTableName(sheet.name),
      ref: 'A1',
      headerRow: true,
      totalsRow: false,
      style: { theme: 'TableStyleMedium2', showRowStripes: true },
      columns: Array.from({ length: columnCount }, (_, index) => ({ name: String(sheet.getCell(1, index + 1).value || `Coluna ${index + 1}`) })),
      rows: Array.from({ length: sheet.rowCount - 1 }, (_, rowIndex) => Array.from({ length: columnCount }, (_, columnIndex) => sheet.getCell(rowIndex + 2, columnIndex + 1).value)),
    });
  }

  sheet.columns.forEach((column) => {
    let width = 12;
    column.eachCell?.({ includeEmpty: true }, (cell) => {
      width = Math.max(width, String(cell.value ?? '').slice(0, 60).length + 2);
      if (cell.value instanceof Date) cell.numFmt = 'dd/mm/yyyy';
      if (typeof cell.value === 'number') cell.numFmt = '#,##0.00';
    });
    column.width = Math.min(width, 60);
  });
}

function validation(type, severity, sheet, line, row, message) {
  return { Tipo: type, Severidade: severity, Aba: sheet, Linha: line, OS_Normalizada: row.__osNormalizada, CodBarra: row.__codBarra, Mensagem: message };
}

function isEmptyObject(object) {
  return Object.values(object).every((value) => value == null || value === '');
}

function stripInternal(row) {
  return Object.fromEntries(Object.entries(row).filter(([key]) => !key.startsWith('__')));
}

function normalizeOS(value) {
  return onlyDigits(value).slice(0, 7);
}

function onlyDigits(value) {
  return String(value ?? '').replace(/\D/g, '');
}

function normalizeMedicineBase(value) {
  return String(value ?? '').split('-')[0].trim();
}

function normalizeText(value) {
  return String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
}

function medicineCompatible(controlMedicineBase, hospital) {
  const control = normalizeText(controlMedicineBase);
  if (!control) return false;
  const medicine = normalizeText(hospital.__medicamento);
  const active = normalizeText(hospital.__principioAtivo);
  return medicine === control || active === control || medicine.includes(control) || control.includes(medicine) || active.includes(control) || (active && control.includes(active));
}

function isExcludedItem(medicine) {
  const normalizedMedicine = normalizeText(medicine);
  return EXCLUDED_ITEMS.some((item) => normalizedMedicine.includes(normalizeText(item)) || normalizeText(item).includes(normalizedMedicine));
}

function buildKey(os, barcode) {
  return os && barcode ? `${os}|${barcode}` : '';
}

function toNumber(value) {
  if (typeof value === 'number') return value;
  const normalized = String(value ?? '').replace(/\./g, '').replace(',', '.').replace(/[^0-9.-]/g, '');
  const number = Number(normalized);
  return Number.isFinite(number) ? number : 0;
}

function parseDate(value) {
  if (value instanceof Date) return value;
  if (typeof value === 'number') return new Date(Math.round((value - 25569) * 86400 * 1000));
  const text = String(value ?? '').trim();
  const br = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (br) return new Date(Number(br[3].padStart(4, '20')), Number(br[2]) - 1, Number(br[1]));
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function sameDayMonth(a, b) {
  const first = parseDate(a);
  const second = parseDate(b);
  if (!first || !second) return true;
  return first.getDate() === second.getDate() && first.getMonth() === second.getMonth();
}

function sameYear(a, b) {
  const first = parseDate(a);
  const second = parseDate(b);
  if (!first || !second) return true;
  return first.getFullYear() === second.getFullYear();
}

function formatDate(value) {
  const date = parseDate(value);
  if (!date) return String(value ?? '');
  return new Intl.DateTimeFormat('pt-BR').format(date);
}

function unique(values) {
  return [...new Set(values)];
}

function groupBy(items, getKey) {
  const map = new Map();
  for (const item of items) {
    const key = getKey(item);
    if (!key) continue;
    map.set(key, [...(map.get(key) ?? []), item]);
  }
  return map;
}

function sanitizeTableName(name) {
  return `Tabela_${name.replace(/[^A-Za-z0-9]/g, '_')}`.slice(0, 31);
}

function setStatus(message, type) {
  statusElement.textContent = message;
  statusElement.className = `status ${type}`;
}
