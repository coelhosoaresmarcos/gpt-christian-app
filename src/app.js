const REQUIRED_SHEETS = [
  'HOSPITAL_ORIGINAL',
  'CONTROLE DE MEDICAMENTOS',
  'ASSOCIACOES',
  'BAIXAR',
  'RELATORIO',
  'VALIDACAO',
  'DIAGNOSTICO_OTIMIZACOES',
];

const MEDICINE_EQUIVALENCE_GROUPS = [
  ['GENLIBBS', 'GENCITABINA', 'GEMCITABINA', 'GEMCITABINE'],
];


const DIAGNOSTIC_COLUMNS = [
  'Linha do controle',
  'Motivo',
  'OS controle original',
  'OS controle normalizada',
  'Data controle',
  'Unidade origem',
  'Unidade destino',
  'Unidade avaliada',
  'Unidade de Destino da otimização',
  'Medicamento controle',
  'Medicamento base controle',
  'Quantidade controle',
  'Lote controle',
  'Validade controle',
  'Hospital informado na tela',
  'Destino compatível',
  'Elegível para otimização',
  'Encontrou OS no hospital',
  'Quantidade de linhas do hospital com a mesma OS',
  'Datas encontradas no hospital para essa OS',
  'Data compatível por dia/mês',
  'Medicamentos hospital coluna I encontrados na mesma OS',
  'Medicamentos alternativos coluna O encontrados na mesma OS',
  'Princípios ativos coluna P encontrados na mesma OS',
  'Medicamento compatível',
  'CodBarra candidato encontrado',
  'Quantidade de CodBarra candidatos',
  'Status final da linha',
  'Motivo final exato',
  'Motivo final',
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

const hasDocument = typeof document !== 'undefined';
const hospitalInput = hasDocument ? document.querySelector('#hospitalFile') : null;
const controlInput = hasDocument ? document.querySelector('#controlFile') : null;
const hospitalFileName = hasDocument ? document.querySelector('#hospitalFileName') : null;
const controlFileName = hasDocument ? document.querySelector('#controlFileName') : null;
const hospitalNameInput = hasDocument ? document.querySelector('#hospitalName') : null;
const processButton = hasDocument ? document.querySelector('#processButton') : null;
const downloadButton = hasDocument ? document.querySelector('#downloadButton') : null;
const statusElement = hasDocument ? document.querySelector('#status') : null;
const resultCard = hasDocument ? document.querySelector('#resultCard') : null;
const summaryElement = hasDocument ? document.querySelector('#summary') : null;

let generatedBlob = null;
let generatedFileName = '';

if (hasDocument) {
  hospitalInput.addEventListener('change', () => updateFileName(hospitalInput, hospitalFileName));
  controlInput.addEventListener('change', () => updateFileName(controlInput, controlFileName));
  processButton.addEventListener('click', processFiles);
  downloadButton.addEventListener('click', downloadGeneratedFile);
}

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
  const optimizationDiagnostics = [];
  const hospitalRows = readHospitalRows(hospitalSheet, validations);
  const controlRows = readControlRows(controlSheet, validations);
  const evaluatedHospitalUnit = resolveEvaluatedHospitalUnit(hospitalRows, hospitalName, validations);
  const associations = associateRows(hospitalRows, controlRows, validations, evaluatedHospitalUnit, optimizationDiagnostics);
  const outputWorkbook = buildOutputWorkbook(hospitalRows, controlRows, associations, validations, hospitalSheet, optimizationDiagnostics);
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

    const physical = (index) => getPhysicalCell(row, index);
    const columnI = physical(9);
    const columnO = physical(15);
    const rawMedicine = getByExactHeader(object, ['Medicamento Hospital', 'Medicamento', 'Produto', 'Descrição', 'Descricao']) ?? columnI;
    const medication = shouldUseHospitalMedicineFallback(rawMedicine) ? columnI : rawMedicine;
    const alternativeMedication = columnO || getByExactHeader(object, ['Medicamento Alternativo', 'Nome Comercial', 'Comercial']);

    const hospitalRow = {
      ...object,
      __rowId: rowNumber,
      __data: getByCandidates(object, ['Data', 'Dt Atendimento', 'Data Atendimento']) ?? physical(3),
      __cliente: String(getByCandidates(object, ['Cliente', 'Hospital']) ?? physical(1) ?? ''),
      __paciente: String(getByCandidates(object, ['Paciente', 'Nome Paciente']) ?? physical(4) ?? ''),
      __os: String(getByCandidates(object, ['OS', 'Ordem de serviço', 'Ordem de Servico', 'Ordem Serviço']) ?? physical(8) ?? ''),
      __medicamento: String(medication ?? ''),
      __produtoHospital: String(getByExactHeader(object, ['Produto Hospital', 'Produto']) ?? ''),
      __medicamentoColunaI: String(columnI ?? ''),
      __medicamentoAlternativo: String(alternativeMedication ?? ''),
      __principioAtivo: String(getByExactHeader(object, ['PrincipioAtivo', 'Princípio Ativo', 'Principio Ativo']) ?? physical(16) ?? ''),
      __qtde: toNumber(getByCandidates(object, ['Qtde', 'Quantidade', 'Qtd']) ?? physical(11)),
      __lote: String(getByCandidates(object, ['Lote']) ?? physical(19) ?? ''),
      __codBarra: normalizeBarcode(getByCandidates(object, ['CodBarra', 'Código de Barras', 'Codigo de Barras', 'EAN']) ?? physical(22)),
    };

    hospitalRow.__osNormalizada = normalizeOS(hospitalRow.__os);
    hospitalRow.__medicamentoNormalizado = normalizeMedicineProduct(hospitalRow.__medicamento);
    hospitalRow.__produtoHospitalNormalizado = normalizeMedicineProduct(hospitalRow.__produtoHospital);
    hospitalRow.__medicamentoAlternativoNormalizado = normalizeMedicineProduct(hospitalRow.__medicamentoAlternativo);
    hospitalRow.__principioAtivoNormalizado = normalizeMedicineProduct(hospitalRow.__principioAtivo);
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
      __data: getByCandidates(object, ['Data', 'Dt', 'Data Otimização', 'Data Otimizacao'], 3),
      __os: String(getByCandidates(object, ['Ordem de serviço', 'Ordem de Servico', 'OS', 'Ordem Serviço'], 2) ?? ''),
      __unidadeOrigem: String(getByExactHeaderOrPosition(object, ['Unidade de Origem', 'Unidade Origem', 'Origem'], 4) ?? ''),
      __unidadeDestino: String(getByExactHeaderOrPosition(object, ['Unidade de Destino', 'Unidade Destino', 'Destino'], 5) ?? ''),
      __paciente: String(getByCandidates(object, ['Paciente', 'Nome Paciente'], 6) ?? ''),
      __medicamento: String(getByCandidates(object, ['Medicamento', 'Produto'], 7) ?? ''),
      __qtde: toNumber(getByCandidates(object, ['Quantidade (mg)', 'Quantidade', 'Qtde', 'Qtd'], 8)),
      __lote: String(getByExactHeaderOrPosition(object, ['Lote'], 9) ?? ''),
      __validade: getByExactHeaderOrPosition(object, ['Validade', 'Data de Validade', 'Val.'], 10),
      __laboratorio: String(getByCandidates(object, ['Laboratório', 'Laboratorio', 'Lab'], 11) ?? ''),
      __motivo: String(getByCandidates(object, ['Tipo/Motivo', 'Tipo', 'Motivo'], 1) ?? ''),
      __codBarra: normalizeBarcode(getByCandidates(object, ['CodBarra', 'Código de Barras', 'Codigo de Barras', 'EAN'])),
      __codBarraProdutoHospital: '',
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
    controlRow.__medicamentoNormalizado = normalizeMedicineProduct(controlRow.__medicamento);
    controlRow.__key = buildKey(controlRow.__osNormalizada, controlRow.__codBarraProdutoHospital || controlRow.__codBarra);

    if (!controlRow.__osNormalizada) {
      validations.push(validation('Divergência de OS', 'BLOQUEIO', 'CONTROLE DE MEDICAMENTOS', rowNumber, controlRow, 'OS ausente ou inválida no controle.'));
    }
    rows.push(controlRow);
  });
  return rows;
}

function associateRows(hospitalRows, controlRows, validations, hospitalName, optimizationDiagnostics = []) {
  const associations = [];
  const usableHospitalRows = hospitalRows.filter((row) => !row.__isExcluded);
  const controlsByKey = new Map();
  const keyUsage = new Map();
  const evaluatedHospitalUnit = resolveEvaluatedHospitalUnit(usableHospitalRows, hospitalName, validations);
  const canApplyOptimizations = Boolean(evaluatedHospitalUnit);

  for (const hospital of usableHospitalRows) {
    hospital.__osNormalizada = normalizeOS(hospital.__osNormalizada || hospital.__os);
    hospital.__codBarra = normalizeBarcode(hospital.__codBarra);
    hospital.__key = buildKey(hospital.__osNormalizada, hospital.__codBarra);
    hospital.__medicamentoNormalizado = normalizeMedicineProduct(hospital.__medicamento);
    hospital.__produtoHospitalNormalizado = normalizeMedicineProduct(hospital.__produtoHospital);
    hospital.__medicamentoColunaINormalizado = normalizeMedicineProduct(hospital.__medicamentoColunaI);
    hospital.__medicamentoAlternativoNormalizado = normalizeMedicineProduct(hospital.__medicamentoAlternativo);
    hospital.__principioAtivoNormalizado = normalizeMedicineProduct(hospital.__principioAtivo);

    if (!hospital.__osNormalizada) {
      validations.push(validation('Divergência de OS', 'BLOQUEIO', 'HOSPITAL_ORIGINAL', hospital.__rowId, hospital, 'OS ausente ou inválida no hospital.'));
    }
    if (!hospital.__codBarra) {
      validations.push(validation('Ausência de CodBarra compatível', 'BLOQUEIO', 'HOSPITAL_ORIGINAL', hospital.__rowId, hospital, 'CodBarra ausente no hospital; a chave OS Normalizada + CodBarra não pode ser criada.'));
    }
  }

  const hospitalKeys = new Set(usableHospitalRows.map((row) => row.__key).filter(Boolean));

  for (const control of controlRows) {
    prepareControlForAssociation(control);
    const isOptimization = normalizeText(control.__motivo).includes('OTIMIZA');
    control.__unidadeAvaliada = evaluatedHospitalUnit;
    control.__destinationOk = false;
    control.__eligibleForOptimization = false;

    if (!isOptimization) {
      control.__status = 'Fora do filtro de otimização';
      control.__observacao = 'Motivo não contém OTIMIZA; linha preservada sem associação.';
      control.__remaining = 0;
      continue;
    }

    if (!canApplyOptimizations) {
      markOptimizationUnitUnknown(control, validations);
      continue;
    }

    const destinationOk = optimizationDestinationCompatible(control.__unidadeDestino, evaluatedHospitalUnit);
    control.__destinationOk = destinationOk;
    control.__eligibleForOptimization = destinationOk;
    if (!destinationOk) {
      markOptimizationDestinationIneligible(control, validations);
      continue;
    }

    const sameOS = usableHospitalRows.filter((hospital) => hospital.__osNormalizada && hospital.__osNormalizada === control.__osNormalizada);
    const medicineCountInOS = unique(sameOS.map((hospital) => normalizeMedicineProduct(hospital.__medicamento || hospital.__principioAtivo)).filter(Boolean)).length;
    if (medicineCountInOS > 1) {
      validations.push(validation('Múltiplos medicamentos na mesma OS', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, 'Há múltiplos medicamentos na mesma OS; CodBarra será obrigatório para a associação final.'));
    }

    if (!control.__osNormalizada) {
      markOptimizationNotUsed(control, validations, 'OS ausente ou inválida no controle; não é permitido associar sem OS Normalizada.');
      continue;
    }
    if (!Number.isFinite(control.__available) || control.__available <= 0) {
      markOptimizationNotUsed(control, validations, validationDetails(control, 'quantidade inválida'));
      continue;
    }
    const candidateResult = findHospitalCandidatesForControl(control, usableHospitalRows);
    const compatible = candidateResult.candidates;

    if (!compatible.length) {
      const candidateContext = sameOS.length ? sameOS : usableHospitalRows;
      setValidationCandidateContext(control, candidateContext, candidateResult.reason === 'OS encontrada, mas Medicamento Base do controle não corresponde a Medicamento Hospital, Medicamento Alternativo ou PrincipioAtivo.' ? 'OS encontrada, mas medicamento diferente' : candidateResult.reason);
      markOptimizationNotUsed(control, validations, candidateResult.reason);
      if (!sameOS.length) {
        validations.push(validation('Divergência de OS', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, candidateResult.reason));
      } else {
        validations.push(validation('Otimização sem correspondência segura', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, candidateResult.reason));
      }
      continue;
    }

    if (!String(control.__data ?? '').trim() || !parseDate(control.__data)) {
      validations.push(validation('Data do controle vazia ou inválida', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, 'Data do controle vazia ou inválida; alerta não bloqueante. A associação continuará por OS Normalizada + medicamento + CodBarra.'));
    }

    const possibleKeys = unique(compatible.map((hospital) => hospital.__key).filter(Boolean));
    const possibleBarcodes = unique(compatible.map((hospital) => hospital.__codBarra).filter(Boolean));
    if (!possibleKeys.length) {
      setValidationCandidateContext(control, compatible, 'medicamento compatível sem CodBarra candidato');
      refuseControl(control, validations, 'Ausência de CodBarra compatível', 'Medicamento compatível encontrado, mas sem CodBarra preenchido no hospital para criar a chave final.');
      continue;
    }
    if (possibleKeys.length > 1) {
      setValidationCandidateContext(control, compatible, 'múltiplos CodBarra candidatos');
      const barcodeDetails = describeBarcodeCandidates(compatible);
      refuseControl(control, validations, 'Múltiplos CodBarra possíveis', 'Múltiplos CodBarra possíveis para mesma OS e Medicamento Base.');
      validations.push(validation('Múltiplos CodBarra possíveis', 'BLOQUEIO', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, `CodBarra candidatos: ${possibleBarcodes.join('; ')}. Detalhes: ${barcodeDetails}.`));
      continue;
    }

    const targetKey = possibleKeys[0];
    const targetHospital = compatible.find((hospital) => hospital.__key === targetKey);
    const identifiedBarcode = targetHospital.__codBarra;
    if (control.__codBarra && control.__codBarra !== identifiedBarcode) {
      setValidationCandidateContext(control, compatible, 'CodBarra do controle diverge do CodBarra candidato');
      refuseControl(control, validations, 'Divergência de CodBarra', 'CodBarra do controle diverge do CodBarra compatível identificado no hospital.');
      continue;
    }

    const dateWithDifferentYear = compatible.find((hospital) => sameDayAndMonth(control.__data, hospital.__data) && !sameYear(control.__data, hospital.__data));
    if (dateWithDifferentYear) {
      validations.push(validation('Correção de data por divergência de ano', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, `Dia/mês coincidem; ano do controle será tratado conforme ano do hospital (${formatDate(dateWithDifferentYear.__data)}).`));
    }

    control.__codBarraProdutoHospital = identifiedBarcode;
    control.__key = targetKey;
    control.__tipoMatch = candidateResult.matchType;
    control.__confianca = candidateResult.confidence;
    control.__status = 'Disponível para otimização';
    control.__observacao = candidateResult.osDivergence
      ? 'OS divergente auditada; associação final autorizada pela chave OS Hospital + CodBarra.'
      : 'OS normalizada, data e medicamento compatíveis; associação final autorizada pela chave OS Normalizada + CodBarra.';
    control.__hospitalAssociado = appendUniqueText(control.__hospitalAssociado, targetHospital.__os);
    control.__dataAssociacao = appendUniqueText(control.__dataAssociacao, formatDate(targetHospital.__data));

    if (candidateResult.osDivergence) {
      validations.push(validation('Divergência de OS', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, `OS divergente entre controle e hospital. Controle: ${control.__osNormalizada}. Hospital: ${targetHospital.__osNormalizada}. Associação realizada por paciente + data + medicamento + CodBarra único.`));
    }

    if (!hospitalKeys.has(control.__key)) {
      validations.push(validation('Ausência de CodBarra compatível', 'BLOQUEIO', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, 'A chave OS Normalizada + CodBarra identificada não existe entre os itens baixáveis do hospital.'));
      control.__status = 'Associação recusada';
      continue;
    }
    controlsByKey.set(control.__key, [...(controlsByKey.get(control.__key) ?? []), control]);
  }

  for (const hospital of usableHospitalRows) {
    let remainingPrescription = Math.max(0, hospital.__qtde);
    const controls = (controlsByKey.get(hospital.__key) ?? []).filter((control) => control.__remaining > 0 && medicineCompatible(control.__medicamentoBase || control.__medicamentoNormalizado, hospital) && sameDayAndMonth(control.__data, hospital.__data));
    let optimizedQuantity = 0;
    const availableBefore = controls.reduce((sum, control) => sum + control.__remaining, 0);

    for (const control of controls) {
      if (remainingPrescription <= 0) break;
      const quantityToUse = Math.min(remainingPrescription, control.__remaining);
      if (quantityToUse <= 0) continue;

      const usedKey = keyUsage.get(control.__key);
      if (usedKey && usedKey !== control.__key) {
        validations.push(validation('Possível duplicidade de uso da otimização', 'BLOQUEIO', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, 'Tentativa de aplicar a mesma otimização em mais de um CodBarra diferente.'));
        continue;
      }
      keyUsage.set(control.__key, control.__key);

      control.__remaining -= quantityToUse;
      control.__usedQuantity += quantityToUse;
      optimizedQuantity += quantityToUse;
      remainingPrescription -= quantityToUse;
      control.__used = true;
      control.__status = control.__remaining > 0 ? 'Parcialmente consumido' : 'Consumido';
      control.__observacao = appendUniqueText(control.__observacao, `Quantidade ${quantityToUse} consumida pela chave final ${control.__key}.`);
      control.__hospitalAssociado = appendUniqueText(control.__hospitalAssociado, hospital.__os);
      control.__dataAssociacao = appendUniqueText(control.__dataAssociacao, formatDate(hospital.__data));

      if (normalizeText(control.__lote) && normalizeText(control.__lote) !== normalizeText(hospital.__lote)) {
        validations.push(validation('Divergência de lote como alerta não bloqueante', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, 'Lote da otimização diverge do lote do hospital; associação mantida por OS Normalizada + CodBarra.'));
      }

      associations.push(associationRecord(hospital, control, quantityToUse, 'Com otimização', `Associação segura pela chave final ${hospital.__key}. Medicamento usado apenas para identificar o CodBarra do hospital.`, availableBefore));
    }

    if (optimizedQuantity > 0 && optimizedQuantity < hospital.__qtde) {
      validations.push(validation('Quantidade otimizada não utilizada', 'ALERTA', 'HOSPITAL_ORIGINAL', hospital.__rowId, hospital, `Qtde prescrita maior que a otimizada; baixa calculada com saldo ${hospital.__qtde - optimizedQuantity}.`));
    }
  }

  for (const control of controlRows) {
    if (control.__status === 'Disponível para otimização') {
      const unused = unusedOptimizationReason(control, usableHospitalRows);
      control.__status = 'Não utilizado';
      control.__observacao = unused.message;
      validations.push(validation(unused.type, unused.severity, 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, unused.message));
    }
    if (control.__usedQuantity > control.__available) {
      validations.push(validation('Possível duplicidade de uso da otimização', 'BLOQUEIO', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, `Uso (${control.__usedQuantity}) maior que saldo disponível (${control.__available}).`));
    }
    if (normalizeText(control.__motivo).includes('OTIMIZA') && control.__remaining > 0 && control.__usedQuantity > 0) {
      validations.push(validation('Quantidade otimizada não utilizada', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, `Sobra de ${control.__remaining} após otimização pela chave ${control.__key}.`));
    }
  }
  optimizationDiagnostics.splice(0, optimizationDiagnostics.length, ...buildOptimizationDiagnostics(controlRows, usableHospitalRows, evaluatedHospitalUnit));
  return associations;
}


function buildOptimizationDiagnostics(controlRows, hospitalRows, evaluatedHospitalUnit) {
  return controlRows
    .filter((control) => normalizeText(control.__motivo).includes('OTIMIZA'))
    .map((control) => optimizationDiagnosticRecord(control, hospitalRows, evaluatedHospitalUnit));
}

function optimizationDiagnosticRecord(control, hospitalRows, evaluatedHospitalUnit) {
  const sameOS = hospitalRows.filter((hospital) => hospital.__osNormalizada && hospital.__osNormalizada === control.__osNormalizada);
  const controlDateParsed = parseDate(control.__data);
  const hasControlDateText = String(control.__data ?? '').trim() !== '';
  const dateCompatibleRows = sameOS.filter((hospital) => sameDayAndMonth(control.__data, hospital.__data));
  const medicineRowsScope = dateCompatibleRows.length ? dateCompatibleRows : sameOS;
  const sameOSCompatible = medicineRowsScope.filter((hospital) => medicineCompatible(control.__medicamentoBase || control.__medicamentoNormalizado, hospital));
  const candidateResult = findHospitalCandidatesForControl(control, hospitalRows);
  const compatible = candidateResult.candidates.length ? candidateResult.candidates : sameOSCompatible;
  const barcodeCandidates = unique(compatible.map((hospital) => hospital.__codBarra).filter(Boolean));
  const destinationOk = typeof control.__destinationOk === 'boolean' ? control.__destinationOk : optimizationDestinationCompatible(control.__unidadeDestino, evaluatedHospitalUnit);
  const dateCompatible = !sameOS.length || !hasControlDateText || !controlDateParsed || dateCompatibleRows.length > 0 || candidateResult.candidates.length > 0;
  const medicineCompatibleInOS = sameOS.some((hospital) => medicineCompatible(control.__medicamentoBase || control.__medicamentoNormalizado, hospital));
  const medicineCompatibleFound = medicineCompatibleInOS || candidateResult.candidates.length > 0;
  const status = diagnosticStatus(control);

  return {
    'Linha do controle': control.__rowId,
    Motivo: control.__motivo,
    'OS controle original': control.__os,
    'OS controle normalizada': control.__osNormalizada,
    'Data controle': control.__data,
    'Unidade origem': control.__unidadeOrigem,
    'Unidade destino': control.__unidadeDestino,
    'Unidade avaliada': evaluatedHospitalUnit || '',
    'Unidade de Destino da otimização': control.__unidadeDestino,
    'Medicamento controle': control.__medicamento,
    'Medicamento base controle': control.__medicamentoBase,
    'Quantidade controle': control.__available ?? control.__qtde ?? '',
    'Lote controle': control.__lote,
    'Validade controle': control.__validade,
    'Hospital informado na tela': evaluatedHospitalUnit || '(não identificado)',
    'Destino compatível': destinationOk ? 'SIM' : 'NÃO',
    'Elegível para otimização': control.__eligibleForOptimization ? 'SIM' : 'NÃO',
    'Encontrou OS no hospital': sameOS.length ? 'SIM' : 'NÃO',
    'Quantidade de linhas do hospital com a mesma OS': sameOS.length,
    'Datas encontradas no hospital para essa OS': unique(sameOS.map((hospital) => formatDate(hospital.__data)).filter(Boolean)).join('; '),
    'Data compatível por dia/mês': dateCompatible ? 'SIM' : 'NÃO',
    'Medicamentos hospital coluna I encontrados na mesma OS': unique(sameOS.map((hospital) => hospital.__medicamentoColunaI || hospital.__medicamento).filter(Boolean)).join('; '),
    'Medicamentos alternativos coluna O encontrados na mesma OS': unique(sameOS.map((hospital) => hospital.__medicamentoAlternativo).filter(Boolean)).join('; '),
    'Princípios ativos coluna P encontrados na mesma OS': unique(sameOS.map((hospital) => hospital.__principioAtivo).filter(Boolean)).join('; '),
    'Medicamento compatível': medicineCompatibleFound ? 'SIM' : 'NÃO',
    'CodBarra candidato encontrado': describeBarcodeCandidates(compatible),
    'Quantidade de CodBarra candidatos': barcodeCandidates.length,
    'Status final da linha': status,
    'Motivo final exato': diagnosticFinalReason(control, { sameOS, dateCompatible, medicineCompatibleInOS, barcodeCandidates, destinationOk, hasControlDateText, controlDateParsed, evaluatedHospitalUnit }),
    'Motivo final': diagnosticFinalReason(control, { sameOS, dateCompatible, medicineCompatibleInOS, barcodeCandidates, destinationOk, hasControlDateText, controlDateParsed, evaluatedHospitalUnit }),
  };
}

function diagnosticStatus(control) {
  if (control.__status === 'Não elegível - destino diferente') return 'NÃO ELEGÍVEL - DESTINO DIFERENTE';
  if (control.__status === 'Não elegível - unidade avaliada não identificada') return 'NÃO ELEGÍVEL - UNIDADE AVALIADA NÃO IDENTIFICADA';
  if (control.__status === 'Consumido') return 'CONSUMIDA';
  if (control.__status === 'Parcialmente consumido') return 'PARCIALMENTE CONSUMIDA';
  if (control.__status === 'Associação recusada') return 'RECUSADA';
  return 'NÃO USADA';
}

function diagnosticFinalReason(control, context) {
  if (control.__status === 'Consumido') return `Consumida pela chave final ${control.__key}.`;
  if (control.__status === 'Parcialmente consumido') return `Parcialmente consumida pela chave final ${control.__key}; saldo restante ${control.__remaining}.`;
  if (!context.evaluatedHospitalUnit) return 'Unidade avaliada não identificada; otimizações não aplicadas por segurança.';
  if (!context.destinationOk) return 'Unidade de Destino da otimização não corresponde à unidade avaliada.';
  if (!Number.isFinite(control.__available) || control.__available <= 0) return validationDetails(control, 'quantidade inválida');
  if (!control.__osNormalizada) return 'OS ausente ou inválida no controle; não é permitido associar sem OS Normalizada.';
  if (!context.sameOS.length) return validationDetails(control, 'OS não encontrada no hospital');
  if (!context.dateCompatible && context.medicineCompatibleInOS) return 'OS encontrada, medicamento compatível, mas data do controle não coincide com dia/mês da data do hospital.';
  if (!context.dateCompatible) return 'OS encontrada, mas data do controle não coincide com dia/mês da data do hospital.';
  if (!context.medicineCompatibleInOS) return 'OS encontrada, mas Medicamento Base do controle não corresponde a Medicamento Hospital, Medicamento Alternativo ou PrincipioAtivo.';
  if (!context.barcodeCandidates.length) return 'Medicamento compatível encontrado na OS, mas sem CodBarra preenchido no hospital para criar a chave final.';
  if (context.barcodeCandidates.length > 1) return `Múltiplos CodBarra possíveis para mesma OS e Medicamento Base: ${describeBarcodeCandidates(context.sameOS.filter((hospital) => medicineCompatible(control.__medicamentoBase || control.__medicamentoNormalizado, hospital)))}.`;
  if (control.__codBarra && control.__codBarra !== context.barcodeCandidates[0]) return 'CodBarra do controle diverge do CodBarra compatível identificado no hospital.';
  if (!context.hasControlDateText || !context.controlDateParsed) return `Data do controle vazia ou inválida gerou alerta não bloqueante; ${control.__observacao || 'associação tentou seguir por OS + medicamento + CodBarra.'}`;
  return control.__observacao || 'Otimização não consumida após a avaliação por OS Normalizada + CodBarra.';
}

function describeBarcodeCandidates(rows) {
  const descriptions = [];
  for (const row of rows) {
    if (!row.__codBarra) continue;
    const description = `${row.__codBarra} (I: ${row.__medicamentoColunaI || row.__medicamento || '(vazio)'}; O: ${row.__medicamentoAlternativo || '(vazio)'}; P: ${row.__principioAtivo || '(vazio)'})`;
    if (!descriptions.includes(description)) descriptions.push(description);
  }
  return descriptions.join('; ');
}


function prepareControlForAssociation(control) {
  control.__osNormalizada = normalizeOS(control.__osNormalizada || control.__os);
  control.__codBarra = normalizeBarcode(control.__codBarra);
  control.__codBarraProdutoHospital = normalizeBarcode(control.__codBarraProdutoHospital || '');
  control.__medicamentoBase = normalizeMedicineBase(control.__medicamento);
  control.__medicamentoNormalizado = normalizeMedicineProduct(control.__medicamento);
  control.__available = Math.max(0, control.__qtde);
  control.__remaining = control.__available;
  control.__usedQuantity = 0;
  control.__key = buildKey(control.__osNormalizada, control.__codBarraProdutoHospital || control.__codBarra);
}

function markOptimizationNotUsed(control, validations, message) {
  control.__status = 'Não utilizado';
  control.__observacao = message;
  control.__remaining = 0;
  validations.push(validation('Otimização sem correspondência segura', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, message));
}

function setValidationCandidateContext(control, candidates, exactReason) {
  const candidate = candidates.find((hospital) => hospital.__osNormalizada === control.__osNormalizada) ?? candidates[0] ?? null;
  control.__validationFoundOS = candidates.some((hospital) => hospital.__osNormalizada === control.__osNormalizada) ? 'SIM' : 'NÃO';
  control.__validationMedicamentoHospitalCandidato = candidate?.__medicamentoColunaI || candidate?.__medicamento || '';
  control.__validationMedicamentoAlternativoCandidato = candidate?.__medicamentoAlternativo || '';
  control.__validationPrincipioAtivoCandidato = candidate?.__principioAtivo || '';
  control.__validationCodBarraCandidato = candidate?.__codBarra || '';
  control.__validationMotivoRecusa = exactReason;
}

function normalizePatientName(value) {
  return normalizeText(value)
    .replace(/\b(DE|DA|DO|DAS|DOS|E)\b/g, ' ')
    .replace(/Z/g, 'S')
    .replace(/[^A-Z0-9 ]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function patientCompatible(a, b) {
  const first = normalizePatientName(a);
  const second = normalizePatientName(b);
  if (!first || !second) return false;
  if (first === second) return true;
  if (first.includes(second) || second.includes(first)) return true;

  const firstTokens = first.split(' ').filter((token) => token.length > 2);
  const secondTokens = second.split(' ').filter((token) => token.length > 2);

  const overlap = firstTokens.filter((token) => secondTokens.includes(token)).length;
  const minTokens = Math.min(firstTokens.length, secondTokens.length);

  if (minTokens === 0) return false;

  return overlap >= Math.min(2, minTokens) && overlap / minTokens >= 0.6;
}

function destinationCompatibleWithHospital(control, hospital) {
  const destination = normalizeText(control.__unidadeDestino);
  if (!destination) return false;

  const hospitalNames = [hospital.__cliente, hospital.__hospital, hospital.__unidade, hospital.__nomeHospital]
    .map(normalizeText)
    .filter(Boolean);

  if (!hospitalNames.length) return false;
  return hospitalNames.some((hospitalName) => optimizationDestinationCompatible(destination, hospitalName));
}

function findHospitalCandidatesForControl(control, hospitalRows) {
  const sameOS = hospitalRows.filter((hospital) =>
    hospital.__osNormalizada &&
    hospital.__osNormalizada === control.__osNormalizada
  );

  const strongCandidates = sameOS.filter((hospital) =>
    sameDayAndMonth(control.__data, hospital.__data) &&
    medicineCompatible(control.__medicamentoBase || control.__medicamentoNormalizado, hospital)
  );

  if (strongCandidates.length) {
    return {
      candidates: strongCandidates,
      matchType: 'MATCH FORTE',
      confidence: 'Alta',
      reason: 'OS normalizada, data e medicamento compatíveis.',
      osDivergence: false,
    };
  }

  const patientDateMedicineCandidates = hospitalRows.filter((hospital) =>
    sameDayAndMonth(control.__data, hospital.__data) &&
    patientCompatible(control.__paciente, hospital.__paciente) &&
    medicineCompatible(control.__medicamentoBase || control.__medicamentoNormalizado, hospital) &&
    destinationCompatibleWithHospital(control, hospital)
  );

  if (patientDateMedicineCandidates.length) {
    return {
      candidates: patientDateMedicineCandidates,
      matchType: 'MATCH MÉDIO AUDITADO',
      confidence: 'Média',
      reason: 'OS divergente; candidato localizado por paciente, data e medicamento compatíveis.',
      osDivergence: true,
    };
  }

  const dateCompatibleRows = sameOS.filter((hospital) => sameDayAndMonth(control.__data, hospital.__data));
  const medicineCompatibleInOS = sameOS.some((hospital) => medicineCompatible(control.__medicamentoBase || control.__medicamentoNormalizado, hospital));
  let reason = 'Nenhum candidato encontrado por OS ou por paciente + data + medicamento.';
  if (sameOS.length && !dateCompatibleRows.length && medicineCompatibleInOS) {
    reason = 'OS encontrada, medicamento compatível, mas data do controle não coincide com dia/mês da data do hospital.';
  } else if (sameOS.length && !dateCompatibleRows.length) {
    reason = 'OS encontrada, mas data do controle não coincide com dia/mês da data do hospital.';
  } else if (sameOS.length && !medicineCompatibleInOS) {
    reason = 'OS encontrada, mas Medicamento Base do controle não corresponde a Medicamento Hospital, Medicamento Alternativo ou PrincipioAtivo.';
  }

  return {
    candidates: [],
    matchType: '',
    confidence: '',
    reason,
    osDivergence: false,
  };
}

function destinationSeemsCompatible(destination, hospitalName) {
  return optimizationDestinationCompatible(destination, hospitalName);
}

function optimizationDestinationCompatible(destination, evaluatedHospitalUnit) {
  const hospital = normalizeHospitalUnitForComparison(evaluatedHospitalUnit);
  const dest = normalizeHospitalUnitForComparison(destination);
  if (!hospital || !dest) return false;
  if (dest === hospital || dest.includes(hospital) || hospital.includes(dest)) return true;

  const hospitalTokens = significantDestinationTokens(hospital);
  const destinationTokens = significantDestinationTokens(dest);
  if (!hospitalTokens.length || !destinationTokens.length) return false;
  const overlap = destinationTokens.filter((token) => hospitalTokens.includes(token));
  return overlap.length >= Math.min(2, hospitalTokens.length, destinationTokens.length);
}

function normalizeHospitalUnitForComparison(value) {
  return normalizeText(value)
    .replace(/\bS\s*A\b/g, ' ')
    .replace(/\bS\s*\/\s*A\b/g, ' ')
    .replace(/\b(HOSPITAL|UNIDADE|REDE|AMERICAS|AMERICA|LTDA|EIRELI)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function resolveEvaluatedHospitalUnit(hospitalRows, hospitalName, validations) {
  const override = String(hospitalName ?? '').trim();
  if (override) return override;

  const counts = new Map();
  for (const hospital of hospitalRows) {
    const value = String(hospital.__cliente || hospital.__hospital || hospital.__unidade || hospital.__nomeHospital || '').trim();
    const normalized = normalizeText(value);
    if (!normalized) continue;
    const current = counts.get(normalized) ?? { value, count: 0 };
    current.count += 1;
    counts.set(normalized, current);
  }

  const inferred = [...counts.values()].sort((a, b) => b.count - a.count)[0]?.value ?? '';
  if (inferred) return inferred;

  if (!validations.some((item) => item.Mensagem === 'Unidade avaliada não identificada; otimizações não aplicadas por segurança.')) {
    validations.push({ Tipo: 'Filtro de otimização', Severidade: 'BLOQUEIO', Aba: 'HOSPITAL_ORIGINAL', Linha: 'todas', OS_Normalizada: '', CodBarra: '', Mensagem: 'Unidade avaliada não identificada; otimizações não aplicadas por segurança.' });
  }
  return '';
}

function significantDestinationTokens(value) {
  const ignored = new Set(['HOSPITAL', 'UNIDADE', 'CENTRO', 'MEDICO', 'MEDICA', 'SAO', 'SANTO', 'SANTA', 'DE', 'DA', 'DO', 'DAS', 'DOS', 'E']);
  return normalizeText(value).split(' ').filter((token) => token.length >= 4 && !ignored.has(token));
}

function markOptimizationDestinationIneligible(control, validations) {
  const message = 'Unidade de Destino da otimização não corresponde à unidade avaliada.';
  control.__status = 'Não elegível - destino diferente';
  control.__remaining = 0;
  control.__codBarraProdutoHospital = '';
  control.__key = '';
  control.__tipoMatch = '';
  control.__confianca = '';
  control.__observacao = message;
  validations.push(validation('Unidade de Destino incompatível', 'BLOQUEIO', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, message));
}

function markOptimizationUnitUnknown(control, validations) {
  const message = 'Unidade avaliada não identificada; otimizações não aplicadas por segurança.';
  control.__status = 'Não elegível - unidade avaliada não identificada';
  control.__remaining = 0;
  control.__codBarraProdutoHospital = '';
  control.__key = '';
  control.__tipoMatch = '';
  control.__confianca = '';
  control.__observacao = message;
  validations.push(validation('Filtro de otimização', 'BLOQUEIO', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, message));
}

function unusedOptimizationReason(control, hospitalRows) {
  const base = {
    type: 'Otimização não utilizada',
    severity: 'ALERTA',
  };
  if (!Number.isFinite(control.__available) || control.__available <= 0) {
    return { ...base, message: validationDetails(control, 'quantidade inválida') };
  }
  if (!String(control.__lote ?? '').trim()) {
    return { ...base, message: validationDetails(control, 'lote vazio') };
  }
  const sameOS = hospitalRows.filter((hospital) => hospital.__osNormalizada === control.__osNormalizada);
  if (!sameOS.length) {
    return { ...base, message: validationDetails(control, 'OS não encontrada no hospital') };
  }
  const compatibleMedicine = sameOS.filter((hospital) => medicineCompatible(control.__medicamentoBase || control.__medicamentoNormalizado, hospital));
  if (!compatibleMedicine.length) {
    return { ...base, message: validationDetails(control, 'OS encontrada, mas medicamento diferente') };
  }
  if (control.__remaining <= 0) {
    return { ...base, message: validationDetails(control, 'já consumida por outra linha') };
  }
  return { ...base, message: validationDetails(control, 'já consumida por outra linha') };
}

function validationDetails(control, exactReason) {
  return `OS controle: ${control.__os || '(vazia)'}; Medicamento controle: ${control.__medicamento || '(vazio)'}; Quantidade: ${control.__available ?? control.__qtde ?? 0}; Lote: ${control.__lote || '(vazio)'}; Motivo de não uso: ${exactReason}.`;
}


function associationRecord(hospital, control, quantity, status, reason, availableBefore = 0) {
  return {
    osHospital: hospital.__os,
    osNormalizada: hospital.__osNormalizada,
    osControle: control?.__os ?? '',
    osNormalizadaControle: control?.__osNormalizada ?? '',
    codBarraHospital: hospital.__codBarra,
    codBarra: hospital.__codBarra,
    key: hospital.__key,
    dataHospital: hospital.__data,
    dataControle: control?.__data ?? '',
    medicamentoHospital: hospital.__medicamento,
    medicamentoControle: control?.__medicamento ?? '',
    medicamentoNormalizadoHospital: hospital.__medicamentoNormalizado ?? normalizeMedicineProduct(hospital.__medicamento),
    medicamentoNormalizadoControle: control?.__medicamentoNormalizado ?? '',
    principioAtivo: hospital.__principioAtivo,
    qtdePrescrita: hospital.__qtde,
    qtdeUsada: quantity,
    qtdeOtimizada: quantity,
    qtdeDisponivelAntes: availableBefore,
    loteHospital: hospital.__lote,
    origemOtimizacao: control?.__unidadeOrigem ?? '',
    validadeOtimizacao: control?.__validade ?? '',
    loteOtimizacao: control?.__lote ?? '',
    loteOtimizacaoComValidade: formatLotWithValidity(control?.__lote, control?.__validade),
    tipoMatch: control?.__tipoMatch ?? '',
    confianca: control?.__confianca ?? '',
    observacao: reason,
    motivo: reason,
    statusAssociacao: status || optimizationStatus(quantity, hospital.__qtde, availableBefore),
    hospitalRowId: hospital.__rowId,
    controlRowId: control?.__rowId ?? '',
  };
}

function optimizationStatus(optimizedQuantity, prescribedQuantity, availableBefore) {
  return optimizedQuantity > 0 ? 'Com otimização' : 'Sem otimização';
}

function summarizeAssociations(associations) {
  const summaries = new Map();
  for (const association of associations) {
    const current = summaries.get(association.hospitalRowId) ?? {
      qtdeOtimizada: 0,
      origens: [],
      lotes: [],
      lotesComValidade: [],
      validades: [],
      status: 'Sem otimização',
      availableBefore: association.qtdeDisponivelAntes ?? 0,
    };
    current.qtdeOtimizada += association.qtdeUsada ?? 0;
    current.availableBefore = Math.max(current.availableBefore, association.qtdeDisponivelAntes ?? 0);
    pushUniqueText(current.origens, association.origemOtimizacao);
    pushUniqueText(current.lotes, association.loteOtimizacao);
    pushUniqueText(current.lotesComValidade, association.loteOtimizacaoComValidade);
    pushUniqueText(current.validades, association.validadeOtimizacao ? formatDate(association.validadeOtimizacao) : '');
    current.status = optimizationStatus(current.qtdeOtimizada, association.qtdePrescrita, current.availableBefore);
    summaries.set(association.hospitalRowId, current);
  }
  return summaries;
}

function pushUniqueText(values, value) {
  const text = String(value ?? '').trim();
  if (text && !values.includes(text)) values.push(text);
}

function appendUniqueText(current, value) {
  const text = String(value ?? '').trim();
  if (!text) return current ?? '';
  const parts = String(current ?? '').split(';').map((part) => part.trim()).filter(Boolean);
  if (!parts.includes(text)) parts.push(text);
  return parts.join('; ');
}

function buildOutputWorkbook(hospitalRows, controlRows, associations, validations, originalHospitalSheet, optimizationDiagnostics = []) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Christian Web App';
  workbook.created = new Date();

  copySheetValues(originalHospitalSheet, workbook.addWorksheet(REQUIRED_SHEETS[0]));

  const controlColumns = Object.keys(stripInternal(controlRows[0] ?? {})).concat(['OS Normalizada', 'Medicamento Base', 'CodBarra Produto Hospital', 'Chave OS+CodBarra', 'Unidade de Origem capturada', 'Validade capturada', 'Lote Otimização com Validade', 'OS Hospital Associada', 'Data Associação', 'Tipo de Match', 'Confiança da Associação', 'Status Associação', 'Observação']);
  addJsonSheet(workbook, REQUIRED_SHEETS[1], controlRows.map((row) => ({
    ...stripInternal(row),
    'OS Normalizada': row.__osNormalizada,
    'Medicamento Base': row.__medicamentoBase,
    'CodBarra Produto Hospital': row.__codBarraProdutoHospital,
    'Chave OS+CodBarra': row.__key,
    'Unidade de Origem capturada': row.__unidadeOrigem,
    'Validade capturada': row.__validade,
    'Lote Otimização com Validade': formatLotWithValidity(row.__lote, row.__validade),
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
    'Origem da Otimização': item.origemOtimizacao,
    'Lote Otimização': item.loteOtimizacao,
    'Validade Otimização': item.validadeOtimizacao,
    'Lote Otimização com Validade': item.loteOtimizacaoComValidade,
    'Tipo de Match': item.tipoMatch,
    'Confiança da Associação': item.confianca,
    Observação: item.observacao,
  })), ['OS Normalizada', 'CodBarra', 'Chave OS+CodBarra', 'Data Hospital', 'Data Controle', 'Medicamento Hospital', 'Medicamento Controle', 'PrincipioAtivo', 'Qtde Prescrita', 'Qtde Otimizada', 'Lote Hospital', 'Origem da Otimização', 'Lote Otimização', 'Validade Otimização', 'Lote Otimização com Validade', 'Tipo de Match', 'Confiança da Associação', 'Observação']);

  const associationByHospitalRow = summarizeAssociations(associations);
  const baixarRows = hospitalRows.filter((row) => !row.__isExcluded).map((row) => downloadRecord(row, associationByHospitalRow.get(row.__rowId)));
  addJsonSheet(workbook, REQUIRED_SHEETS[3], baixarRows);
  addJsonSheet(workbook, REQUIRED_SHEETS[4], baixarRows.map((row) => ({
    Data: row.Data,
    Cliente: row.Cliente,
    Paciente: row.Paciente,
    'Medicamento Hospital': row['Medicamento Hospital'],
    'Qtde Prescrita': row['Qtde Prescrita'],
    'Qtde Otimizada': row['Qtde Otimizada'],
    'Origem da Otimização': row['Origem da Otimização'],
    'Lote Otimização': row['Lote Otimização'],
    'Qtde Baixa': row['Qtde Baixa'],
    'Lote da Baixa': row['Lote da Baixa'],
    'Status Otimização': row['Status Otimização'],
  })), ['Data', 'Cliente', 'Paciente', 'Medicamento Hospital', 'Qtde Prescrita', 'Qtde Otimizada', 'Origem da Otimização', 'Lote Otimização', 'Qtde Baixa', 'Lote da Baixa', 'Status Otimização']);
  addJsonSheet(workbook, REQUIRED_SHEETS[5], validations);
  addJsonSheet(workbook, REQUIRED_SHEETS[6], optimizationDiagnostics, DIAGNOSTIC_COLUMNS);

  workbook.worksheets.forEach(formatWorksheet);
  return workbook;
}

function downloadRecord(row, association) {
  const optimizedQuantity = Math.min(row.__qtde, association?.qtdeOtimizada ?? 0);
  return {
    Data: row.__data,
    Cliente: row.__cliente,
    Paciente: row.__paciente,
    'Medicamento Hospital': row.__medicamento,
    OS: row.__os,
    'OS Normalizada': row.__osNormalizada,
    CodBarra: row.__codBarra,
    'Chave OS+CodBarra': row.__key,
    'Qtde Prescrita': row.__qtde,
    'Qtde Otimizada': optimizedQuantity,
    'Origem da Otimização': association?.origens?.length ? unique(association.origens).join('; ') : 'Sem otimização',
    'Lote Otimização': association?.lotesComValidade?.length ? unique(association.lotesComValidade).join('; ') : 'Sem otimização',
    'Qtde Baixa': row.__qtde - optimizedQuantity,
    'Lote da Baixa': row.__lote,
    'Status Otimização': association?.status ?? 'Sem otimização',
  };
}


function formatLotWithValidity(lote, validade) {
  const lotText = String(lote ?? '').trim();
  if (!lotText) return 'Sem otimização';
  const validityText = String(formatDate(validade) ?? '').trim();
  return validityText ? `${lotText} - Val.: ${validityText}` : lotText;
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

function getByPosition(object, position) {
  return Object.entries(object)[position - 1]?.[1] ?? null;
}

function getPhysicalCell(row, position) {
  return normalizeCell(row.getCell(position).value);
}

function getByExactHeader(object, candidates) {
  const entries = Object.entries(object);
  for (const candidate of candidates) {
    const direct = entries.find(([key]) => normalizeText(key) === normalizeText(candidate));
    if (direct) return direct[1];
  }
  return null;
}

function shouldUseHospitalMedicineFallback(value) {
  const normalized = normalizeMedicineNameForComparison(value);
  return !normalized || normalized === 'MG';
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


function getByExactHeaderOrPosition(object, candidates, fallbackIndex) {
  const entries = Object.entries(object);
  for (const candidate of candidates) {
    const direct = entries.find(([key]) => normalizeText(key) === normalizeText(candidate));
    if (direct) return direct[1];
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
  return {
    Tipo: type,
    Severidade: severity,
    Aba: sheet,
    Linha: line,
    'OS controle': row.__os ?? '',
    OS_Normalizada: row.__osNormalizada,
    'OS normalizada controle': row.__osNormalizada ?? '',
    CodBarra: row.__codBarra,
    'Medicamento controle': row.__medicamento ?? '',
    'Medicamento Base controle': row.__medicamentoBase ?? '',
    'Data controle': row.__data ?? '',
    Lote: row.__lote ?? '',
    Quantidade: row.__available ?? row.__qtde ?? '',
    'Unidade destino': row.__unidadeDestino ?? '',
    'Unidade avaliada': row.__unidadeAvaliada ?? '',
    'Encontrou OS no hospital': row.__validationFoundOS ?? '',
    'Medicamento Hospital candidato coluna I': row.__validationMedicamentoHospitalCandidato ?? '',
    'Medicamento Alternativo candidato coluna O': row.__validationMedicamentoAlternativoCandidato ?? '',
    'PrincipioAtivo candidato coluna P': row.__validationPrincipioAtivoCandidato ?? '',
    'CodBarra candidato coluna V': row.__validationCodBarraCandidato ?? '',
    'Motivo exato da recusa': row.__validationMotivoRecusa ?? message,
    'Quantidade disponível': row.__remaining ?? row.__available ?? row.__qtde ?? '',
    Mensagem: message,
  };
}

function isEmptyObject(object) {
  return Object.values(object).every((value) => value == null || value === '');
}

function stripInternal(row) {
  return Object.fromEntries(Object.entries(row).filter(([key]) => !key.startsWith('__')));
}

function normalizeOS(value) {
  return onlyDigits(removeExcelDecimal(value)).slice(0, 7);
}

function normalizeBarcode(value) {
  return onlyDigits(removeExcelDecimal(value));
}

function removeExcelDecimal(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim().replace(/\.0$/, '');
}

function onlyDigits(value) {
  return String(value ?? '').replace(/\D/g, '');
}

function normalizeMedicineBase(value) {
  const beforeDash = String(value ?? '').split('-')[0];
  return normalizeMedicineProduct(beforeDash).trim();
}

function normalizeText(value) {
  return String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
}

function normalizeMedicineProduct(value) {
  return normalizeText(value)
    .replace(/[-_/\\]+/g, ' ')
    .replace(/(\d+(?:[.,]\d+)?)\s*(MG|G|MCG|ML|L|UI)\b/g, '$1$2')
    .replace(/\s*%/g, '%')
    .replace(/[^A-Z0-9%]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function medicineCompatible(controlMedicine, hospital) {
  const control = normalizeMedicineNameForComparison(controlMedicine);
  if (!control) return false;
  const candidates = [
    hospital.__medicamento,
    hospital.__produtoHospital,
    hospital.__medicamentoColunaI,
    hospital.__medicamentoAlternativo,
    hospital.__principioAtivo,
  ];
  return candidates.some((candidate) => compatibleNormalizedMedicine(control, normalizeMedicineNameForComparison(candidate)));
}

function normalizeMedicineNameForComparison(value) {
  const beforeDash = String(value ?? '').split('-')[0];
  return normalizeMedicineProduct(beforeDash)
    .replace(/\b\d+(?:[.,]\d+)?(?:MG|G|MCG|ML|L|UI)\b/g, '')
    .replace(/\b\d+(?:[.,]\d+)?\s*(?:MG|G|MCG|ML|L|UI)\b/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function compatibleNormalizedMedicine(control, candidate) {
  if (!control || !candidate) return false;
  if (control === candidate) return true;
  if (equivalentMedicine(control, candidate)) return true;
  return candidate.includes(control) || control.includes(candidate);
}

function equivalentMedicine(control, candidate) {
  return MEDICINE_EQUIVALENCE_GROUPS.some((group) => {
    const normalizedGroup = group.map(normalizeMedicineNameForComparison);
    return normalizedGroup.includes(control) && normalizedGroup.includes(candidate);
  });
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

function sameDayAndMonth(a, b) {
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

export {
  associateRows,
  buildKey,
  downloadRecord,
  normalizeBarcode,
  normalizeMedicineBase,
  normalizeMedicineProduct,
  normalizeOS,
  formatLotWithValidity,
  optimizationStatus,
  readControlRows,
  readHospitalRows,
  summarizeAssociations,
  validation,
};
