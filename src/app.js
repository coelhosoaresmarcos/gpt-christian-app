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
      __codBarra: normalizeBarcode(getByCandidates(object, ['CodBarra', 'Código de Barras', 'Codigo de Barras', 'EAN'], 22)),
      __qtde: toNumber(getByCandidates(object, ['Qtde', 'Quantidade', 'Qtd'], 11)),
      __lote: String(getByCandidates(object, ['Lote'], 19) ?? ''),
      __os: String(getByCandidates(object, ['OS', 'Ordem de serviço', 'Ordem de Servico', 'Ordem Serviço']) ?? ''),
    };

    hospitalRow.__osNormalizada = normalizeOS(hospitalRow.__os);
    hospitalRow.__medicamentoNormalizado = normalizeMedicineProduct(hospitalRow.__medicamento);
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
      __unidadeOrigem: String(getByCandidates(object, ['Unidade Origem', 'Unidade de Origem', 'Origem'], 4) ?? ''),
      __unidadeDestino: String(getByCandidates(object, ['Unidade de Destino', 'Unidade Destino', 'Destino'], 5) ?? ''),
      __paciente: String(getByCandidates(object, ['Paciente', 'Nome Paciente'], 6) ?? ''),
      __medicamento: String(getByCandidates(object, ['Medicamento', 'Produto'], 7) ?? ''),
      __qtde: toNumber(getByCandidates(object, ['Quantidade (mg)', 'Quantidade', 'Qtde', 'Qtd'], 8)),
      __lote: String(getByCandidates(object, ['Lote'], 9) ?? ''),
      __validade: getByCandidates(object, ['Validade', 'Vencimento'], 10),
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

function associateRows(hospitalRows, controlRows, validations, hospitalName) {
  const associations = [];
  const usableHospitalRows = hospitalRows.filter((row) => !row.__isExcluded);
  const controlsByKey = new Map();
  const keyUsage = new Map();

  if (!hospitalName) {
    validations.push({ Tipo: 'Filtro de otimização', Severidade: 'ALERTA', Aba: 'CONTROLE DE MEDICAMENTOS', Linha: 'todas', OS_Normalizada: '', CodBarra: '', Mensagem: 'Nome do hospital não informado; o filtro por Unidade de Destino não foi aplicado.' });
  }

  for (const hospital of usableHospitalRows) {
    hospital.__osNormalizada = normalizeOS(hospital.__osNormalizada || hospital.__os);
    hospital.__codBarra = normalizeBarcode(hospital.__codBarra);
    hospital.__key = buildKey(hospital.__osNormalizada, hospital.__codBarra);
    hospital.__medicamentoNormalizado = normalizeMedicineProduct(hospital.__medicamento);
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
    const destinationOk = destinationSeemsCompatible(control.__unidadeDestino, hospitalName);
    const sameOS = usableHospitalRows.filter((hospital) => hospital.__osNormalizada && hospital.__osNormalizada === control.__osNormalizada);
    const medicineCountInOS = unique(sameOS.map((hospital) => normalizeMedicineProduct(hospital.__medicamento || hospital.__principioAtivo)).filter(Boolean)).length;
    if (medicineCountInOS > 1) {
      validations.push(validation('Múltiplos medicamentos na mesma OS', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, 'Há múltiplos medicamentos na mesma OS; CodBarra será obrigatório para a associação final.'));
    }

    if (!isOptimization) {
      control.__status = 'Fora do filtro de otimização';
      control.__observacao = 'Motivo não contém OTIMIZA; linha preservada sem associação.';
      control.__remaining = 0;
      continue;
    }
    if (!destinationOk) {
      markOptimizationNotUsed(control, validations, 'Unidade de Destino não contém o hospital informado; otimização fora do filtro obrigatório.');
      continue;
    }
    if (!control.__osNormalizada) {
      markOptimizationNotUsed(control, validations, 'OS ausente ou inválida no controle; não é permitido associar sem OS Normalizada.');
      continue;
    }
    if (!Number.isFinite(control.__available) || control.__available <= 0) {
      markOptimizationNotUsed(control, validations, validationDetails(control, 'quantidade inválida'));
      continue;
    }
    if (!sameOS.length) {
      markOptimizationNotUsed(control, validations, validationDetails(control, 'OS não encontrada no hospital'));
      validations.push(validation('Divergência de OS', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, 'Nenhuma OS normalizada correspondente na HOSPITAL_ORIGINAL.'));
      continue;
    }

    const dateCompatibleRows = sameOS.filter((hospital) => sameDayAndMonth(control.__data, hospital.__data));
    if (!dateCompatibleRows.length) {
      markOptimizationNotUsed(control, validations, validationDetails(control, 'data sem coincidência de dia e mês com o hospital'));
      validations.push(validation('Otimização sem correspondência segura', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, 'OS encontrada, mas a data do controle não coincide com dia/mês da data do hospital.'));
      continue;
    }

    const compatible = dateCompatibleRows.filter((hospital) => medicineCompatible(control.__medicamentoBase || control.__medicamentoNormalizado, hospital));
    if (!compatible.length) {
      refuseControl(control, validations, 'Tentativa de associação apenas por OS', 'MATCH FRACO recusado: OS encontrada, mas medicamento não compatível.');
      validations.push(validation('Divergência de medicamento', 'BLOQUEIO', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, 'OS encontrada, mas medicamento não compatível.'));
      continue;
    }

    const possibleBarcodes = unique(compatible.map((hospital) => hospital.__codBarra).filter(Boolean));
    if (!possibleBarcodes.length) {
      refuseControl(control, validations, 'Ausência de CodBarra compatível', 'Medicamento compatível encontrado na OS, mas sem CodBarra preenchido no hospital para criar a chave final.');
      continue;
    }
    if (possibleBarcodes.length > 1) {
      refuseControl(control, validations, 'Múltiplos CodBarra possíveis na mesma OS', 'Múltiplos CodBarra possíveis para a mesma OS e medicamento.');
      validations.push(validation('Múltiplos CodBarra possíveis na mesma OS', 'BLOQUEIO', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, 'Múltiplos CodBarra possíveis para a mesma OS e medicamento.'));
      continue;
    }

    const identifiedBarcode = possibleBarcodes[0];
    if (control.__codBarra && control.__codBarra !== identifiedBarcode) {
      refuseControl(control, validations, 'Divergência de CodBarra', 'CodBarra do controle diverge do CodBarra compatível identificado no hospital.');
      continue;
    }

    const dateWithDifferentYear = compatible.find((hospital) => sameDayAndMonth(control.__data, hospital.__data) && !sameYear(control.__data, hospital.__data));
    if (dateWithDifferentYear) {
      validations.push(validation('Correção de data por divergência de ano', 'ALERTA', 'CONTROLE DE MEDICAMENTOS', control.__rowId, control, `Dia/mês coincidem; ano do controle será tratado conforme ano do hospital (${formatDate(dateWithDifferentYear.__data)}).`));
    }

    control.__codBarraProdutoHospital = identifiedBarcode;
    control.__key = buildKey(control.__osNormalizada, identifiedBarcode);
    control.__tipoMatch = control.__codBarra ? 'MATCH FORTE' : 'MATCH MÉDIO';
    control.__confianca = control.__codBarra ? 'Alta' : 'Média';
    control.__status = 'Disponível para otimização';
    control.__observacao = control.__codBarra
      ? 'CodBarra confirmado no controle; associação final autorizada pela chave OS Normalizada + CodBarra.'
      : 'CodBarra Produto Hospital identificado por OS normalizada e medicamento compatível; associação final autorizada pela chave OS Normalizada + CodBarra.';
    control.__hospitalAssociado = appendUniqueText(control.__hospitalAssociado, unique(compatible.map((hospital) => hospital.__os)).join('; '));
    control.__dataAssociacao = appendUniqueText(control.__dataAssociacao, unique(compatible.map((hospital) => formatDate(hospital.__data))).join('; '));

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
  return associations;
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

function destinationSeemsCompatible(destination, hospitalName) {
  const hospital = normalizeText(hospitalName);
  if (!hospital) return true;
  const dest = normalizeText(destination);
  if (!dest) return false;
  return dest.includes(hospital) || hospital.includes(dest);
}

function significantDestinationTokens(value) {
  const ignored = new Set(['HOSPITAL', 'UNIDADE', 'CENTRO', 'MEDICO', 'MEDICA', 'SAO', 'SANTO', 'SANTA', 'DE', 'DA', 'DO', 'DAS', 'DOS', 'E']);
  return normalizeText(value).split(' ').filter((token) => token.length >= 4 && !ignored.has(token));
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
  const compatibleMedicine = sameOS.filter((hospital) => medicineCompatible(control.__medicamentoNormalizado, hospital));
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
    loteOtimizacao: control?.__lote ?? '',
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
      lotes: [],
      status: 'Sem otimização',
      availableBefore: association.qtdeDisponivelAntes ?? 0,
    };
    current.qtdeOtimizada += association.qtdeUsada ?? 0;
    current.availableBefore = Math.max(current.availableBefore, association.qtdeDisponivelAntes ?? 0);
    if (association.loteOtimizacao) current.lotes.push(association.loteOtimizacao);
    current.status = optimizationStatus(current.qtdeOtimizada, association.qtdePrescrita, current.availableBefore);
    summaries.set(association.hospitalRowId, current);
  }
  return summaries;
}

function appendUniqueText(current, value) {
  const text = String(value ?? '').trim();
  if (!text) return current ?? '';
  const parts = String(current ?? '').split(';').map((part) => part.trim()).filter(Boolean);
  if (!parts.includes(text)) parts.push(text);
  return parts.join('; ');
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
    'CodBarra Produto Hospital': row.__codBarraProdutoHospital,
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
  })), ['OS Normalizada', 'CodBarra', 'Chave OS+CodBarra', 'Data Hospital', 'Data Controle', 'Medicamento Hospital', 'Medicamento Controle', 'PrincipioAtivo', 'Qtde Prescrita', 'Qtde Otimizada', 'Lote Hospital', 'Lote Otimização', 'Tipo de Match', 'Confiança da Associação', 'Observação']);

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
    'Lote Otimização': association?.lotes?.length ? unique(association.lotes).join('; ') : 'Sem otimização',
    'Qtde Baixa': row.__qtde - optimizedQuantity,
    'Lote da Baixa': row.__lote,
    'Status Otimização': association?.status ?? 'Sem otimização',
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
  return {
    Tipo: type,
    Severidade: severity,
    Aba: sheet,
    Linha: line,
    'OS controle': row.__os ?? '',
    OS_Normalizada: row.__osNormalizada,
    CodBarra: row.__codBarra,
    'Medicamento controle': row.__medicamento ?? '',
    Lote: row.__lote ?? '',
    Quantidade: row.__available ?? row.__qtde ?? '',
    'Unidade destino': row.__unidadeDestino ?? '',
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
  const control = normalizeMedicineProduct(controlMedicine);
  if (!control) return false;
  const medicine = normalizeMedicineProduct(hospital.__medicamento);
  const active = normalizeMedicineProduct(hospital.__principioAtivo);
  return compatibleNormalizedMedicine(control, medicine) || compatibleNormalizedMedicine(control, active);
}

function compatibleNormalizedMedicine(control, candidate) {
  if (!control || !candidate) return false;
  if (control === candidate) return true;
  return candidate.includes(control) || control.includes(candidate);
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
  normalizeBarcode,
  normalizeMedicineBase,
  normalizeMedicineProduct,
  normalizeOS,
  optimizationStatus,
  readControlRows,
  summarizeAssociations,
  validation,
};
