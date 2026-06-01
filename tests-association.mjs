import assert from 'node:assert/strict';
import { associateRows, downloadRecord, formatLotWithValidity, normalizeMedicineBase, normalizeMedicineProduct, normalizeOS, readControlRows, readHospitalRows, summarizeAssociations } from './src/app.js';

function fakeRow(values) {
  return {
    eachCell(_options, callback) {
      values.forEach((value, index) => callback({ value }, index + 1));
    },
    getCell(index) {
      return { value: values[index - 1] ?? null };
    },
  };
}

function fakeSheet(rows) {
  const fakeRows = rows.map(fakeRow);
  return {
    eachRow(optionsOrCallback, maybeCallback) {
      const callback = typeof optionsOrCallback === 'function' ? optionsOrCallback : maybeCallback;
      fakeRows.forEach((row, index) => callback(row, index + 1));
    },
    getRow(index) {
      return fakeRows[index - 1];
    },
  };
}

function hospital(overrides) {
  const osNormalizada = overrides.__osNormalizada ?? normalizeOS(overrides.__os);
  const codBarra = overrides.__codBarra ?? '7891000';
  return {
    __rowId: overrides.__rowId,
    __data: overrides.__data ?? '01/06/2026',
    __cliente: overrides.__cliente ?? 'Hospital Teste',
    __paciente: overrides.__paciente ?? 'Paciente',
    __medicamento: overrides.__medicamento ?? 'FAULDFLUOR 500MG',
    __produtoHospital: overrides.__produtoHospital ?? '',
    __medicamentoColunaI: overrides.__medicamentoColunaI ?? overrides.__medicamento ?? 'FAULDFLUOR 500MG',
    __medicamentoAlternativo: overrides.__medicamentoAlternativo ?? '',
    __principioAtivo: overrides.__principioAtivo ?? 'FAULDFLUOR 500MG',
    __codBarra: codBarra,
    __qtde: overrides.__qtde,
    __lote: overrides.__lote ?? 'LOTE-H',
    __os: overrides.__os,
    __osNormalizada: osNormalizada,
    __key: osNormalizada && codBarra ? `${osNormalizada}|${codBarra}` : '',
    __isExcluded: false,
  };
}

function control(overrides) {
  const osNormalizada = overrides.__osNormalizada ?? normalizeOS(overrides.__os);
  const codBarra = overrides.__codBarra ?? '';
  return {
    __rowId: overrides.__rowId,
    __data: overrides.__data ?? '01/06/2026',
    __os: overrides.__os,
    __osNormalizada: osNormalizada,
    __medicamento: overrides.__medicamento ?? 'FAULDFLUOR - 500 mg',
    __medicamentoBase: normalizeMedicineBase(overrides.__medicamento ?? 'FAULDFLUOR - 500 mg'),
    __medicamentoNormalizado: normalizeMedicineProduct(overrides.__medicamento ?? 'FAULDFLUOR - 500 mg'),
    __qtde: overrides.__qtde,
    __lote: overrides.__lote ?? '25D0738',
    __motivo: overrides.__motivo ?? 'OTIMIZAÇÃO',
    __unidadeOrigem: overrides.__unidadeOrigem ?? 'Centro Médico Origem',
    __unidadeDestino: overrides.__unidadeDestino ?? 'Hospital Teste',
    __validade: overrides.__validade ?? '30/04/2027',
    __codBarra: codBarra,
    __codBarraProdutoHospital: '',
    __key: '',
    __status: 'Não avaliado como otimização',
    __observacao: '',
    __tipoMatch: '',
    __confianca: '',
    __hospitalAssociado: '',
    __dataAssociacao: '',
    __used: false,
  };
}

assert.equal(normalizeOS('3130624-1'), '3130624');
assert.equal(normalizeOS('13157787'), '1315778');
assert.equal(normalizeMedicineBase('FAULDFLUOR - 500 mg'), 'FAULDFLUOR');
assert.equal(normalizeMedicineProduct('FAULDFLUOR - 500 mg'), 'FAULDFLUOR 500MG');

const fallbackControlRows = readControlRows(fakeSheet([
  ['Tipo/Motivo', 'OS', 'Data', 'Unidade Origem', 'Unidade Destino', 'Paciente', 'Medicamento', 'Quantidade', 'Lote', 'Validade', 'Laboratório'],
  ['OTIMIZAÇÃO', '13157787', '01/06/2026', 'CENTRO MEDICO PITANGUEIRAS', 'HOSPITAL SANTA HELENA', 'VLADEMIR PERNIQUELLI', 'FAULDFLUOR - 500 mg', 200, '25D0738', '30/04/2027', 'LIBBS'],
]), []);
assert.equal(fallbackControlRows[0].__os, '13157787', 'readControlRows lê OS pela posição 2 quando necessário');
assert.equal(fallbackControlRows[0].__osNormalizada, '1315778', 'OS do controle é normalizada pelos 7 primeiros dígitos');
assert.equal(fallbackControlRows[0].__medicamento, 'FAULDFLUOR - 500 mg', 'readControlRows lê medicamento pela posição 7');
assert.equal(fallbackControlRows[0].__qtde, 200, 'readControlRows lê quantidade pela posição 8');
assert.equal(fallbackControlRows[0].__lote, '25D0738', 'readControlRows lê lote pela posição 9');
assert.equal(fallbackControlRows[0].__unidadeOrigem, 'CENTRO MEDICO PITANGUEIRAS', 'readControlRows lê Unidade de Origem pela posição 4');
assert.equal(fallbackControlRows[0].__unidadeDestino, 'HOSPITAL SANTA HELENA', 'readControlRows lê Unidade de Destino pela posição 5');
assert.equal(fallbackControlRows[0].__validade, '30/04/2027', 'readControlRows lê Validade pela posição 10');
assert.equal(Object.keys(fallbackControlRows[0]).filter((key) => key === '__unidadeOrigem').length, 1, 'controlRow contém __unidadeOrigem apenas uma vez');


const fallbackHospitalRows = readHospitalRows(fakeSheet([
  ['Coluna A', 'Coluna B', 'Data', 'Paciente', 'Tipo', 'Documento', 'Unidade', 'OS', 'Coluna I', 'UM', 'Qtde', 'Desc 1', 'Desc 2', 'Setor', 'Coluna O', 'PrincipioAtivo', 'Carteira', 'Convenio', 'Lote', 'Extra 1', 'Extra 2', 'CodBarra'],
  ['HOSPITAL ANA COSTA', 'HOSPITAL ANA COSTA S/A', '01/06/2026', 'ODETE DE SOUSA', 'NFe', '1122056', 'QT', '1.315.669,00', 'GENLIBBS', 'mg', '1.300,00', '0,00', '0,00', 'AMBULATORIO', 'GENLIBBS', 'GENCITABINA', 'Corporativo', 'AMIL', 'L:26A0739 V:31/01/28; L:26A0728 V:31/01/28', '', '1082', '7896094207257'],
]), []);
assert.equal(fallbackHospitalRows[0].__os, '1.315.669,00', 'readHospitalRows lê OS pela coluna H quando cabeçalho de OS existe ou por fallback');
assert.equal(fallbackHospitalRows[0].__medicamento, 'GENLIBBS', 'readHospitalRows usa a coluna I como fallback de Medicamento Hospital');
assert.equal(fallbackHospitalRows[0].__medicamentoColunaI, 'GENLIBBS', 'readHospitalRows mantém campo interno da coluna I');
assert.equal(fallbackHospitalRows[0].__medicamentoAlternativo, 'GENLIBBS', 'readHospitalRows usa a coluna O como Medicamento Alternativo');
assert.equal(fallbackHospitalRows[0].__principioAtivo, 'GENCITABINA', 'readHospitalRows usa a coluna P como PrincipioAtivo');
assert.equal(fallbackHospitalRows[0].__codBarra, '7896094207257', 'readHospitalRows usa a coluna V como CodBarra');
assert.equal(fallbackHospitalRows[0].__qtde, 1300, 'readHospitalRows usa a coluna K como Qtde');
assert.equal(fallbackHospitalRows[0].__lote, 'L:26A0739 V:31/01/28; L:26A0728 V:31/01/28', 'readHospitalRows usa a coluna S como Lote');

const hospitalRowsWithWrongMedicineHeader = readHospitalRows(fakeSheet([
  ['Cliente', 'Coluna B', 'Data', 'Paciente', 'Tipo', 'Documento', 'Unidade', 'OS', 'Produto físico', 'Medicamento', 'Qtde', 'Desc 1', 'Desc 2', 'Setor', 'Coluna O', 'PrincipioAtivo', 'Carteira', 'Convenio', 'Lote', 'Extra 1', 'Extra 2', 'CodBarra'],
  ['HOSPITAL ANA COSTA', 'HOSPITAL ANA COSTA S/A', '01/06/2026', 'ODETE DE SOUSA', 'NFe', '1122056', 'QT', '1.315.669,00', 'GENLIBBS', 'mg', '1.300,00', '0,00', '0,00', 'AMBULATORIO', 'GENLIBBS', 'GENCITABINA', 'Corporativo', 'AMIL', 'L:26A0739 V:31/01/28; L:26A0728 V:31/01/28', '', '1082', '7896094207257'],
]), []);
assert.equal(hospitalRowsWithWrongMedicineHeader[0].__medicamento, 'GENLIBBS', 'readHospitalRows nunca usa coluna J como Medicamento Hospital quando ela contém mg');
assert.equal(hospitalRowsWithWrongMedicineHeader[0].__medicamentoAlternativo, 'GENLIBBS', 'readHospitalRows preenche Medicamento Alternativo pela coluna O mesmo com cabeçalho diferente');

assert.equal(Object.keys(fallbackControlRows[0]).filter((key) => key === '__unidadeDestino').length, 1, 'controlRow contém __unidadeDestino apenas uma vez');
assert.equal(Object.keys(fallbackControlRows[0]).filter((key) => key === '__validade').length, 1, 'controlRow contém __validade apenas uma vez');
assert.equal(Object.keys(fallbackControlRows[0]).filter((key) => key === '__lote').length, 1, 'controlRow contém __lote apenas uma vez');
assert.equal(formatLotWithValidity(fallbackControlRows[0].__lote, fallbackControlRows[0].__validade), '25D0738 - Val.: 30/04/2027', 'lote e validade são combinados no formato do relatório');
assert.equal(formatLotWithValidity('25D0738', ''), '25D0738', 'lote sem validade retorna apenas lote');
assert.equal(formatLotWithValidity('', '30/04/2027'), 'Sem otimização', 'ausência de lote retorna Sem otimização');

const positionalFallbackRows = readControlRows(fakeSheet([
  ['Tipo/Motivo', 'OS', 'Data', 'Hospital Origem', 'Hospital Destino', 'Paciente', 'Medicamento', 'Quantidade', 'Lote', 'Data Venc.', 'Laboratório'],
  ['OTIMIZAÇÃO', '13157787', '01/06/2026', 'CENTRO MEDICO PITANGUEIRAS', 'HOSPITAL SANTA HELENA', 'VLADEMIR PERNIQUELLI', 'FAULDFLUOR - 500 mg', 200, '25D0738', '30/04/2027', 'LIBBS'],
]), []);
assert.equal(positionalFallbackRows[0].__unidadeOrigem, 'CENTRO MEDICO PITANGUEIRAS', 'readControlRows usa a coluna D como fallback quando o cabeçalho de origem é diferente');
assert.equal(positionalFallbackRows[0].__unidadeDestino, 'HOSPITAL SANTA HELENA', 'readControlRows usa a coluna E como fallback quando o cabeçalho de destino é diferente');
assert.equal(positionalFallbackRows[0].__validade, '30/04/2027', 'readControlRows usa a coluna J como fallback quando o cabeçalho de validade é diferente');

const validations = [];
const hospitalRows = [
  hospital({ __rowId: 2, __os: '1234567-1', __codBarra: '7891000', __qtde: 5 }),
  hospital({ __rowId: 3, __os: '1234567-2', __codBarra: '7892000', __qtde: 5, __medicamento: 'KEYTRUDA 100MG', __principioAtivo: 'PEMBROLIZUMABE' }),
  hospital({ __rowId: 4, __os: '9999999', __codBarra: '7893000', __qtde: 4 }),
  hospital({ __rowId: 5, __os: '9999999', __codBarra: '7893000', __qtde: 4 }),
];
const controlRows = [
  control({ __rowId: 2, __os: '1234567', __codBarra: '7891000', __qtde: 5, __lote: 'OK' }),
  control({ __rowId: 3, __os: '1234567', __codBarra: '7899999', __qtde: 10, __lote: 'BARCODE-DIVERGENTE' }),
  control({ __rowId: 4, __os: '0000000', __codBarra: '7891000', __qtde: 10, __lote: 'BARCODE-ONLY-BLOCKED' }),
  control({ __rowId: 5, __os: '9999999', __qtde: 6, __lote: 'SALDO' }),
];

const associations = associateRows(hospitalRows, controlRows, validations, 'Hospital Teste');

const row2Used = associations.filter((item) => item.hospitalRowId === 2).reduce((sum, item) => sum + item.qtdeUsada, 0);
assert.equal(row2Used, 5, 'associa quando OS + CodBarra são iguais e medicamento é compatível');
assert.equal(controlRows[0].__tipoMatch, 'MATCH FORTE');

const row3Used = associations.filter((item) => item.hospitalRowId === 3).reduce((sum, item) => sum + item.qtdeUsada, 0);
assert.equal(row3Used, 0, 'não associa quando só a OS é igual e o CodBarra diverge');
assert.ok(validations.some((item) => item.Tipo === 'Tentativa de associação apenas por OS' || item.Tipo === 'Divergência de CodBarra'));

const row4Used = associations.filter((item) => item.hospitalRowId === 4).reduce((sum, item) => sum + item.qtdeUsada, 0);
const row5Used = associations.filter((item) => item.hospitalRowId === 5).reduce((sum, item) => sum + item.qtdeUsada, 0);
assert.equal(row4Used, 4, 'primeira linha da mesma chave consome saldo disponível');
assert.equal(row5Used, 2, 'segunda linha da mesma chave usa somente saldo remanescente');
assert.equal(controlRows[3].__codBarraProdutoHospital, '7893000', 'controle sem CodBarra recebe CodBarra do hospital');
assert.equal(controlRows[3].__key, '9999999|7893000', 'controle cria chave final OS Normalizada + CodBarra');

const mandatoryValidations = [];
const mandatoryHospitalRows = [
  hospital({ __rowId: 7, __os: '13157787', __codBarra: '7895000', __qtde: 300, __cliente: 'HOSPITAL SANTA HELENA', __paciente: 'VLADEMIR PERNIQUELLI', __medicamento: 'FAULDFLUOR 500MG', __principioAtivo: 'FAULDFLUOR', __lote: 'LOTE-HOSPITAL' }),
];
const mandatoryControlRows = [
  control({ __rowId: 6, __os: '13157787', __qtde: 200, __lote: '25D0738', __validade: '30/04/2027', __data: '01/06/2026', __unidadeOrigem: 'CENTRO MEDICO PITANGUEIRAS', __unidadeDestino: 'HOSPITAL SANTA HELENA', __medicamento: 'FAULDFLUOR - 500 mg' }),
];
const mandatoryAssociations = associateRows(mandatoryHospitalRows, mandatoryControlRows, mandatoryValidations, 'Hospital Santa Helena');
const mandatoryAssociation = mandatoryAssociations.find((item) => item.hospitalRowId === 7 && item.controlRowId === 6);
assert.ok(mandatoryAssociation, 'associa a linha OTIMIZAÇÃO do controle após descobrir CodBarra do hospital');
assert.equal(mandatoryHospitalRows[0].__osNormalizada, '1315778', 'OS 13157787 deve ser truncada para os 7 primeiros dígitos');
assert.equal(mandatoryControlRows[0].__osNormalizada, '1315778', 'OS 13157787 do controle deve ser truncada para os 7 primeiros dígitos');
assert.equal(mandatoryControlRows[0].__codBarraProdutoHospital, '7895000', 'controle recebe o CodBarra Produto Hospital');
assert.equal(mandatoryControlRows[0].__key, '1315778|7895000', 'controle recebe chave OS_NORMALIZADA|CODBARRA');
assert.equal(mandatoryAssociation.key, '1315778|7895000');
assert.equal(mandatoryAssociation.qtdeOtimizada, 200, 'Qtde Otimizada vem da quantidade registrada na linha OTIMIZAÇÃO, limitada pela prescrição');
assert.equal(mandatoryAssociation.origemOtimizacao, 'CENTRO MEDICO PITANGUEIRAS', 'Origem da Otimização vem da mesma linha OTIMIZAÇÃO do controle');
assert.equal(mandatoryAssociation.loteOtimizacao, '25D0738', 'Lote Otimização vem exatamente da linha OTIMIZAÇÃO do controle');
assert.equal(mandatoryAssociation.validadeOtimizacao, '30/04/2027', 'Validade Otimização vem da mesma linha OTIMIZAÇÃO do controle');
assert.equal(mandatoryAssociation.loteOtimizacaoComValidade, '25D0738 - Val.: 30/04/2027', 'Lote Otimização com Validade combina lote e validade da mesma linha do controle');
assert.equal(mandatoryAssociation.statusAssociacao, 'Com otimização');
assert.equal(mandatoryAssociation.tipoMatch, 'MATCH MÉDIO');
assert.equal(mandatoryAssociation.confianca, 'Média');
const mandatoryReportSummary = summarizeAssociations(mandatoryAssociations).get(7);
assert.deepEqual(mandatoryReportSummary.origens, ['CENTRO MEDICO PITANGUEIRAS'], 'RELATORIO/BAIXAR usa a origem da otimização resumida');
assert.deepEqual(mandatoryReportSummary.lotes, ['25D0738'], 'RELATORIO/BAIXAR mantém o lote 25D0738 resumido para auditoria');
assert.deepEqual(mandatoryReportSummary.lotesComValidade, ['25D0738 - Val.: 30/04/2027'], 'RELATORIO/BAIXAR usa lote e validade resumidos para a otimização');
assert.equal(mandatoryReportSummary.status, 'Com otimização', 'RELATORIO/BAIXAR usa status Com otimização');
const mandatoryDownloadRecord = downloadRecord(mandatoryHospitalRows[0], mandatoryReportSummary);
assert.equal(mandatoryDownloadRecord['Origem da Otimização'], 'CENTRO MEDICO PITANGUEIRAS', 'RELATORIO mostra a origem da otimização vinda da coluna D');
assert.equal(mandatoryDownloadRecord['Lote Otimização'], '25D0738 - Val.: 30/04/2027', 'RELATORIO mostra o lote da otimização com a validade da coluna J');
assert.equal(mandatoryDownloadRecord['Status Otimização'], 'Com otimização', 'RELATORIO mostra status Com otimização quando houve consumo');
assert.equal(mandatoryHospitalRows[0].__qtde - mandatoryAssociation.qtdeOtimizada, 100, 'Qtde Baixa esperada é 300 - 200 = 100');
const genlibbsValidations = [];
const genlibbsHospitalRows = [
  hospital({ __rowId: 9, __os: '1.315.669,00', __codBarra: '7896094207257', __qtde: 1300, __cliente: 'HOSPITAL ANA COSTA', __paciente: 'ODETE DE SOUSA', __medicamento: 'GENLIBBS', __medicamentoAlternativo: 'GENLIBBS', __principioAtivo: 'GENCITABINA', __lote: 'L:26A0739 V:31/01/28; L:26A0728 V:31/01/28' }),
];
const genlibbsControlRows = [
  control({ __rowId: 10, __os: '1.315.669,00', __qtde: 100, __lote: '26A0739', __validade: '31/01/2028', __data: '01/06/2026', __unidadeOrigem: 'AMBULATORIO', __unidadeDestino: 'HOSPITAL ANA COSTA', __medicamento: 'GENLIBBS - qualquer concentração' }),
];
const genlibbsAssociations = associateRows(genlibbsHospitalRows, genlibbsControlRows, genlibbsValidations, 'Hospital Ana Costa');
assert.equal(genlibbsControlRows[0].__medicamentoBase, 'GENLIBBS', 'controle GENLIBBS cria Medicamento Base apenas com texto antes do hífen');
assert.equal(genlibbsControlRows[0].__codBarraProdutoHospital, '7896094207257', 'controle GENLIBBS localiza CodBarra Produto Hospital por coluna I/O');
assert.equal(genlibbsControlRows[0].__key, '1315669|7896094207257', 'controle GENLIBBS cria chave final OS_NORMALIZADA|CodBarra');
assert.ok(genlibbsAssociations.find((item) => item.hospitalRowId === 9 && item.controlRowId === 10), 'controle GENLIBBS associa somente após criar chave OS+CodBarra');
assert.equal(genlibbsAssociations[0].statusAssociacao, 'Com otimização', 'controle GENLIBBS gera status Com otimização');

const genlibbsEquivalenceValidations = [];
const genlibbsEquivalenceHospitalRows = [
  hospital({ __rowId: 16, __os: '1.315.669,00', __codBarra: '7896094207257', __qtde: 1300, __medicamento: '', __medicamentoColunaI: '', __medicamentoAlternativo: '', __principioAtivo: 'GENCITABINA' }),
];
const genlibbsEquivalenceControlRows = [
  control({ __rowId: 17, __os: '1.315.669,00', __qtde: 100, __medicamento: 'GENLIBBS - qualquer concentração', __unidadeDestino: 'HOSPITAL ANA COSTA' }),
];
const genlibbsEquivalenceAssociations = associateRows(genlibbsEquivalenceHospitalRows, genlibbsEquivalenceControlRows, genlibbsEquivalenceValidations, 'Hospital Ana Costa');
assert.equal(genlibbsEquivalenceControlRows[0].__codBarraProdutoHospital, '7896094207257', 'controle GENLIBBS localiza CodBarra por equivalência com GENCITABINA');
assert.ok(genlibbsEquivalenceAssociations.find((item) => item.hospitalRowId === 16 && item.controlRowId === 17), 'equivalência GENLIBBS/GENCITABINA serve para identificar CodBarra antes da associação OS+CodBarra');

const mismatchValidations = [];
associateRows(
  [hospital({ __rowId: 11, __os: '1234567', __codBarra: '7897000', __qtde: 10, __medicamento: 'GENLIBBS', __medicamentoAlternativo: 'GENLIBBS', __principioAtivo: 'GENCITABINA' })],
  [control({ __rowId: 12, __os: '1234567', __qtde: 1, __medicamento: 'KEYTRUDA - 100MG' })],
  mismatchValidations,
  'Hospital Teste',
);
const mismatchValidation = mismatchValidations.find((item) => item.Mensagem === 'OS encontrada, mas Medicamento Base do controle não corresponde a Medicamento Hospital, Medicamento Alternativo ou PrincipioAtivo.');
assert.ok(mismatchValidation, 'registra mensagem específica quando OS existe mas Medicamento Base não corresponde');
assert.equal(mismatchValidation['OS controle'], '1234567', 'VALIDACAO mostra OS controle quando recusa uma otimização');
assert.equal(mismatchValidation['OS normalizada controle'], '1234567', 'VALIDACAO mostra OS normalizada controle quando recusa uma otimização');
assert.equal(mismatchValidation['Medicamento controle'], 'KEYTRUDA - 100MG', 'VALIDACAO mostra Medicamento controle quando recusa uma otimização');
assert.equal(mismatchValidation['Medicamento Base controle'], 'KEYTRUDA', 'VALIDACAO mostra Medicamento Base controle quando recusa uma otimização');
assert.equal(mismatchValidation['Encontrou OS no hospital'], 'SIM', 'VALIDACAO mostra se encontrou OS no hospital quando recusa uma otimização');
assert.equal(mismatchValidation['Medicamento Hospital candidato coluna I'], 'GENLIBBS', 'VALIDACAO mostra candidato da coluna I quando recusa uma otimização');
assert.equal(mismatchValidation['Medicamento Alternativo candidato coluna O'], 'GENLIBBS', 'VALIDACAO mostra candidato da coluna O quando recusa uma otimização');
assert.equal(mismatchValidation['PrincipioAtivo candidato coluna P'], 'GENCITABINA', 'VALIDACAO mostra PrincipioAtivo candidato quando recusa uma otimização');
assert.equal(mismatchValidation['CodBarra candidato coluna V'], '7897000', 'VALIDACAO mostra CodBarra candidato quando recusa uma otimização');
assert.equal(mismatchValidation['Motivo exato da recusa'], 'OS encontrada, mas medicamento diferente', 'VALIDACAO mostra motivo exato da recusa quando recusa uma otimização');

const multiBarcodeValidations = [];
associateRows(
  [
    hospital({ __rowId: 13, __os: '7654321', __codBarra: '7898000', __qtde: 10, __medicamento: 'GENLIBBS', __medicamentoAlternativo: 'GENLIBBS', __principioAtivo: 'GENCITABINA' }),
    hospital({ __rowId: 14, __os: '7654321', __codBarra: '7898001', __qtde: 10, __medicamento: 'GENLIBBS', __medicamentoAlternativo: 'GENLIBBS', __principioAtivo: 'GENCITABINA' }),
  ],
  [control({ __rowId: 15, __os: '7654321', __qtde: 1, __medicamento: 'GENLIBBS - 1000 mg' })],
  multiBarcodeValidations,
  'Hospital Teste',
);
assert.ok(multiBarcodeValidations.some((item) => item.Mensagem === 'Múltiplos CodBarra possíveis para mesma OS e Medicamento Base.'), 'bloqueia associação automática com múltiplos CodBarra para mesma OS e Medicamento Base');


const noOptimizationRecord = downloadRecord(hospital({ __rowId: 8, __os: '7654321', __codBarra: '7896000', __qtde: 10 }), undefined);
assert.equal(noOptimizationRecord['Origem da Otimização'], 'Sem otimização', 'RELATORIO mantém origem Sem otimização sem associação');
assert.equal(noOptimizationRecord['Lote Otimização'], 'Sem otimização', 'RELATORIO mantém lote Sem otimização sem associação');
assert.equal(noOptimizationRecord['Status Otimização'], 'Sem otimização', 'RELATORIO mantém status Sem otimização sem associação');
assert.ok(mandatoryValidations.some((item) => item.Tipo === 'Divergência de lote como alerta não bloqueante'), 'lote diferente gera alerta não bloqueante');

console.log('association optimization tests passed');
